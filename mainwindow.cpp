#include "mainwindow.h"
#include "ui_mainwindow.h"

#include <QFileDialog>
#include <QDebug>
#include <QDateTime>
#include <QMessageBox>
#include <QFile>
#include <QCryptographicHash>

#include "xlsxdocument.h"


#define SETTINGS_LAST_PATH  "last_path"


MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    ui->scan_toolButton->setIcon(QIcon(QStringLiteral(":/img/img/start.svg")));
    ui->scan_toolButton->setToolTip(QStringLiteral("Сканировать"));
    ui->scan_toolButton->setStyleSheet(QStringLiteral("border: 0;"));

    ui->toTxt_toolButton->setIcon(QIcon(QStringLiteral(":/img/img/txt.svg")));
    ui->toTxt_toolButton->setToolTip(QStringLiteral("Экспорт в .txt"));
    ui->toTxt_toolButton->setStyleSheet(QStringLiteral("border: 0;"));

    ui->toXlsx_toolButton->setIcon(QIcon(QStringLiteral(":/img/img/xls.svg")));
    ui->toXlsx_toolButton->setToolTip(QStringLiteral("Экспорт в .xlsx"));
    ui->toXlsx_toolButton->setStyleSheet(QStringLiteral("border: 0;"));

    ui->tableWidget->horizontalHeader()->setSectionResizeMode(COL_NAME, QHeaderView::Stretch);
    ui->tableWidget->horizontalHeader()->setSectionResizeMode(COL_DATE_TIME, QHeaderView::Fixed);
    ui->tableWidget->setColumnWidth(COL_DATE_TIME, 150);
    ui->tableWidget->verticalHeader()->setVisible(false);
    ui->tableWidget->verticalHeader()->setSectionResizeMode(QHeaderView::ResizeToContents);

    connect(ui->browse_pushButton, &QPushButton::clicked, this, &MainWindow::slotBrowse);
    connect(ui->scan_toolButton, &QToolButton::clicked, this, &MainWindow::slotScan);
    connect(ui->toTxt_toolButton, &QToolButton::clicked, this, &MainWindow::slotWriteTxt);
    connect(ui->toXlsx_toolButton, &QToolButton::clicked, this, &MainWindow::slotWriteXlsx);
    connect(ui->path_lineEdit, &QLineEdit::textChanged, this, &MainWindow::slotPathChanged);

    auto lastPath = settings->value(SETTINGS_LAST_PATH, "").toString();
    if (lastPath.isEmpty())
    {
        lastPath = QDir::homePath();
    }
    ui->path_lineEdit->setText(lastPath);

    setTxtXlsxEnabled();
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::setTxtXlsxEnabled()
{
    ui->toTxt_toolButton->setEnabled(ui->tableWidget->rowCount() > 0);
    ui->toXlsx_toolButton->setEnabled(ui->tableWidget->rowCount() > 0);
}

QByteArray MainWindow::getMd5Checksumm(const QString &filename)
{
    auto result = QByteArray();
    QFile f(filename);
    if (!f.open(QFile::ReadOnly))
    {
        return result;
    }

    QCryptographicHash hash(QCryptographicHash::Md5);
    if (hash.addData(&f))
    {
        result = hash.result();
    }
    f.close();
    return result;
}

QString MainWindow::createSaveFilename(const QString &folderPath, EXPORT_MODES mode)
{
    int max = 0;
    QString extention = mode == TXT_EXPORT ? QStringLiteral(".txt")
                                           : QStringLiteral(".xlsx");
    QRegExp rex(QStringLiteral("Отчет\\s\\d+\\") + extention);
    const QFileInfoList fileList = QDir(folderPath).entryInfoList(QStringList(), QDir::Files);
    for (const auto &info : fileList)
    {
        if (info.isHidden())
        {
            continue;
        }
        if (!rex.exactMatch(info.fileName()))
        {
            continue;
        }
        int number = (info.fileName().remove(0, 6).remove(extention)).toInt();
        if (number > max)
        {
            max = number;
        }
    }
    return QStringLiteral("Отчет ") + QString::number(++max) + extention;
}

void MainWindow::slotBrowse()
{
    auto currentPath = ui->path_lineEdit->text();
    if (currentPath.isEmpty())
    {
        currentPath = QDir::homePath();
    }
    auto folderPath = QFileDialog::getExistingDirectory(this, QStringLiteral("Папка"), currentPath);
    if (!folderPath.isEmpty())
    {
        ui->path_lineEdit->setText(folderPath);
    }
}

void MainWindow::slotScan()
{
    auto folderPath = ui->path_lineEdit->text();
    if (folderPath.isEmpty())
    {
        return;
    }

    settings->setValue(SETTINGS_LAST_PATH, folderPath);
    setCursor(Qt::WaitCursor);
    ui->tableWidget->clearContents();
    ui->tableWidget->setRowCount(0);
    const QFileInfoList fileList = QDir(folderPath).entryInfoList(QStringList(), QDir::Files);
    for (const auto &info : fileList)
    {
        if (info.isHidden())
        {
            continue;
        }
        int row = ui->tableWidget->rowCount();
        ui->tableWidget->insertRow(row);
        {
            auto item = new QTableWidgetItem();
            item->setFlags(item->flags() &~Qt::ItemIsEditable);
            item->setText(info.fileName());

            QByteArray data = getMd5Checksumm(info.filePath()).toHex();
            item->setData(ROLE_MD5, QString::fromStdString(data.toStdString()));
            item->setData(ROLE_FILE_SIZE, info.size());

            ui->tableWidget->setItem(row, COL_NAME, item);
        }
        {
            QDateTime dateTime(info.lastModified().date(), info.lastModified().time());
            auto item = new QTableWidgetItem();
            item->setFlags(item->flags() &~Qt::ItemIsEditable);
            item->setTextAlignment(Qt::AlignCenter);
            item->setData(ROLE_DATE_TIME, dateTime);
            item->setText(dateTime.date().toString(QStringLiteral("dd.MM.yyyy"))
                          + QStringLiteral("  ") + dateTime.time().toString(QStringLiteral("hh:mm:ss")));
            ui->tableWidget->setItem(row, COL_DATE_TIME, item);
        }
    }
    setTxtXlsxEnabled();
    setCursor(Qt::ArrowCursor);
}

void MainWindow::slotWriteTxt()
{
    if (!ui->tableWidget->rowCount())
    {
        return;
    }

    QString saveName = ui->path_lineEdit->text() + QDir::separator() + createSaveFilename(ui->path_lineEdit->text(), TXT_EXPORT);
    const auto savePath = QFileDialog::getSaveFileName(this, QStringLiteral("Сохранение")
                                                       , saveName
                                                       , QStringLiteral("*.txt"));
    qDebug() << savePath;
    if (savePath.isEmpty())
    {
        return;
    }

    QFile file(savePath);
    if (!file.open(QFile::WriteOnly))
    {
        return;
    }

    QTextStream stream(&file);
    for (int i = 0; i < ui->tableWidget->rowCount(); ++i)
    {
        stream << ui->tableWidget->item(i, COL_DATE_TIME)->text() << QStringLiteral("\r\n");
    }
    file.close();
    QMessageBox::information(this, QStringLiteral("Успех")
                             , QStringLiteral("Запись в файл завершена успешно"));
}

void MainWindow::slotWriteXlsx()
{
    if (!ui->tableWidget->rowCount())
    {
        return;
    }

    QString saveName = ui->path_lineEdit->text() + QDir::separator() + createSaveFilename(ui->path_lineEdit->text(), XLSX_EXPORT);
    const auto savePath = QFileDialog::getSaveFileName(this, QStringLiteral("Сохранение")
                                                       , saveName
                                                       , QStringLiteral("*.xlsx"));
    if (savePath.isEmpty())
    {
        return;
    }

    QXlsx::Document xlsx;
    xlsx.setColumnWidth(1, 80.0);
    xlsx.setColumnWidth(2, 20.0);
    xlsx.setColumnWidth(3, 35.0);
    xlsx.setColumnWidth(4, 15.0);
    QXlsx::Format txtFormat;
    txtFormat.setNumberFormatIndex(49);
    QXlsx::Format dateTimeFormat;
    dateTimeFormat.setNumberFormat(QStringLiteral("dd.MM.yyyy hh:mm:ss"));
    for (int i = 0; i < ui->tableWidget->rowCount(); ++i)
    {
        const auto xlsxRow = i + 1;
        xlsx.write(QStringLiteral("A%1").arg(xlsxRow), ui->tableWidget->item(i, COL_NAME)->text(), txtFormat);
        xlsx.write(QStringLiteral("B%1").arg(xlsxRow), ui->tableWidget->item(i, COL_DATE_TIME)->data(ROLE_DATE_TIME).toDateTime(), dateTimeFormat);
        xlsx.write(QStringLiteral("C%1").arg(xlsxRow), ui->tableWidget->item(i, COL_NAME)->data(ROLE_MD5).toString(), txtFormat);
        xlsx.write(QStringLiteral("D%1").arg(xlsxRow), ui->tableWidget->item(i, COL_NAME)->data(ROLE_FILE_SIZE).value<qint64>(), txtFormat);
    }
    xlsx.saveAs(savePath);

    QMessageBox::information(this, QStringLiteral("Успех")
                             , QStringLiteral("Запись в файл завершена успешно"));
}

void MainWindow::slotPathChanged()
{
    ui->scan_toolButton->setEnabled(!ui->path_lineEdit->text().isEmpty());
}


