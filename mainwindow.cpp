#include "mainwindow.h"
#include "ui_mainwindow.h"

#include "xlsxdocument.h"

#include <QFileDialog>
#include <QDateTime>
#include <QMessageBox>
#include <QFile>
#include <QTimer>
#include <QTextStream>
#include <QDesktopServices>

#define SETTINGS_LAST_PATH      "last_path"
#define SETTINGS_CHECKSUM_TYPE  "checksum_type"
#define SETTINGS_OPEN_REPORT    "open_report"

#define MAJOR_VERSION 1
#define MINOR_VERSION 2


MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    setWindowTitle(QStringLiteral("%1 %2.%3").arg(windowTitle()).arg(MAJOR_VERSION).arg(MINOR_VERSION));

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

    ui->checksum_comboBox->addItem("CRC32", static_cast<uint>(ChecksumCalculator::CHECKSUM_TYPES::CRC32));
    ui->checksum_comboBox->addItem("MD5", static_cast<uint>(ChecksumCalculator::CHECKSUM_TYPES::MD5));
    ui->checksum_comboBox->addItem("SHA-1", static_cast<uint>(ChecksumCalculator::CHECKSUM_TYPES::SHA_1));
    const uint last_checksum_type = settings->value(SETTINGS_CHECKSUM_TYPE, 0).toInt();
    if (last_checksum_type < static_cast<uint>(ChecksumCalculator::CHECKSUM_TYPES::MAX))
        ui->checksum_comboBox->setCurrentIndex(last_checksum_type);
    else
        ui->checksum_comboBox->setCurrentIndex(0);
    const bool open_report = settings->value(SETTINGS_OPEN_REPORT, 0).toBool();
    ui->open_checkBox->setCheckState(open_report ? Qt::Checked : Qt::Unchecked);

    connect(ui->browse_pushButton, &QPushButton::clicked, this, &MainWindow::slotBrowse);
    connect(ui->scan_toolButton, &QToolButton::clicked, this, &MainWindow::slotScan);
    connect(ui->toTxt_toolButton, &QToolButton::clicked, this, &MainWindow::slotWriteTxt);
    connect(ui->toXlsx_toolButton, &QToolButton::clicked, this, &MainWindow::slotWriteXlsx);
    connect(ui->path_lineEdit, &QLineEdit::textChanged, this, &MainWindow::slotPathChanged);
    connect(this, &MainWindow::signalReportFileWritten, this, &MainWindow::slotReportFileWritten);

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
    settings->setValue(SETTINGS_OPEN_REPORT, ui->open_checkBox->checkState() == Qt::Checked ? 1 : 0);
    delete ui;
}

void MainWindow::setTxtXlsxEnabled()
{
    ui->toTxt_toolButton->setEnabled(ui->tableWidget->rowCount() > 0);
    ui->toXlsx_toolButton->setEnabled(ui->tableWidget->rowCount() > 0);
}

QString MainWindow::createSavePath(const EXPORT_MODES mode)
{
    const QString &folderPath = ui->path_lineEdit->text();
    if (folderPath.isEmpty())
    {
        return QString();
    }

    int max = 0;
    const QString extention = mode == TXT_EXPORT ? QStringLiteral(".txt")
                                                 : QStringLiteral(".xlsx");
    const QRegExp rex(QStringLiteral("Отчет\\s\\d+\\") + extention);
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
        const int number = (info.fileName().remove(0, 6).remove(extention)).toInt();
        if (number > max)
        {
            max = number;
        }
    }
    return folderPath + '/' + QStringLiteral("Отчет ") + QString::number(++max) + extention;
}

void MainWindow::showSuccessMessage(const QString &savePath)
{
    QMessageBox msgBox(this);
    msgBox.setIcon(QMessageBox::Information);
    msgBox.setWindowTitle(QStringLiteral("Экспорт"));
    msgBox.setText(QStringLiteral("Сохранено в %1").arg(savePath));
    msgBox.setStandardButtons(QMessageBox::NoButton);
    QTimer::singleShot(1000, &msgBox, &QMessageBox::accept);
    msgBox.exec();
}

QSharedPointer<ChecksumCalculator> MainWindow::makeChecksumCalculator(const ChecksumCalculator::CHECKSUM_TYPES type)
{
    if (type == ChecksumCalculator::CHECKSUM_TYPES::CRC32)
        return QSharedPointer<CRC32_ChecksumCalculator>(new CRC32_ChecksumCalculator);
    if (type == ChecksumCalculator::CHECKSUM_TYPES::MD5)
        return QSharedPointer<MD5_ChecksumCalculator>(new MD5_ChecksumCalculator);
    if (type == ChecksumCalculator::CHECKSUM_TYPES::SHA_1)
        return QSharedPointer<SHA1_ChecksumCalculator>(new SHA1_ChecksumCalculator);
    return nullptr;
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
    const QString &folderPath = ui->path_lineEdit->text();
    if (folderPath.isEmpty())
    {
        return;
    }

    settings->setValue(SETTINGS_LAST_PATH, folderPath);
    setCursor(Qt::WaitCursor);
    ui->tableWidget->clearContents();
    ui->tableWidget->setRowCount(0);

    const int checksum_type = ui->checksum_comboBox->currentData().toInt();
    settings->setValue(SETTINGS_CHECKSUM_TYPE, checksum_type);
    checksumCalculator = makeChecksumCalculator(static_cast<ChecksumCalculator::CHECKSUM_TYPES>(checksum_type));
    if (!checksumCalculator)
    {
        QMessageBox::critical(this, QStringLiteral("Ошибка"), QStringLiteral("Проблема с расчетом чексуммы"));
        return;
    }

    const QFileInfoList fileList = QDir(folderPath).entryInfoList(QStringList(), QDir::Files);
    for (const auto &info : fileList)
    {
        if (info.isHidden())
        {
            continue;
        }
        const int row = ui->tableWidget->rowCount();
        ui->tableWidget->insertRow(row);
        {
            auto item = new QTableWidgetItem();
            item->setFlags(item->flags() &~Qt::ItemIsEditable);
            item->setText(info.fileName());
            item->setData(ROLE_CHECKSUM, checksumCalculator->calcChecksum(info.filePath()));
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

    const QString savePath = createSavePath(TXT_EXPORT);
    if (savePath.isEmpty())
    {
        return;
    }

    QFile file(savePath);
    if (!file.open(QFile::WriteOnly))
    {
        return;
    }

    const std::size_t wName = 50;
    const std::size_t wDate = 25;
    const std::size_t wChecksum = checksumCalculator->maxLen();
    const std::size_t wSize = 15;
    QTextStream stream(&file);
    stream << Qt::left  << qSetFieldWidth(wName)     << QStringLiteral("Filename")
                        << qSetFieldWidth(wDate)     << QStringLiteral("Last edit date time")
                        << qSetFieldWidth(wChecksum) << QStringLiteral("Checksum (%1)").arg(checksumCalculator->name())
           << Qt::right << qSetFieldWidth(wSize)     << QStringLiteral("File size")
                        << qSetFieldWidth(0)         << Qt::endl;
    for (int i = 0; i < ui->tableWidget->rowCount(); ++i)
    {
        const QString filename = QString::fromLocal8Bit(ui->tableWidget->item(i, COL_NAME)->data(Qt::DisplayRole).toByteArray());
        stream << Qt::left  << qSetFieldWidth(wName)     << filename
                            << qSetFieldWidth(wDate)     << ui->tableWidget->item(i, COL_DATE_TIME)->text()
                            << qSetFieldWidth(wChecksum) << ui->tableWidget->item(i, COL_NAME)->data(ROLE_CHECKSUM).toString()
               << Qt::right << qSetFieldWidth(wSize)     << ui->tableWidget->item(i, COL_NAME)->data(ROLE_FILE_SIZE).value<qint64>()
                            << qSetFieldWidth(0)         << Qt::endl;
    }
    file.close();

    emit signalReportFileWritten(savePath);
}

void MainWindow::slotWriteXlsx()
{
    if (!ui->tableWidget->rowCount())
    {
        return;
    }

    const QString savePath = createSavePath(XLSX_EXPORT);
    if (savePath.isEmpty())
    {
        return;
    }

    QXlsx::Document xlsx;
    xlsx.setColumnWidth(1, 80.0);
    xlsx.setColumnWidth(2, 20.0);
    xlsx.setColumnWidth(3, 40.0);
    xlsx.setColumnWidth(4, 15.0);

    {
        QXlsx::Format headerFormat;
        headerFormat.setNumberFormatIndex(49);
        headerFormat.setFontBold(true);
        headerFormat.setHorizontalAlignment(QXlsx::Format::HorizontalAlignment::AlignHCenter);
        xlsx.write(QStringLiteral("A1"), QStringLiteral("Filename"), headerFormat);
        xlsx.write(QStringLiteral("B1"), QStringLiteral("Last edit date time"), headerFormat);
        xlsx.write(QStringLiteral("C1"), QStringLiteral("Checksum (%1)").arg(checksumCalculator->name()), headerFormat);
        xlsx.write(QStringLiteral("D1"), QStringLiteral("File size"), headerFormat);
    }

    QXlsx::Format txtFormat;
    txtFormat.setNumberFormatIndex(49);
    QXlsx::Format dateTimeFormat;
    dateTimeFormat.setNumberFormat(QStringLiteral("dd.MM.yyyy hh:mm:ss"));
    for (int i = 0; i < ui->tableWidget->rowCount(); ++i)
    {
        const auto row = i + 2;
        xlsx.write(QStringLiteral("A%1").arg(row), ui->tableWidget->item(i, COL_NAME)->text(), txtFormat);
        xlsx.write(QStringLiteral("B%1").arg(row), ui->tableWidget->item(i, COL_DATE_TIME)->data(ROLE_DATE_TIME).toDateTime(), dateTimeFormat);
        xlsx.write(QStringLiteral("C%1").arg(row), ui->tableWidget->item(i, COL_NAME)->data(ROLE_CHECKSUM).toString(), txtFormat);
        xlsx.write(QStringLiteral("D%1").arg(row), ui->tableWidget->item(i, COL_NAME)->data(ROLE_FILE_SIZE).value<qint64>(), txtFormat);
    }

    if (!xlsx.saveAs(savePath))
    {
        QMessageBox::critical(this, QStringLiteral("Ошибка"), QStringLiteral("Не удалось сохранить в %1").arg(savePath));
        return;
    }

    emit signalReportFileWritten(savePath);
}

void MainWindow::slotPathChanged()
{
    ui->scan_toolButton->setEnabled(!ui->path_lineEdit->text().isEmpty());
}

void MainWindow::slotReportFileWritten(const QString &savePath)
{
    showSuccessMessage(savePath);
    if (ui->open_checkBox->checkState() == Qt::Checked)
        QDesktopServices::openUrl(QUrl::fromLocalFile(savePath));
}


