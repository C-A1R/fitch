#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include "ChecksumCalculator.h"

#include <QMainWindow>
#include <QSettings>

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class QDir;

class MainWindow : public QMainWindow
{
    Q_OBJECT

    enum COLUMNS
    {
        COL_NAME,
        COL_DATE_TIME,
        COL_MAX
    };

    enum ROLES
    {
        ROLE_DATE_TIME = Qt::UserRole,
        ROLE_CHECKSUM,
        ROLE_FILE_SIZE
    };

    enum EXPORT_MODES
    {
        TXT_EXPORT,
        XLSX_EXPORT
    };

    Ui::MainWindow *ui;
    const QString settingsFilename = QStringLiteral("settings.conf");
    QScopedPointer<QSettings> settings {new QSettings(QStringLiteral("settings.conf"), QSettings::IniFormat)};
    QSharedPointer<ChecksumCalculator> checksumCalculator;

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

private:
    void setTxtXlsxEnabled();
    QString createSavePath(const EXPORT_MODES mode);
    void showSuccessMessage(const QString &savePath);
    static QSharedPointer<ChecksumCalculator> makeChecksumCalculator(const ChecksumCalculator::CHECKSUM_TYPES type);

private slots:
    void slotBrowse();
    void slotScan();
    void slotWriteTxt();
    void slotWriteXlsx();
    void slotPathChanged();
    void slotReportFileWritten(const QString &savePath);

signals:
    void signalReportFileWritten(const QString &savePath);
};
#endif // MAINWINDOW_H
