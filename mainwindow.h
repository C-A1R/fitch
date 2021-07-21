#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QSettings>

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

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
        ROLE_MD5,
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

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

private:
    void setTxtXlsxEnabled();
    QByteArray getMd5Checksumm(const QString &filename);
    QString createSaveFilename(const QString &folderPath, EXPORT_MODES exportMode);

private slots:
    void slotBrowse();
    void slotScan();
    void slotWriteTxt();
    void slotWriteXlsx();
    void slotPathChanged();
};
#endif // MAINWINDOW_H
