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

    enum ROlES
    {
        ROLE_DATE_TIME = Qt::UserRole,
        ROLE_MD5,
        ROLE_FILE_SIZE
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

private slots:
    void slotBrowse();
    void slotScan();
    void slotWriteTxt();
    void slotWriteXlsx();
    void slotPathChanged();
};
#endif // MAINWINDOW_H
