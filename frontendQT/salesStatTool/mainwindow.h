#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>

#include <QFileDialog> // 文件/文件夹选择对话框
#include <QSettings> // 保存和读取设置
#include <QMessageBox> // 显示警告、信息等对话框
#include <QProcess> // 执行外部程序
#include <QDebug> // 在控制台打印调试信息
#include <QScrollBar> // 控制滚动条

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

private slots:
    void on_browseInputButton_clicked();
    void on_browseOutputButton_clicked();
    void on_processButton_clicked();

private:
    Ui::MainWindow *ui;

    // 用于加载和保存设置
    void loadSettings();
    void saveSettings();

};
#endif // MAINWINDOW_H
