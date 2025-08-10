#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QFileDialog>
#include <QSettings>
#include <QMessageBox>
#include <QProcess>
#include <QDebug>
#include <QScrollBar>
#include <QStringList>
#include <QDragEnterEvent>
#include <QDropEvent>
#include <QCheckBox>
#include <QMouseEvent>
#include <QTableWidgetItem>
#include <QHeaderView>      // 表格头视图
#include <QMimeData>        // 存储拖拽数据
#include <QUrl>             // 用于处理拖拽的文件路径
#include <QDirIterator>     // 目录迭代器，用于递归扫描文件夹
#include <QElapsedTimer>    // 高精度计时器，用于实现防抖功能

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

/**
 * @brief 应用程序主窗口类。
 *
 * 管理UI交互、文件列表、模式切换及与后端进程的通信。
 */
class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

protected:
    // 事件处理重写
    void dragEnterEvent(QDragEnterEvent *event) override; // 处理文件拖拽进入
    void dropEvent(QDropEvent *event) override;           // 处理文件拖拽放下
    bool eventFilter(QObject *watched, QEvent *event) override; // 拦截并处理特定控件事件

private slots:
    // UI控件槽函数
    void on_addFilesButton_clicked();           // “添加文件”按钮点击
    void on_removeSelectedButton_clicked();     // “删除选中”按钮点击
    void on_selectAllButton_clicked();          // “全选/全不选”按钮点击
    void on_invertSelectionButton_clicked();    // “反选”按钮点击
    void on_browseInputDirButton_clicked();       // “选择输入文件夹”按钮点击
    void on_inputDirPathLineEdit_editingFinished(); // 输入文件夹路径编辑完成
    void on_refreshDirButton_clicked();           // “刷新”按钮点击
    void on_browseOutputButton_clicked();         // “选择输出文件夹”按钮点击
    void on_outputToSourceCheckBox_stateChanged(int state); // “输出到源路径”复选框状态改变
    void on_startProcessButton_clicked();         // “开始/取消处理”按钮点击

    // 逻辑与进程控制槽函数
    void handleModeChangeRequest();                  // 核心：处理模式切换请求（由事件过滤器调用）
    void onProcessStarted();                         // 后端进程已启动
    void onProcessFinished(int exitCode, QProcess::ExitStatus exitStatus); // 后端进程已结束
    void onProcessError(QProcess::ProcessError error); // 后端进程发生错误
    void readProcessOutput();                        // 读取后端进程的标准输出
    void readProcessError();                         // 读取后端进程的标准错误

private:
    Ui::MainWindow *ui;

    // 核心状态与数据
    QProcess *m_process;
    bool m_isProcessing;
    QStringList m_fileList;
    QString m_lastSelectedPath;
    bool m_dontAskOnModeChange;
    QElapsedTimer m_modeChangeDebounceTimer;

    // 私有辅助函数
    void loadSettings();                // 加载配置
    void saveSettings();                // 保存配置
    void addFileToList(const QString &filePath); // 添加文件到列表
    void updateUiForProcessingState(bool isProcessing); // 根据处理状态更新UI
    void updateActionButtonStates();   // 根据列表内容更新按钮状态
    void findAndUpdatetTableRow(const QString &filePath, const QString &status, const QString &message); // 更新表格行状态
    void scanAndPopulateFiles(const QString &dirPath); // 扫描并填充文件
    void updateUiForMode(bool isDirectoryMode);        // 根据模式更新UI

};
#endif // MAINWINDOW_H
