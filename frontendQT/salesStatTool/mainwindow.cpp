#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    loadSettings(); // 启动时加载设置
}

MainWindow::~MainWindow()
{
    saveSettings(); // 关闭时保存设置
    delete ui;
}

void MainWindow::loadSettings()
{
    // 获取可执行文件所在的目录路径
    QString configFilePath = QApplication::applicationDirPath() + "/config.ini";
    // 使用完整文件路径和ini格式来创建QSettings对象
    QSettings settings(configFilePath, QSettings::IniFormat);

    // 读取路径，如果键不存在，则返回一个空字符串作为默认值
    QString inputPath = settings.value("paths/inputPath", "").toString();
    QString outputPath = settings.value("paths/outputPath", "").toString();
    ui->inputPathLineEdit->setText(inputPath);
    ui->outputPathLineEdit->setText(outputPath);

    // 读取冲突处理选项的索引，如果不存在，则返回0作为默认值
    int conflictIndex = settings.value("options/conflictIndex", 2).toInt();
    ui->conflictComboBox->setCurrentIndex(conflictIndex);
}

void MainWindow::saveSettings()
{
    // 获取可执行文件所在的目录路径
    QString configFilePath = QApplication::applicationDirPath() + "/config.ini";
    // 使用完整文件路径和ini格式来创建QSettings对象
    QSettings settings(configFilePath, QSettings::IniFormat);

    // 保存当前文本框中的路径
    settings.setValue("paths/inputPath", ui->inputPathLineEdit->text());
    settings.setValue("paths/outputPath", ui->outputPathLineEdit->text());

    // 保存下拉框当前选中的索引
    settings.setValue("options/conflictIndex", ui->conflictComboBox->currentIndex());
}

void MainWindow::on_browseInputButton_clicked()
{
    // 以当前文本框中的路径作为默认打开路径，如果为空则使用用户主目录
    QString currentPath = ui->inputPathLineEdit->text();
    QString dir = QFileDialog::getExistingDirectory(this, "选择输入文件夹", currentPath.isEmpty() ? QDir::homePath() : currentPath);

    // 只有当用户确实选择了文件夹（而不是点击取消），才更新文本框
    if (!dir.isEmpty()) {
        ui->inputPathLineEdit->setText(dir);
    }
}

void MainWindow::on_browseOutputButton_clicked()
{
    QString currentPath = ui->outputPathLineEdit->text();
    QString dir = QFileDialog::getExistingDirectory(this, "选择输出文件夹", currentPath.isEmpty() ? QDir::homePath() : currentPath);

    if (!dir.isEmpty()) {
        ui->outputPathLineEdit->setText(dir);
    }
}

void MainWindow::on_processButton_clicked()
{
    // 1. 获取并验证用户输入
    QString inputPath = ui->inputPathLineEdit->text();
    QString outputPath = ui->outputPathLineEdit->text();

    if (inputPath.isEmpty() || outputPath.isEmpty()) {
        QMessageBox::warning(this, "输入错误", "输入和输出文件夹路径不能为空！");
        return;
    }

    // 2. 准备命令行参数
    QString conflictPolicy;
    switch (ui->conflictComboBox->currentIndex()) {
        case 0: conflictPolicy = "rename"; break;
        case 1: conflictPolicy = "overwrite"; break;
        case 2: conflictPolicy = "skip"; break;
        default: conflictPolicy = "skip";
    }

    // 3. 禁用按钮、清空log、显示状态
    ui->processButton->setEnabled(false);
    ui->plainTextEdit->clear(); // 在开始处理前清空日志区
    ui->plainTextEdit->appendPlainText("--- 开始处理 ---");
    statusBar()->showMessage("正在处理中，请稍候...");

    // 4. 配置和启动外部进程
    QProcess *process = new QProcess(this);

    // 默认后端引擎在可执行文件同级的 "backend_engine" 文件夹内
    QString programPath = QDir::toNativeSeparators(QApplication::applicationDirPath() + "/backend_engine/backend_engine.exe");

    QStringList arguments;
    arguments << inputPath << outputPath << "--on-conflict" << conflictPolicy;

    // 5. 连接信号和槽
    connect(process, static_cast<void (QProcess::*)(int, QProcess::ExitStatus)>(&QProcess::finished), this, [=](int exitCode, QProcess::ExitStatus exitStatus){
        // 恢复按钮
        ui->processButton->setEnabled(true);

        if (exitStatus == QProcess::NormalExit && exitCode == 0) {
            statusBar()->showMessage("处理成功！", 5000);
            ui->plainTextEdit->appendPlainText("\n--- 处理成功！ ---");
            //QMessageBox::information(this, "完成", "所有文件处理完毕！");
        } else {
            QString errorOutput = process->readAllStandardError();
            statusBar()->showMessage("处理失败！详情见弹窗。", 5000);
            ui->plainTextEdit->appendPlainText("\n--- 处理失败！ ---");
            ui->plainTextEdit->appendPlainText("错误码: " + QString::number(exitCode));
            ui->plainTextEdit->appendPlainText("详细信息:\n" + errorOutput);
            QMessageBox::critical(this, "处理失败", "后端引擎执行出错。\n\n错误码: " + QString::number(exitCode) + "\n\n详细信息:\n" + errorOutput);
        }

        // 在进程结束后，将滚动条滚动到底部
        QScrollBar *scrollbar = ui->plainTextEdit->verticalScrollBar();
        scrollbar->setValue(scrollbar->maximum());

        process->deleteLater();
    });

    // 连接错误处理信号
    connect(process, &QProcess::errorOccurred, this, [=](QProcess::ProcessError error){
        if (process->state() == QProcess::NotRunning) {
            ui->processButton->setEnabled(true);
            statusBar()->showMessage("启动失败！", 5000);
            ui->plainTextEdit->appendPlainText("\n--- 启动失败！ ---");
            ui->plainTextEdit->appendPlainText("无法启动后端引擎。请检查程序目录下 'backend_engine' 文件夹及其中的 'backend_engine.exe' 是否存在。");
            QMessageBox::critical(this, "启动失败", "无法启动后端引擎。\n请检查程序目录下 'backend_engine' 文件夹及其中的 'backend_engine.exe' 是否存在。\n\n错误类型: " + QString::number(error));
            process->deleteLater();
        }
    });

    // 连接标准输出信号，实时显示日志
    connect(process, &QProcess::readyReadStandardOutput, this, [=](){
        // 将读取到的字节流用本地编码转为QString
        const QByteArray data = process->readAllStandardOutput();
        ui->plainTextEdit->insertPlainText(QString::fromLocal8Bit(data));
    });

    // 6. 启动
    process->start(programPath, arguments);
}
