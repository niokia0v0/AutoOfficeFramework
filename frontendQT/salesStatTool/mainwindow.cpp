#include "mainwindow.h"
#include "ui_mainwindow.h"

/**
 * @brief MainWindow构造函数
 *
 * 负责初始化UI、连接信号与槽、安装事件过滤器、加载用户设置，
 * 并根据加载的设置完成窗口的初始状态配置。
 */
MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
    , m_process(nullptr)
    , m_isProcessing(false)
    , m_dontAskOnModeChange(false)
{
    ui->setupUi(this);

    // 1. 通用UI初始化
    setAcceptDrops(true); // 使主窗口能够接收拖拽事件

    // 2. 文件列表表格(QTableWidget)配置
    ui->fileListTableWidget->setColumnCount(4); // 设置4列
    // 设置列宽调整策略，以优化显示效果
    ui->fileListTableWidget->horizontalHeader()->setSectionResizeMode(0, QHeaderView::ResizeToContents); // 复选框列，根据内容自适应
    ui->fileListTableWidget->horizontalHeader()->setSectionResizeMode(1, QHeaderView::ResizeToContents); // 状态列，根据内容自适应
    ui->fileListTableWidget->horizontalHeader()->setSectionResizeMode(2, QHeaderView::Interactive);    // 文件名列，用户可手动调整
    ui->fileListTableWidget->horizontalHeader()->setSectionResizeMode(3, QHeaderView::Stretch);         // 路径列，自动拉伸以填满剩余空间
    ui->fileListTableWidget->setColumnWidth(2, 250); // 为文件名设置初始宽度

    // 3. 信号与槽连接
    // 当用户在输入文件夹路径框中完成编辑（按回车或焦点离开）时，触发扫描
    connect(ui->inputDirPathLineEdit, &QLineEdit::editingFinished, this, &MainWindow::on_inputDirPathLineEdit_editingFinished);

    // 4. 事件过滤器安装
    // 这是解决模式切换复选框快速点击问题的核心。
    // 将MainWindow自身作为事件过滤器安装到复选框上，
    // 从而可以在eventFilter函数中拦截并完全控制其鼠标事件。
    ui->useDirectoryModeCheckBox->installEventFilter(this);

    // 5. 加载设置并完成最终初始化
    loadSettings(); // 从配置文件加载上次的状态
    updateUiForMode(ui->useDirectoryModeCheckBox->isChecked()); // 根据加载的模式，更新UI控件的初始状态
    // 如果启动时就是文件夹模式，则自动执行一次文件扫描
    if (ui->useDirectoryModeCheckBox->isChecked()) {
        scanAndPopulateFiles(ui->inputDirPathLineEdit->text());
    }
    updateActionButtonStates(); // 根据列表是否为空，更新按钮的初始可用状态
    on_outputToSourceCheckBox_stateChanged(ui->outputToSourceCheckBox->checkState()); // 更新输出路径控件的初始状态

    // 启动防抖计时器
    m_modeChangeDebounceTimer.start();
}

/**
 * @brief MainWindow析构函数
 *
 * 负责在程序关闭前进行资源清理，如强制终止仍在运行的后端进程，
 * 以及调用saveSettings()将当前配置写入文件。
 */
MainWindow::~MainWindow()
{
    if (m_process && m_process->state() != QProcess::NotRunning) {
        m_process->kill();
    }
    saveSettings();
    delete ui;
}

/**
 * @brief 事件过滤器
 *
 * 这是整个模式切换交互逻辑中最关键的部分。
 * 它通过拦截鼠标的按下和释放事件，完全阻止了QCheckBox的默认行为，
 * 从而避免了因快速点击导致的UI视觉状态与程序逻辑状态不一致的问题。
 * 所有的逻辑处理都转交给了handleModeChangeRequest函数。
 */
bool MainWindow::eventFilter(QObject *watched, QEvent *event)
{
    // 检查事件是否来自文件夹输入模式(文件夹模式)的复选框
    if (watched == ui->useDirectoryModeCheckBox) {

        // 拦截鼠标左键按下的事件
        if (event->type() == QEvent::MouseButtonPress) {
            QMouseEvent *mouseEvent = static_cast<QMouseEvent*>(event);
            if (mouseEvent->button() == Qt::LeftButton) {
                // 调用自定义的逻辑处理函数
                handleModeChangeRequest();
                // 返回true，表示事件已被处理，应停止传播，QCheckBox自身不会再响应它
                return true;
            }
        }

        // 拦截鼠标左键释放的事件
        // 防止在鼠标松开时，QCheckBox出现视觉上的“闪烁”
        if (event->type() == QEvent::MouseButtonRelease) {
            QMouseEvent *mouseEvent = static_cast<QMouseEvent*>(event);
            if (mouseEvent->button() == Qt::LeftButton) {
                // 直接“吞掉”这个事件，不让QCheckBox做任何事
                return true;
            }
        }
    }

    // 对于所有其他的对象和事件，执行默认的父类处理
    return QMainWindow::eventFilter(watched, event);
}

/**
 * @brief 拖拽进入事件处理
 *
 * 当用户拖拽文件进入窗口区域时触发。
 * 如果当前处于文件夹模式或正在处理中，则忽略拖拽。
 * 否则，接受拖拽，并显示可放置的光标。
 */
void MainWindow::dragEnterEvent(QDragEnterEvent *event)
{
    if (ui->useDirectoryModeCheckBox->isChecked() || m_isProcessing) {
        event->ignore();
        return;
    }
    if (event->mimeData()->hasUrls()) {
        event->acceptProposedAction();
    }
}

/**
 * @brief 拖拽放下事件处理
 *
 * 当用户在窗口中释放拖拽的文件时触发。
 * 遍历所有被拖入的项，如果是文件夹则递归扫描添加，如果是文件则直接添加。
 */
void MainWindow::dropEvent(QDropEvent *event)
{
    // 若启用指定输入文件夹的模式则禁用拖拽导入
    if (ui->useDirectoryModeCheckBox->isChecked() || m_isProcessing) {
        event->ignore();
        return;
    }
    for (const QUrl &url : event->mimeData()->urls()) {
        if (url.isLocalFile()) {
            QString path = url.toLocalFile();
            QFileInfo fileInfo(path);
            if (fileInfo.isDir()) {
                QDirIterator it(path, {"*.csv", "*.xlsx"}, QDir::Files, QDirIterator::Subdirectories);
                while (it.hasNext()) {
                    addFileToList(it.next());
                }
            } else if (fileInfo.isFile()) {
                if (fileInfo.suffix().toLower() == "csv" || fileInfo.suffix().toLower() == "xlsx") {
                    addFileToList(path);
                }
            }
        }
    }
}

/**
 * @brief 添加单个文件到列表
 *
 * 这是所有文件添加操作的统一入口。
 * 它负责检查重复、更新内部数据列表、在UI表格中创建新行并填充所有单元格。
 * 对于复选框单元格，创建了一个容器QWidget来实现居中显示。
 */
void MainWindow::addFileToList(const QString &filePath)
{
    if (m_fileList.contains(filePath)) {
        return;
    }
    m_fileList.append(filePath);
    int currentRow = ui->fileListTableWidget->rowCount();
    ui->fileListTableWidget->insertRow(currentRow);

    // 创建并居中放置复选框
    QCheckBox *checkBox = new QCheckBox();
    checkBox->setChecked(true);
    QWidget *widgetContainer = new QWidget();
    QHBoxLayout *layout = new QHBoxLayout(widgetContainer);
    layout->addWidget(checkBox);
    layout->setAlignment(Qt::AlignCenter);
    layout->setContentsMargins(0, 0, 0, 0);
    widgetContainer->setLayout(layout);
    ui->fileListTableWidget->setCellWidget(currentRow, 0, widgetContainer);

    // 填充其他单元格
    ui->fileListTableWidget->setItem(currentRow, 1, new QTableWidgetItem("待处理"));
    ui->fileListTableWidget->setItem(currentRow, 2, new QTableWidgetItem(QFileInfo(filePath).fileName()));
    ui->fileListTableWidget->setItem(currentRow, 3, new QTableWidgetItem(filePath));

    updateActionButtonStates();
}

// 按钮槽函数实现

void MainWindow::on_addFilesButton_clicked()
{
    QStringList files = QFileDialog::getOpenFileNames(this, "选择数据文件", m_lastSelectedPath, "数据文件 (*.xlsx *.csv)");
    if (!files.isEmpty()) {
        m_lastSelectedPath = QFileInfo(files.first()).absolutePath(); // 记住本次选择的路径
        for (const QString &file : files) {
            addFileToList(file);
        }
    }
}

void MainWindow::on_removeSelectedButton_clicked()
{
    // 从后向前遍历以避免删除时索引错乱
    for (int i = ui->fileListTableWidget->rowCount() - 1; i >= 0; --i) {
        QWidget *widgetContainer = ui->fileListTableWidget->cellWidget(i, 0);
        if (widgetContainer) {
            QCheckBox *checkBox = widgetContainer->findChild<QCheckBox *>(); // 从容器中查找复选框
            if (checkBox && checkBox->isChecked()) {
                QString filePath = ui->fileListTableWidget->item(i, 3)->text();
                m_fileList.removeAll(filePath);
                ui->fileListTableWidget->removeRow(i);
            }
        }
    }
    updateActionButtonStates();
}

void MainWindow::on_selectAllButton_clicked()
{
    bool shouldSelectAll = false;
    // 检查是否已全部选中，如果不是，则目标是全选
    for (int i = 0; i < ui->fileListTableWidget->rowCount(); ++i) {
        QWidget *widgetContainer = ui->fileListTableWidget->cellWidget(i, 0);
        if (widgetContainer) {
            QCheckBox *checkBox = widgetContainer->findChild<QCheckBox *>();
            if (checkBox && !checkBox->isChecked()) {
                shouldSelectAll = true;
                break;
            }
        }
    }
    // 应用最终状态
    for (int i = 0; i < ui->fileListTableWidget->rowCount(); ++i) {
        QWidget *widgetContainer = ui->fileListTableWidget->cellWidget(i, 0);
        if (widgetContainer) {
            QCheckBox *checkBox = widgetContainer->findChild<QCheckBox *>();
            if (checkBox) {
                checkBox->setChecked(shouldSelectAll);
            }
        }
    }
}

void MainWindow::on_invertSelectionButton_clicked()
{
    for (int i = 0; i < ui->fileListTableWidget->rowCount(); ++i) {
        QWidget *widgetContainer = ui->fileListTableWidget->cellWidget(i, 0);
        if(widgetContainer) {
            QCheckBox *checkBox = widgetContainer->findChild<QCheckBox *>();
            if (checkBox) {
                checkBox->setChecked(!checkBox->isChecked());
            }
        }
    }
}

/**
 * @brief 模式切换的核心逻辑处理函数
 *
 * 由事件过滤器在捕获到用户点击时调用。
 * 1. 使用QElapsedTimer实现防抖，防止快速点击导致的逻辑混乱。
 * 2. 根据列表是否为空，决定是否需要弹窗确认。
 * 3. 如果用户取消，则无任何操作。
 * 4. 如果用户确认，则执行清空列表、切换UI模式、扫描文件等一系列操作。
 * 5. 所有对指定的复选框状态的改变都通过在此函数中调用setChecked()完成。
 */
void MainWindow::handleModeChangeRequest()
{
    // 核心防抖逻辑
    const qint64 DEBOUNCE_INTERVAL_MS = 300; // 设置防抖间隔
    if (m_modeChangeDebounceTimer.elapsed() < DEBOUNCE_INTERVAL_MS) {
        return; // 如果点击过于频繁，则直接忽略
    }
    m_modeChangeDebounceTimer.restart(); // 重置计时器，记录本次有效操作的时间

    // 获取当前状态和用户的意图状态
    bool currentState = ui->useDirectoryModeCheckBox->isChecked();
    bool intendedState = !currentState;

    // 如果列表为空，直接切换，无需确认
    if (ui->fileListTableWidget->rowCount() == 0) {
        ui->useDirectoryModeCheckBox->setChecked(intendedState);
        updateUiForMode(intendedState);
        if (intendedState) {
            scanAndPopulateFiles(ui->inputDirPathLineEdit->text());
        }
        updateActionButtonStates();
        return;
    }

    // 列表不为空，弹窗确认
    if (!m_dontAskOnModeChange) {
        QMessageBox msgBox(this);
        msgBox.setWindowTitle("确认操作");
        msgBox.setText("切换模式将清空当前文件列表，是否继续？");
        msgBox.setIcon(QMessageBox::Question);
        msgBox.addButton("是", QMessageBox::YesRole);
        msgBox.addButton("否", QMessageBox::NoRole);
        QCheckBox *dontAskAgain = new QCheckBox("不再提示");
        msgBox.setCheckBox(dontAskAgain);

        if (msgBox.exec() == 1) { // 用户点击“否”
            return; // 直接返回，不执行任何操作
        }
        if (dontAskAgain->isChecked()) {
            m_dontAskOnModeChange = true;
        }
    }

    // 用户同意切换，执行实际的切换逻辑
    ui->useDirectoryModeCheckBox->setChecked(intendedState);
    m_fileList.clear();
    ui->fileListTableWidget->setRowCount(0);
    updateUiForMode(intendedState);
    if (intendedState) {
        scanAndPopulateFiles(ui->inputDirPathLineEdit->text());
    }
    updateActionButtonStates();
}

void MainWindow::on_browseInputDirButton_clicked()
{
    QString startDir = ui->inputDirPathLineEdit->text().isEmpty() ? m_lastSelectedPath : ui->inputDirPathLineEdit->text();
    QString dir = QFileDialog::getExistingDirectory(this, "选择输入文件夹", startDir);
    if (!dir.isEmpty()) {
        ui->inputDirPathLineEdit->setText(dir);
        m_lastSelectedPath = dir;
        scanAndPopulateFiles(dir); // 选择后立即扫描
    }
}

void MainWindow::on_inputDirPathLineEdit_editingFinished()
{
    scanAndPopulateFiles(ui->inputDirPathLineEdit->text());
}

void MainWindow::on_refreshDirButton_clicked()
{
    scanAndPopulateFiles(ui->inputDirPathLineEdit->text());
}

// 设置的加载与保存

void MainWindow::loadSettings()
{
    QString configFilePath = QApplication::applicationDirPath() + "/config.ini";
    QSettings settings(configFilePath, QSettings::IniFormat);
    ui->outputPathLineEdit->setText(settings.value("paths/outputPath", "").toString());
    ui->inputDirPathLineEdit->setText(settings.value("paths/inputPath", "").toString());
    m_lastSelectedPath = settings.value("paths/lastSelectedPath", QDir::homePath()).toString();
    ui->conflictComboBox->setCurrentIndex(settings.value("options/conflictIndex", 2).toInt());
    ui->outputToSourceCheckBox->setChecked(settings.value("options/outputToSource", false).toBool());
    ui->useDirectoryModeCheckBox->setChecked(settings.value("options/useDirectoryMode", false).toBool());
    m_dontAskOnModeChange = settings.value("options/dontAskOnModeChange", false).toBool();
}

void MainWindow::saveSettings()
{
    QString configFilePath = QApplication::applicationDirPath() + "/config.ini";
    QSettings settings(configFilePath, QSettings::IniFormat);
    settings.setValue("paths/outputPath", ui->outputPathLineEdit->text());
    settings.setValue("paths/inputPath", ui->inputDirPathLineEdit->text());
    settings.setValue("paths/lastSelectedPath", m_lastSelectedPath);
    settings.setValue("options/conflictIndex", ui->conflictComboBox->currentIndex());
    settings.setValue("options/outputToSource", ui->outputToSourceCheckBox->isChecked());
    settings.setValue("options/useDirectoryMode", ui->useDirectoryModeCheckBox->isChecked());
    settings.setValue("options/dontAskOnModeChange", m_dontAskOnModeChange);
}

void MainWindow::on_browseOutputButton_clicked()
{
    QString startDir = ui->outputPathLineEdit->text().isEmpty() ? m_lastSelectedPath : ui->outputPathLineEdit->text();
    QString dir = QFileDialog::getExistingDirectory(this, "选择输出文件夹", startDir);
    if (!dir.isEmpty()) {
        ui->outputPathLineEdit->setText(dir);
        m_lastSelectedPath = dir;
    }
}

void MainWindow::on_outputToSourceCheckBox_stateChanged(int state)
{
    bool enabled = (state == Qt::Unchecked);
    ui->outputPathLineEdit->setEnabled(enabled);
    ui->browseOutputButton->setEnabled(enabled);
}

// 核心处理流程

void MainWindow::on_startProcessButton_clicked()
{
    // 如果正在处理，则按钮功能为“取消”
    if (m_isProcessing) {
        if (m_process) {
            m_process->kill();
        }
        return;
    }

    // 收集所有已勾选的任务
    QStringList tasks;
    for (int i = 0; i < ui->fileListTableWidget->rowCount(); ++i) {
        QWidget *widgetContainer = ui->fileListTableWidget->cellWidget(i, 0);
        if (widgetContainer) {
            QCheckBox *checkBox = widgetContainer->findChild<QCheckBox *>();
            if (checkBox && checkBox->isChecked()) {
                tasks.append(ui->fileListTableWidget->item(i, 3)->text());
            }
        }
    }

    // 启动前检查
    if (tasks.isEmpty()) {
        QMessageBox::warning(this, "没有任务", "请至少勾选一个要处理的文件！");
        return;
    }
    if (!ui->outputToSourceCheckBox->isChecked() && ui->outputPathLineEdit->text().isEmpty()) {
        QMessageBox::warning(this, "输出未指定", "请指定一个输出文件夹，或勾选“输出到源文件路径”！");
        return;
    }

    // 准备启动进程
    updateUiForProcessingState(true); // 锁定UI
    ui->plainTextEdit->clear();
    ui->plainTextEdit->appendPlainText("--- 开始处理 ---");

    m_process = new QProcess(this);
    // 设置Python脚本的环境变量，强制使用UTF-8，防止中文路径乱码
    QProcessEnvironment env = QProcessEnvironment::systemEnvironment();
    env.insert("PYTHONIOENCODING", "utf-8");
    m_process->setProcessEnvironment(env);

    // 连接所有进程信号
    connect(m_process, &QProcess::started, this, &MainWindow::onProcessStarted);
    connect(m_process, static_cast<void (QProcess::*)(int, QProcess::ExitStatus)>(&QProcess::finished), this, &MainWindow::onProcessFinished);
    connect(m_process, &QProcess::errorOccurred, this, &MainWindow::onProcessError);
    connect(m_process, &QProcess::readyReadStandardOutput, this, &MainWindow::readProcessOutput);
    connect(m_process, &QProcess::readyReadStandardError, this, &MainWindow::readProcessError);

    // 准备程序路径和命令行参数
    QString programPath = QDir::toNativeSeparators(QApplication::applicationDirPath() + "/backend_engine/backend_engine.exe");
    QStringList arguments;
    QString conflictPolicy;
    switch (ui->conflictComboBox->currentIndex()) {
        case 0: conflictPolicy = "rename"; break;
        case 1: conflictPolicy = "overwrite"; break;
        case 2: conflictPolicy = "skip"; break;
        default: conflictPolicy = "skip";
    }
    arguments << "--on-conflict" << conflictPolicy;
    if (!ui->outputToSourceCheckBox->isChecked()) {
        arguments << "--output-dir" << ui->outputPathLineEdit->text();
    }

    // 将任务列表暂存到进程对象属性中，待进程启动后写入
    m_process->setProperty("tasks", tasks);
    m_process->start(programPath, arguments);
}

// 进程信号处理

void MainWindow::onProcessStarted()
{
    QStringList tasks = m_process->property("tasks").toStringList();
    if (!tasks.isEmpty()) {
        QByteArray data;
        for(const QString& task : tasks) {
            data.append(task.toUtf8() + "\n");
        }
        m_process->write(data);
        m_process->closeWriteChannel(); // 写入完成后必须关闭写入通道
    }
}

void MainWindow::onProcessFinished(int exitCode, QProcess::ExitStatus exitStatus)
{
    // 根据不同的退出状态给出不同的反馈
    if (exitStatus == QProcess::CrashExit) { // 用户通过“取消”按钮终止
        statusBar()->showMessage("处理已由用户取消。", 5000);
        ui->plainTextEdit->appendPlainText("\n--- 处理已取消 ---");
        for (int i = 0; i < ui->fileListTableWidget->rowCount(); ++i) {
            if (ui->fileListTableWidget->item(i, 1)->text() == "正在处理...") {
                ui->fileListTableWidget->item(i, 1)->setText("已取消");
            }
        }
    } else if (exitCode == 0) { // 正常完成
        statusBar()->showMessage("处理成功！", 5000);
        ui->plainTextEdit->appendPlainText("\n--- 处理成功！ ---");
    } else { // 后端出错
        statusBar()->showMessage("处理失败！详情见日志区。", 5000);
        ui->plainTextEdit->appendPlainText("\n--- 处理失败！ ---");
        ui->plainTextEdit->appendPlainText("后端进程异常退出，错误码: " + QString::number(exitCode));
    }

    updateUiForProcessingState(false); // 解锁UI
    m_process->deleteLater(); // 安全地删除进程对象
    m_process = nullptr;

    // 滚动日志到底部
    QScrollBar *scrollbar = ui->plainTextEdit->verticalScrollBar();
    scrollbar->setValue(scrollbar->maximum());
}

void MainWindow::onProcessError(QProcess::ProcessError error)
{
    // 只处理启动失败的错误
    if (error == QProcess::FailedToStart && m_process) {
        QMessageBox::critical(this, "启动失败", "无法启动后端引擎。\n请检查程序目录下 'backend_engine' 文件夹及其中的 'main_processor.exe' 是否存在。");
        updateUiForProcessingState(false);
        m_process->deleteLater();
        m_process = nullptr;
    }
}

void MainWindow::readProcessOutput()
{
    if (!m_process) return;
    QByteArray data = m_process->readAllStandardOutput();
    QString output = QString::fromUtf8(data);
    QStringList lines = output.split('\n', Qt::SkipEmptyParts);
    for (const QString &line : lines) {
        // 解析协议，更新UI状态
        if (line.startsWith("##STATUS##|")) {
            QStringList parts = line.split('|');
            if (parts.count() >= 4) {
                findAndUpdatetTableRow(parts[1], parts[2], parts[3]);
            }
        } else { // 普通日志直接显示
            ui->plainTextEdit->insertPlainText(line + "\n");
        }
    }
    ui->plainTextEdit->verticalScrollBar()->setValue(ui->plainTextEdit->verticalScrollBar()->maximum());
}

void MainWindow::readProcessError()
{
    if (!m_process) return;
    ui->plainTextEdit->appendPlainText("【错误】: " + QString::fromLocal8Bit(m_process->readAllStandardError()));
}

// 辅助函数

void MainWindow::updateUiForMode(bool isDirectoryMode)
{
    // 根据模式，启用/禁用对应的UI控件组
    ui->browseInputDirButton->setEnabled(isDirectoryMode);
    ui->inputDirPathLineEdit->setEnabled(isDirectoryMode);
    ui->refreshDirButton->setEnabled(isDirectoryMode);
    ui->addFilesButton->setEnabled(!isDirectoryMode);
    ui->fileListTableWidget->setAcceptDrops(!isDirectoryMode);

    // 更新表头提示文本
    if (isDirectoryMode) {
        ui->fileListTableWidget->setHorizontalHeaderLabels({"", "状态", "文件名", "文件路径(指定输入文件夹模式下不支持拖入)"});
    } else {
        ui->fileListTableWidget->setHorizontalHeaderLabels({"", "状态", "文件名", "文件路径(可拖入文件/文件夹)"});
    }
}

void MainWindow::scanAndPopulateFiles(const QString &dirPath)
{
    if (dirPath.isEmpty() || !QDir(dirPath).exists()) {
        m_fileList.clear();
        ui->fileListTableWidget->setRowCount(0);
        updateActionButtonStates();
        return;
    }

    m_fileList.clear();
    ui->fileListTableWidget->setRowCount(0);

    // 在耗时操作前后，改变鼠标光标和状态栏提示，提升用户体验
    QApplication::setOverrideCursor(Qt::WaitCursor);
    statusBar()->showMessage("正在扫描文件夹，请稍候...");
    QApplication::processEvents(); // 强制UI刷新，确保提示能立即显示

    QDirIterator it(dirPath, {"*.csv", "*.xlsx"}, QDir::Files, QDirIterator::Subdirectories);
    while (it.hasNext()) {
        addFileToList(it.next());
    }

    QApplication::restoreOverrideCursor();
    statusBar()->showMessage("扫描完成。", 3000);
    updateActionButtonStates();
}

void MainWindow::updateUiForProcessingState(bool isProcessing)
{
    m_isProcessing = isProcessing;
    ui->startProcessButton->setText(isProcessing ? "取消处理" : "开始处理");

    // --- 统一锁定/解锁所有交互控件 ---

    // 模式切换和配置控件
    ui->useDirectoryModeCheckBox->setEnabled(!isProcessing);
    ui->outputToSourceCheckBox->setEnabled(!isProcessing);
    ui->conflictComboBox->setEnabled(!isProcessing);

    // 文件夹模式控件
    ui->browseInputDirButton->setEnabled(!isProcessing && ui->useDirectoryModeCheckBox->isChecked());
    ui->inputDirPathLineEdit->setEnabled(!isProcessing && ui->useDirectoryModeCheckBox->isChecked());
    ui->refreshDirButton->setEnabled(!isProcessing && ui->useDirectoryModeCheckBox->isChecked());

    // 手动模式控件
    ui->addFilesButton->setEnabled(!isProcessing && !ui->useDirectoryModeCheckBox->isChecked());
    ui->fileListTableWidget->setAcceptDrops(!isProcessing && !ui->useDirectoryModeCheckBox->isChecked());

    // 列表操作控件
    bool hasFiles = ui->fileListTableWidget->rowCount() > 0;
    ui->removeSelectedButton->setEnabled(!isProcessing && hasFiles);
    ui->selectAllButton->setEnabled(!isProcessing && hasFiles);
    ui->invertSelectionButton->setEnabled(!isProcessing && hasFiles);

    // 输出路径控件
    if (!ui->outputToSourceCheckBox->isChecked()) {
        ui->outputPathLineEdit->setEnabled(!isProcessing);
        ui->browseOutputButton->setEnabled(!isProcessing);
    }
}

void MainWindow::updateActionButtonStates()
{
    bool hasFiles = (ui->fileListTableWidget->rowCount() > 0);
    ui->removeSelectedButton->setEnabled(hasFiles);
    ui->selectAllButton->setEnabled(hasFiles);
    ui->invertSelectionButton->setEnabled(hasFiles);
    ui->startProcessButton->setEnabled(hasFiles);
}

void MainWindow::findAndUpdatetTableRow(const QString &filePath, const QString &status, const QString &message)
{
    for (int i = 0; i < ui->fileListTableWidget->rowCount(); ++i) {
        QTableWidgetItem *pathItem = ui->fileListTableWidget->item(i, 3);
        if (pathItem && pathItem->text() == filePath) {
            QTableWidgetItem *statusItem = ui->fileListTableWidget->item(i, 1);
            if(statusItem) {
                // 根据后端协议更新状态文本
                if (status == "PROCESSING") statusItem->setText("正在处理...");
                else if (status == "SUCCESS") statusItem->setText("处理完成");
                else if (status == "FAILURE") statusItem->setText("处理失败");
                else if (status == "SKIPPED") statusItem->setText("已跳过");
                else if (status == "UNIDENTIFIED") statusItem->setText("未知平台");
                else statusItem->setText(status); // 其他未知状态直接显示

                statusItem->setToolTip(message); // 将详细信息设置到Tooltip
            }
            return;
        }
    }
}
