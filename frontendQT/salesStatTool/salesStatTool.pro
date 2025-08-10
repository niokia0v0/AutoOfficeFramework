QT       += core gui

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

CONFIG += c++11

# The following define makes your compiler emit warnings if you use
# any Qt feature that has been marked deprecated (the exact warnings
# depend on your compiler). Please consult the documentation of the
# deprecated API in order to know how to port your code away from it.
DEFINES += QT_DEPRECATED_WARNINGS

# You can also make your code fail to compile if it uses deprecated APIs.
# In order to do so, uncomment the following line.
# You can also select to disable deprecated APIs only up to a certain version of Qt.
#DEFINES += QT_DISABLE_DEPRECATED_BEFORE=0x060000    # disables all the APIs deprecated before Qt 6.0.0

SOURCES += \
    main.cpp \
    mainwindow.cpp

HEADERS += \
    mainwindow.h

FORMS += \
    mainwindow.ui

# ===================================================================
#           自动化部署脚本
# ===================================================================

# 只在Windows平台的Release模式下生效
win32:CONFIG(release, debug|release) {

    # 1. 指定最终可执行文件的生成目录 (核心修改)
    # $$PWD 代表当前 .pro 文件所在的目录。
    # 使用相对路径向上两级，进入 'Releases' 文件夹，再进入 'salesStatTool' 文件夹。
    # 强制 qmake 将 salesStatTool.exe 直接生成在这个目录下。
    DESTDIR = $$PWD/../../Releases/salesStatTool

    # 2. 定义链接后要执行的命令
    # 因为 .exe 已经生成在目标目录了，所以 .bat 脚本也应该在这个目录里执行。
    # $$shell_path() 用于处理路径中的空格和斜杠，确保在命令行中正确传递。
    QMAKE_POST_LINK = "\"$$shell_path($$PWD/deploy.bat)\" \"$$shell_path($$DESTDIR)\" \"$$TARGET\" \"$$shell_path($$[QT_INSTALL_BINS])\""
}
