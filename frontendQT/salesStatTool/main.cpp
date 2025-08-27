#include "mainwindow.h"

#include <QApplication>
#include <QDir>

int main(int argc, char *argv[])
{
    // 获取可执行文件所在的目录
    QString appDir = QCoreApplication::applicationDirPath();
    // 构建 frontend_runtime 文件夹的绝对路径
    QString libraryPath = QDir(appDir).filePath("frontend_runtime");
    // 将这个新路径添加到Qt的库搜索路径列表中
    QCoreApplication::addLibraryPath(libraryPath);

    QApplication a(argc, argv);
    MainWindow w;
    w.show();
    return a.exec();
}
