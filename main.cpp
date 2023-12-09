#include "mainwindow.h"
#include <QApplication>
#include <QDir>
#include <QFontDatabase>
#include <xlnt/xlnt.hpp>

bool validatePath();

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    if(!validatePath()){
        QString message = "مسیر برنامه دارای کارکترهای غیراستاندارد است";
        QMessageBox::critical(nullptr,"Error",message);
        return -1;
    }
    QDir fontDir (":/fonts");
    QFileInfoList infoList = fontDir.entryInfoList(QDir::Filter::Files | QDir::Filter::NoDotAndDotDot);
    for(const QFileInfo &fi:infoList){
        QFontDatabase::addApplicationFont(fi.filePath());
    }
    MainWindow w;
    w.show();
    return a.exec();
}

bool validatePath(){
    xlnt::path path (QDir::currentPath().toStdString());
    if(!path.exists()){
        return false;
    }
    return true;
}
