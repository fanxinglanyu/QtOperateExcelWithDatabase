#ifndef CREATE_CONNECTION_H
#define CREATE_CONNECTION_H

#include<QSqlDatabase>
#include<QCoreApplication>
#include<QFile>
#include<QSysInfo>
#include<QtGlobal>
#include<QDebug>

#include<QTimer>
#include<QObject>
#include <QSqlQuery>

static bool CreateConnection(){
    //        qDebug()<<"查看目前可用驱动";
    //        QStringList drivers = QSqlDatabase::drivers();
    //        for(auto driver: drivers){
    //            qDebug()<<driver<<" ";
    //        }

    //设置数据库驱动
    QSqlDatabase mysqlDB = QSqlDatabase::addDatabase("QMYSQL", "mysql_connection1");
    mysqlDB.setHostName("192.168.0.104");
    mysqlDB.setUserName("root");
    mysqlDB.setPassword("password");
    mysqlDB.setPort(8889);
    mysqlDB.setDatabaseName("test_db");

    //根据系统环境设计数据库路径
    // Q_OS_LINUX：Q_OS_WIN： Q_OS_MAC   Q_OS_WIN32

    //如果远程mysql数据库没有打开
    if(!mysqlDB.open()){
        return false;
    }else{

        // #ifdef Q_OS_WIN
        mysqlDB.exec("SET NAMES 'GBK'");
        // #endif
        // #ifdef Q_OS_MAC

        // #endif
    }

    //QSqlDatabase sqliteDB = QSqlDatabase::addDatabase("QSQLITE", "sqlite_connection1");

    // #ifdef Q_OS_WIN //Q_OS_WIN32
    //     qDebug()<<"QCoreApplication::applicationDirPath():"<<QCoreApplication::applicationDirPath();
    //     sqliteDB.setDatabaseName(QCoreApplication::applicationDirPath() + QString("/database/LocalSystemDatabse.db"));
    // #endif

    //如果本地sqlite数据库没有打开
    // if(!sqliteDB.open()){
    //     QMessageBox* databaseInformationBox = new QMessageBox(QMessageBox::Critical, ("信息提示"),  ("不能建立本地数据库连接！"), QMessageBox::Yes);
    //     auto button =  databaseInformationBox->exec();
    //     if(button == QMessageBox::Yes){
    //         databaseInformationBox->deleteLater();
    //     }
    //     return false;
    // }

    return true;

}

#endif // CREATE_CONNECTION_H
