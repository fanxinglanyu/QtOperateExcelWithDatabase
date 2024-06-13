@[toc]
# 0 背景
因为需要批量导入和导出数据，所以需要用到excel。实现把数据库的数据导入到excel中，把excel中的数据导出到数据库。这里使用了开源代码库QXlsx。

# 1 准备QXlsx环境
官网中的[qmake的使用方法](https://github.com/QtExcel/QXlsx/blob/master/HowToSetProject.md)，[cmake的使用方法。](https://github.com/QtExcel/QXlsx/blob/master/HowToSetProject-cmake.md)

## 1.1 cmake安装使用

* 1，输入下列指令安装：

```shell
mkdir build
cd build
cmake ../QXlsx/ -DCMAKE_INSTALL_PREFIX=... -DCMAKE_BUILD_TYPE=Release
cmake --build .
cmake --install .
```

在CMakeLists.txt中添加如下内容:

```shell
find_package(QXlsxQt5 REQUIRED) # or QXlsxQt6
target_link_libraries(myapp PRIVATE QXlsx::QXlsx)
```


* 2，下面是无需安装的两种使用方法：

使用cmake的子目录在 CMakeLists.txt:

```powershell
add_subdirectory(QXlsx)
target_link_libraries(myapp PRIVATE QXlsx::QXlsx)
```

使用 cmake FetchContent 在 CMakeLists.txt:

```powershell
FetchContent_Declare(
  QXlsx
  GIT_REPOSITORY https://github.com/QtExcel/QXlsx.git
  GIT_TAG        sha-of-the-commit
  SOURCE_SUBDIR  QXlsx
)
FetchContent_MakeAvailable(QXlsx)
target_link_libraries(myapp PRIVATE QXlsx::QXlsx)
```


如果 `QT_VERSION_MAJOR`没有设置, QXlsx's的 CMakeLists.txt 将尝试自己寻找 Qt 版本（5 或 6）。

## 1.2 qmake使用
下载[QXsx的github项目代码](https://github.com/QtExcel/QXlsx)。

* 1，把QXsx项目中的代码（选中的三个项目）复制到自己项目下；

![在这里插入图片描述](https://img-blog.csdnimg.cn/direct/7b82c67b7a034d4bba386335622dc7ae.png)


复制到自己项目下（新建一个QXlxs文件夹，存储文件）：
![在这里插入图片描述](https://img-blog.csdnimg.cn/direct/50013e91ccc24e48ac446fa2bb66fcd7.png)
![在这里插入图片描述](https://img-blog.csdnimg.cn/direct/438a281c6adc4f5ba8689047a81d99ae.png)

* 2，在pro中添加如下代码；

```cpp
QXLSX_PARENTPATH=./         # current QXlsx path is . (. means curret directory)
QXLSX_HEADERPATH=./QXlsx/header/  # current QXlsx header path is ./header/
QXLSX_SOURCEPATH=./QXlsx/source/  # current QXlsx source path is ./source/
include(./QXlsx/QXlsx.pri)

```
* 3，编译文件后，会自动把文件添加到项目中（绿色的那一部分）；

![在这里插入图片描述](https://img-blog.csdnimg.cn/direct/c9a8aac56507416bb39c18b1cfa1e2d0.png)

* 4，添加如下头文件，就可以开始项目编写；

```cpp
#include "xlsxdocument.h"
#include "xlsxchartsheet.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"
#include "xlsxrichstring.h"
#include "xlsxworkbook.h"
```

测试程序：

```cpp
// main.cpp

#include <QCoreApplication>

#include "xlsxdocument.h"
#include "xlsxchartsheet.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"
#include "xlsxrichstring.h"
#include "xlsxworkbook.h"
using namespace QXlsx;

int main(int argc, char *argv[])
{
    QCoreApplication a(argc, argv);

    QXlsx::Document xlsx;
    xlsx.write("A1", "Hello Qt!"); // write "Hello Qt!" to cell(A,1). it's shared string.
    xlsx.saveAs("Test.xlsx"); // save the document as 'Test.xlsx'

    return 0;
    // return a.exec();
}
```

# 2 把excel数据导出到mysql数据库

* 1，准备要导入的账号和密码的excel表（第一行为数据库的字段名，必须一样；如果数据库中字段值不能为空，excel中数据也不能为空）；

![在这里插入图片描述](https://img-blog.csdnimg.cn/direct/188fe0b08f724b7fb294153a71ce5b2c.png)
账号信息.xlsx

![在这里插入图片描述](https://img-blog.csdnimg.cn/direct/6a9abec6dcb54f4c89e376c5c99574d1.png)
数据库中的login_information表

* 2，在数据库中创建表格；

```sql
DROP TABLE IF EXISTS `login_information`;
CREATE TABLE `login_information`  (
  `account` varchar(255) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci NOT NULL,
  `password` varchar(255) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci NOT NULL,
  PRIMARY KEY (`account`) USING BTREE
) ENGINE = InnoDB CHARACTER SET = utf8mb4 COLLATE = utf8mb4_unicode_ci ROW_FORMAT = DYNAMIC;

```

* 3，建立数据库连接；

方法：

```cpp
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
```

调用：

```cpp
//main中创建数据库连接
int main(int argc, char *argv[])
{
    QCoreApplication a(argc, argv);

    //连接数据库
    if (!CreateConnection()){
        qDebug()<<"数据库连接失败";
    }
    return a.exec();
}
```

* 4，把excel中的数据导入到数据库中；

```cpp
bool exportExcel2Database(QStringList filePaths,  QString xlsxName, QString sqlSentence){
    QList<bool> execResultList;//操作的结果集
    bool execResult = false;
    QSqlDatabase db = QSqlDatabase::database("mysql_connection1");
    QSqlQuery query(db);
    if(db.transaction()){
        foreach(QString filePath, filePaths) {
            QXlsx::Document xlsx(filePath);

            if(!xlsx.selectSheet(xlsxName)){/*在当前打开的xlsx文件中，找一个名字为ziv的sheet*/
                //xlsx.addSheet(xlsxName);//找不到的话就添加一个名为ziv的sheet
                qDebug()<<"没有对应的xlsx表";
                return false;
            }else{

            }

            QQueue<QString> tableFieldQueue;
            QHash<QString, QVariantList> tableAlterFiledValue;

            for(int row = 1; row <= xlsx.dimension().rowCount(); row++) {
                // 获取每行的数据并插入到数据库中
                for(int col = 1; col <= xlsx.dimension().columnCount();col++){
                    if(row == 1){
                        tableFieldQueue.enqueue(xlsx.read(row, col).toString());

                    }else{
                        tableAlterFiledValue[tableFieldQueue[col-1]].append(xlsx.read(row, col));
                    }
                }
            }

            query.prepare(sqlSentence);
            foreach (QString tableFiled, tableFieldQueue) {
                query.addBindValue(tableAlterFiledValue[tableFiled]);
            }

            execResult = query.execBatch();
            execResultList.append(execResult);
            if(!execResult) {//批量执行数据插入
                qDebug() <<  query.lastError().databaseText();
            }

        }

        foreach (bool result, execResultList) {
            if(result == false){
                if(!db.rollback())
                {
                    qDebug() << "数据库回滚失败"<<db.lastError().databaseText(); //回滚
                }else{
                    qDebug()<<"数据库回滚成功";
                }
                return false;
            }

        }

        if(db.commit()){
            return true;
        }else{
            return false;
        }
    }
    return false;
}

```

调用：

```cpp
    QStringList filePaths;
    filePaths<<"D:/test/账号信息.xlsx";

    //考试细节步骤
    QString sql2 = QString("INSERT INTO  login_information(account, password)  VALUES (?, ?)");
    QString xlsxName2 = "账号信息";
    // qDebug()<<sql2;

    if(exportExcel2Database(filePaths, xlsxName2, sql2)){
        qDebug()<<"导入成功";
    }else{
        qDebug()<<"导入失败";
    }
```
,
# 3 把mysql数据库的数据写入到excel
* 1，建立数据库连接，同上；

* 2，把数据库中表的数据导出到excel中；

```cpp
bool exportData2XLSX(QString fileName, QString tableName)
{

    QXlsx::Document xlsx;
    QXlsx::Format format1;/*设置标题单元的样式*/
    format1.setFontSize(12);/*设置字体大小*/
    format1.setHorizontalAlignment(QXlsx::Format::AlignHCenter);/*横向居中*/
    //format1.setBorderStyle(QXlsx::Format::BorderThin);/*边框样式*/
    //format1.setFontBold(true);/*设置加粗*/

    if(!xlsx.selectSheet("表格数据")){/*在当前打开的xlsx文件中，找一个名字为ziv的sheet*/
        xlsx.addSheet("表格数据");//找不到的话就添加一个名为ziv的sheet
    }

    QSqlDatabase db = QSqlDatabase::database("mysql_connection1");
    QString tmpSql = QString("SELECT * FROM %1").arg(tableName);
    QSqlQuery query(db);
    if(query.exec(tmpSql)){
        //表头列
        QSqlRecord queryRecord(query.record());
        qDebug()<<"queryRecord.count():"<<queryRecord.count();
        for(int colNum = 0; colNum < queryRecord.count(); colNum++){
            //qDebug() <<  queryRecord.fieldName(colNum);
            xlsx.write(1, colNum+1,  queryRecord.fieldName(colNum),format1);
        }

        //表格数据
        int rowNum = 2;
        while(query.next()){
            for(int colNum = 0; colNum < queryRecord.count(); colNum++){
                xlsx.write(rowNum, colNum + 1, query.value(colNum),format1);
            }
            rowNum++;
        }
    }else{
        return false;
    }

    if(fileName.isEmpty())
        return false;

    xlsx.saveAs(fileName);//保存文件

    return true;
}

```

调用：

```cpp

int main(int argc, char *argv[])
{
    QCoreApplication a(argc, argv);

    //连接数据库
    if (!CreateConnection()){
        qDebug()<<"数据库连接失败";
    }
    
    QString tableName = "login_information";
    QString fileName = "D:/账号.xlsx";
    if(exportData2XLSX(fileName, tableName)){
        qDebug()<<"导入excel成功";
    }else{
        qDebug()<<"导入excel失败";
    }
    
    return a.exec();
}

```

# 4 完整代码

```cpp
#include <QCoreApplication>

#include "create_connection.h"

#include "xlsxdocument.h"
#include "xlsxchartsheet.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"
#include "xlsxrichstring.h"
#include "xlsxworkbook.h"

#include <QSqlError>
#include <QQueue>
#include <QHash>
#include <QSqlRecord>

bool exportExcel2Database(QStringList filePaths,  QString xlsxName, QString sqlSentence){
    QList<bool> execResultList;
    bool execResult = false;
    QSqlDatabase db = QSqlDatabase::database("mysql_connection1");
    QSqlQuery query(db);
    if(db.transaction()){
        foreach(QString filePath, filePaths) {
            QXlsx::Document xlsx(filePath);

            if(!xlsx.selectSheet(xlsxName)){/*在当前打开的xlsx文件中，找一个名字为ziv的sheet*/
                //xlsx.addSheet(xlsxName);//找不到的话就添加一个名为ziv的sheet
                qDebug()<<"没有对应的xlsx表";
                return false;
            }else{

            }

            QQueue<QString> tableFieldQueue;
            QHash<QString, QVariantList> tableAlterFiledValue;

            for(int row = 1; row <= xlsx.dimension().rowCount(); row++) {
                // 获取每行的数据并插入到数据库中
                for(int col = 1; col <= xlsx.dimension().columnCount();col++){
                    if(row == 1){
                        tableFieldQueue.enqueue(xlsx.read(row, col).toString());

                    }else{
                        tableAlterFiledValue[tableFieldQueue[col-1]].append(xlsx.read(row, col));
                    }
                }
            }

            query.prepare(sqlSentence);
            foreach (QString tableFiled, tableFieldQueue) {
                query.addBindValue(tableAlterFiledValue[tableFiled]);
            }

            execResult = query.execBatch();
            execResultList.append(execResult);
            if(!execResult) {//批量执行数据插入
                qDebug() <<  query.lastError().databaseText();
            }

        }

        foreach (bool result, execResultList) {
            if(result == false){
                if(!db.rollback())
                {
                    qDebug() << "数据库回滚失败"<<db.lastError().databaseText(); //回滚
                }else{
                    qDebug()<<"数据库回滚成功";
                }
                return false;
            }

        }

        if(db.commit()){
            return true;
        }else{
            return false;
        }
    }
    return false;
}

bool exportData2XLSX(QString fileName, QString tableName)
{

    QXlsx::Document xlsx;
    QXlsx::Format format1;/*设置标题单元的样式*/
    format1.setFontSize(12);/*设置字体大小*/
    format1.setHorizontalAlignment(QXlsx::Format::AlignHCenter);/*横向居中*/
    //format1.setBorderStyle(QXlsx::Format::BorderThin);/*边框样式*/
    //format1.setFontBold(true);/*设置加粗*/

    if(!xlsx.selectSheet("表格数据")){/*在当前打开的xlsx文件中，找一个名字为ziv的sheet*/
        xlsx.addSheet("表格数据");//找不到的话就添加一个名为ziv的sheet
    }

    QSqlDatabase db = QSqlDatabase::database("mysql_connection1");
    QString tmpSql = QString("SELECT * FROM %1").arg(tableName);
    QSqlQuery query(db);
    if(query.exec(tmpSql)){
        //表头列
        QSqlRecord queryRecord(query.record());
        qDebug()<<"queryRecord.count():"<<queryRecord.count();
        for(int colNum = 0; colNum < queryRecord.count(); colNum++){
            //qDebug() <<  queryRecord.fieldName(colNum);
            xlsx.write(1, colNum+1,  queryRecord.fieldName(colNum),format1);
        }

        //表格数据
        int rowNum = 2;
        while(query.next()){
            for(int colNum = 0; colNum < queryRecord.count(); colNum++){
                xlsx.write(rowNum, colNum + 1, query.value(colNum),format1);
            }
            rowNum++;
        }
    }else{
        return false;
    }

    if(fileName.isEmpty())
        return false;

    xlsx.saveAs(fileName);//保存文件

    return true;
}

int main(int argc, char *argv[])
{
    QCoreApplication a(argc, argv);

    //连接数据库
    if (!CreateConnection()){
        qDebug()<<"数据库连接失败";
    }
    
    QString tableName = "login_information";
    QString fileName = "D:/账号.xlsx";
    if(exportData2XLSX(fileName, tableName)){
        qDebug()<<"导入excel成功";
    }else{
        qDebug()<<"导入excel失败";
    }

    return a.exec();
}

```

