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


   // QString tableName = "login_information";
  //  QString fileName = "D:/账号.xlsx";
    //if(exportData2XLSX(fileName, tableName)){
   //     qDebug()<<"导入excel成功";
   // }else{
   //     qDebug()<<"导入excel失败";
  //  }


    QString tableName = "login_information";
    QString fileName = "D:/账号.xlsx";
    if(exportData2XLSX(fileName, tableName)){

    }

    return a.exec();
}
