#include <QApplication>
#include <QDebug>

#include <QAxObject>
#include <QAxWidget>
#include <QString>

struct UserInfo
{
    QString name;  // 用户名
    int total;  // 密码
};

UserInfo g_userInfo = {0, 0};

void parseArguments()
{
    // 获取命令行参数
    QStringList arguments = QCoreApplication::arguments();

    qDebug() << "Arguments : " << arguments;

    if (arguments.count() < 2)
        return;

     g_userInfo.name = arguments.at(1);

    qDebug() << "startNumber : " << g_userInfo.name;

    g_userInfo.total = arguments.at(2).toInt();

    qDebug() << "total : " << g_userInfo.total;
}

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);

    parseArguments();

    QString label;

    for (int i = 0; i < g_userInfo.total; ++i) {
        label = QByteArray::number(g_userInfo.name.toInt() + i, 10).insert(0, 2, '0');

        qDebug() << "Label : " << label;
        QAxObject *word = new QAxObject();
        word->setControl("word.Application");

    //    if (!bFlag) {
    //        return false;
    //    }

        word->setProperty("Visible", false);

        QAxObject *document = word->querySubObject("Documents");

    //    if (!document) {
    //        return false;
    //    }

        QString sFile(QObject::tr("E:/myCoding/Qt/build-ForECN-Desktop_Qt_5_12_3_MinGW_64_bit-Debug/debug/0015程序文件.dotx"));
        document->dynamicCall("Add(QString)", sFile);

        QAxObject *workDocument = word->querySubObject("ActiveDocument");

        QAxObject *nameLabel = workDocument->querySubObject("Bookmarks(QString)",QObject::tr("name"));
        if(nameLabel)
        {
            nameLabel->dynamicCall("Select(void)");
            nameLabel->querySubObject("Range")->setProperty("Text",  QObject::tr("%1程序文件").arg(label));
            delete nameLabel;
        }

        QAxObject *numberLabel = workDocument->querySubObject("Bookmarks(QString)",QObject::tr("number"));
        if(numberLabel)
        {
            numberLabel->dynamicCall("Select(void)");
            numberLabel->querySubObject("Range")->setProperty("Text",  QObject::tr("WI-ERD-%0").arg(label));
            delete numberLabel;
        }

    //    QAxObject *document = word->querySubObject("Documents");

        workDocument->dynamicCall("SaveAs (const QString&)", QObject::tr("E:/myCoding/Qt/build-ForECN-Desktop_Qt_5_12_3_MinGW_64_bit-Debug/debug/%1程序文件.docx").arg(label));
        QVariant OutputFileName(QObject::tr("E:/myCoding/Qt/build-ForECN-Desktop_Qt_5_12_3_MinGW_64_bit-Debug/debug/%1程序文件.pdf").arg(label));
            QVariant ExportFormat(17);      //17是pdf
            QVariant OpenAfterExport(false); //保存后是否自动打开
            //调用接口进行转换
            workDocument->querySubObject("ExportAsFixedFormat(const QVariant&,const QVariant&,const QVariant&)",
                                            OutputFileName,
                                            ExportFormat,
                                            OpenAfterExport
                                            );

        if (word)
            word->setProperty("DisplayAlerts", true);
        if (workDocument)
            workDocument->dynamicCall("Close(bool)", true);

        if (word)
            word->dynamicCall("Quit()");
        if (workDocument)
            delete workDocument;
        if (word)
            delete word;
    }

    return a.exec();
}
