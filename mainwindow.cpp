#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    this->setFixedSize(562,220);
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::on_pushButton_clicked()
{
    int count = 0 ,flag = 0;
    QSqlDatabase db = QSqlDatabase::addDatabase("QODBC");
    QString cmd = "select MD001,I.MB002 AS MMB002,MD003,II.MB002 AS CMB002,II.MB003,isNull(MD006,0) as MD006,MD015,MD019,MD020,MD021,MD022,MD023,V.MA002,isNull(II.MB036,0) as MB036,II.MB009,II.MB039,I.MB009 as Idesc from BOMMD left join INVMB I on MD001=I.MB001 left join INVMB II on MD003=II.MB001 left join PURMA V on II.MB032=V.MA001 where MD001 in ('" + ui->lineEdit->text() + "') and ( (MD011<=convert(CHAR(8),getdate(),112) or MD011='') and (MD012>convert(CHAR(8),getdate(),112) or MD012='') ) ";
    QString dsn("DRIVER={SQL Server};SERVER=203.70.94.121;UID=SQLreader;PWD=12170;DATABASE=ADLINK");
//    int clevel = 1;
    db.setDatabaseName(dsn);
    if(!db.open())
    {
        qDebug()<<" XXXXXXXXXXXXXXXXX";
        QMessageBox::information(this, "Message", "Open database error.");
        return;
    }
    else
    {
        QSqlQuery sq(db);
//        IniTabelView();
        sq.exec(cmd + " order by MD001,MD003");
        while(sq.isActive())
        {
            QString level;
            QString bPN;
            while(sq.next())
            {
                qDebug()<< count++;
                PNDiscription = sq.value(1).toString().simplified();
                qDebug()<< PNDiscription;
                bPN = sq.value(2).toString().simplified();
                level = level + "'" + bPN + "',";

                if(PNDiscription.contains("(EA)") || PNDiscription.contains("(GEA)") || PNDiscription.contains("(L)") || PNDiscription.contains("(EP)") || PNDiscription.contains("(CA)") || PNDiscription.contains("(G)"))
                {
                    flag = 1;
                    break;
                }
//                model->setRowCount(model->rowCount() + 1);
//                model->setData(model->index(model->rowCount() - 1, 0), QString::number(clevel));
//                model->setData(model->index(model->rowCount() - 1, 1), sq.value(0).toString().simplified());
//                model->setData(model->index(model->rowCount() - 1, 2), sq.value(1).toString().simplified());
//                model->setData(model->index(model->rowCount() - 1, 3), bPN);
//                model->setData(model->index(model->rowCount() - 1, 4), sq.value(3).toString().simplified());
//                model->setData(model->index(model->rowCount() - 1, 5), sq.value(4).toString().simplified());
//                model->setData(model->index(model->rowCount() - 1, 6), sq.value(5).toString().simplified());

            }
            if(flag == 1) break;
            level = level.remove(level.length() - 1, 1);    //移除最後一個逗號
//            clevel++;
            sq.clear();
            cmd = "select MD001,I.MB002 AS MMB002,MD003,II.MB002 AS CMB002,II.MB003,isNull(MD006,0) as MD006,MD015,MD019,MD020,MD021,MD022,MD023,V.MA002,isNull(II.MB036,0) as MB036,II.MB009,II.MB039 from BOMMD left join INVMB I on MD001=I.MB001 left join INVMB II on MD003=II.MB001 left join PURMA V on II.MB032=V.MA001 where MD001 in (" + level + ")  and ( (MD011<=convert(CHAR(8),getdate(),112) or MD011='') and (MD012>convert(CHAR(8),getdate(),112) or MD012='') ) order by MD001,MD003";
            sq.exec(cmd);
        }
//        ui->tableView->resizeColumnsToContents();
        sq.clear();
        db.close();
    }
    if(count == 0 || flag == 0)
    {
        QMessageBox::warning(this, "Message", "Could not search BOM.");
        return;
    }

    PNpart = ui->lineEdit->text().section('-', 2, 2);

    if(PNDiscription.contains("(G)"))
    {
        GP = "GP";
    }else
    {
        GP = "LF";
    }
    ui->label_2->setText(PNpart + GP + "\n" + PNDiscription);
    qDebug() << PNpart + GP ;
    printLabel(GP, PNpart);
}

void MainWindow::printLabel(const QString &GP, const QString &PN)
{
    QAxObject bt("BarTender.Application");

    bt.setProperty("Visible", false);

    QAxObject *format = bt.querySubObject("Formats");

    QAxObject *open = format->querySubObject("Open(QVariant, QVariant, QVariant)", QDir::currentPath() + "/Sample.btw", false, "");

    open->dynamicCall("SetNamedSubStringValue(QVariant, QVariant)", "GP", GP);

    open->dynamicCall("SetNamedSubStringValue(QVariant, QVariant)", "PN", PN);

    open->dynamicCall("PrintOut(QVariant, QVariant)", true, true);

    bt.dynamicCall("Quit(QVariant)", 1);
}

void MainWindow::keyPressEvent(QKeyEvent *event)
{
    if(event->key() == Qt::Key_F1)
    {
        QMessageBox::information(this, "Information", "Enviroment request: Windows XP/7/8/10 & ODBC driver");
    }
}


