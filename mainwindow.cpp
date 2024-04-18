#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <qsqlrecord>
#include <QMessageBox>
#include <windows.h>
#include <tchar.h>
#include <stdio.h>
#include "Sorting.h"
#include "OpenApps.h"
#include <QApplication>

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    db = QSqlDatabase::addDatabase("QSQLITE");
    db.setDatabaseName("./TourDb");
    if(db.open())
    {
        qDebug("open");
    }
    else
    {
        qDebug("close");
    }

    query = new QSqlQuery(db);
    query->exec("CREATE TABLE Touring(Маршрут TEXT, Прізвище TEXT, Відстань маршруту FLOAT, Ціна путівки INT )");

    model = new QSqlTableModel(this,db);
    model -> setTable("Touring");
    model -> select();


    ui->tableView->setModel(model);
    ui->tableView->resizeColumnsToContents(); // Автоматично регулює ширину стовпців



}


void InsertionSort(QSqlTableModel* model , int column) {

    int rowCount = model->rowCount();

    for (int i = 0; i < rowCount - 1; ++i) {
        int minIndex = i;
        int minPrice = model->data(model->index(i, model->fieldIndex("Ціна"))).toInt();

        for (int j = i + 1; j < rowCount; ++j) {
            int currentPrice = model->data(model->index(j, model->fieldIndex("Ціна"))).toInt();

            if (currentPrice < minPrice) {
                minIndex = j;
                minPrice = currentPrice;
            }
        }

        if (minIndex != i) {
            for (int column = 0; column < model->columnCount(); ++column) {
                QVariant temp = model->data(model->index(i, column));
                model->setData(model->index(i, column), model->data(model->index(minIndex, column)));
                model->setData(model->index(minIndex, column), temp);
            }
        }
    }
}


bool OpenExcel(int argc, TCHAR *argv[]) {
    // Путь к Excel
    TCHAR excelPath[] = TEXT("C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE");

    STARTUPINFO si;
    PROCESS_INFORMATION pi;

    ZeroMemory(&si, sizeof(si));
    si.cb = sizeof(si);
    ZeroMemory(&pi, sizeof(pi));

    // Запускаем Excel без указания файла
    if (!CreateProcess(NULL, excelPath, NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi)) {
        printf("CreateProcess failed (%d).\n", GetLastError());
        return 1;
    }

    // Закрываем дескрипторы процесса и потока
    // CloseHandle(pi.hProcess);
    // CloseHandle(pi.hThread);

    return 0;
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::on_Add_clicked()
{
    model->insertRow(model->rowCount());
}


void MainWindow::on_Delete_clicked()
{
    model->removeRow(row);
}


void MainWindow::on_tableView_clicked(const QModelIndex &index)
{
    row = index.row();
}


void MainWindow::on_Sort_clicked()
{


    int priceColumnIndex = model->fieldIndex("Ціна");
    model->setSort(priceColumnIndex, Qt::AscendingOrder);
    model->select(); // Обновляем модель, чтобы отобразить отсортированные данные


    InsertionSort(model, priceColumnIndex);

}





void MainWindow::on_Average_clicked()
{
    int rowCount = model->rowCount();
    int priceColumnIndex = model->fieldIndex("Ціна");

    int total = 0;

    for (int i = 0; i < rowCount; ++i) {
        QModelIndex index = model->index(i, priceColumnIndex);
        total += model->data(index).toInt();
    }

    double average = static_cast<double>(total) / rowCount;

    ui->lineEdit->setText(QString::number(average));
}



void MainWindow::on_Search_clicked()

{
    QString surname = ui->searchLine->text(); // Отримати прізвище з lineEdit

    // Виконати запит до бази даних
    QSqlQuery query;
    query.prepare("SELECT * FROM Touring WHERE Прізвище = ?");
    query.addBindValue(surname);

    if (query.exec()) {
        // Вдало виконаний запит
        QSqlTableModel *searchModel = new QSqlTableModel(this, db);
        searchModel->setQuery(query);
        ui->tableView->setModel(searchModel); // Встановити модель для відображення результатів
    } else {
        QMessageBox searchingError;
        searchingError.setText("Помилка пошуку! /n Введіть інше значення або переконайтесь що воно існує");
        searchingError.exec();
    }
}



void MainWindow::on_Download_clicked()
{

    //Звертаємось через базу данних до функції select, щоб оновити таблицю
    model->select();

}


void MainWindow::on_OpenExcel_clicked()
{
    Excel::OpenExcel();
}



void MainWindow::on_OpenWord_clicked()
{
    Word::OpenWord();
}


void MainWindow::on_OpenAccess_clicked()
{
    Access::OpenAccess();
}


void MainWindow::on_Close_triggered()
{
   QApplication::quit();
}


void MainWindow::on_actionFullInfo_triggered()
{
    QMessageBox msgBox;
    msgBox.setText("Курсовий проєкт з дисципліни програмування \nСтворив Огли Рустам, група КІ-2-21\n Керівник проєкту Віктор Герасимюк");
    msgBox.exec();
}


void MainWindow::on_actionInform_triggered()
{
    QMessageBox msgBox;
    msgBox.setText("Керівник проекту - викладач Павлоградського фахового колледжу Віктор Герасимюк");
    msgBox.exec();
}


void MainWindow::on_action_Excel_triggered()
{
    Excel::OpenExcel();
}


void MainWindow::on_action_Word_triggered()
{
    Word::OpenWord();
}


void MainWindow::on_action_Access_triggered()
{
    Access::OpenAccess();
}


void MainWindow::on_actionSort_triggered()
{
    int priceColumnIndex = model->fieldIndex("Ціна");
    model->setSort(priceColumnIndex, Qt::AscendingOrder);
    model->select(); // Обновляем модель, чтобы отобразить отсортированные данные


    InsertionSort(model, priceColumnIndex);
}


void MainWindow::on_actionAverage_triggered()
{

    int rowCount = model->rowCount();
    int priceColumnIndex = model->fieldIndex("Ціна");

    int total = 0;

    for (int i = 0; i < rowCount; ++i) {
        QModelIndex index = model->index(i, priceColumnIndex);
        total += model->data(index).toInt();
    }

    double average = static_cast<double>(total) / rowCount;

    ui->lineEdit->setText(QString::number(average));

}


void MainWindow::on_actionInfoApp_triggered()
{
    QMessageBox msgBox;
    msgBox.setText("База даних Туристичної фірми");
    msgBox.exec();
}

