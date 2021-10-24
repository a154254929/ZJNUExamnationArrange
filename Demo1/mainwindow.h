#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include<QStandardItem>
QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();
    QVariantList varRows;
private slots:
    void on_chooseFile_clicked();
    void on_generateTable_clicked();
    void setMessage(QString message);
private:
    Ui::MainWindow *ui;
    QString fileName;
    QStandardItemModel *ItemModel;
};
#endif // MAINWINDOW_H
