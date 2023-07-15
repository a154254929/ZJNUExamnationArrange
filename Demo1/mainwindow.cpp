#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QFileDialog>
#include <QAxObject>
#include <QFileDialog>
#include <QDebug>
#include <QTime>
#include <QMessageBox>
#include <algorithm>
#include <QPair>
struct student{
    QString stuId=NULL,
            name = NULL,
            className = NULL,
            college = NULL,
            professionalClass = NULL,
            examProperty = NULL;
    bool isChecked;
};
struct Class{
    QString className = NULL,
            courseName = NULL,
            teacher = NULL;
    QVector<student> stus;
};

struct course{
    QString courseName = NULL,
            courseId = NULL;
    bool chineseStuExist, internationalStuExist;
    QVector<int> Classes;
};


QVector<Class> classes;
QVector<course> courses;
QMap<QPair<QString, QString>, int> classMap;
QMap<QString, int> courseMap;
QVector<course> ordinary,math,computerA,computerB;


bool stuCmp(student A, student B){
    if(A.className != B.className){
        return A.className < B.className;
    }
    return  A.stuId < B.stuId;
}
bool classCmp(int A, int B){
    return classes[A].stus.size() < classes[B].stus.size();
}

struct examinationRoom{
    QVector<int> arrangedClasses;
    int totStudents = 0;
    int start = 0, end = 0;
};
int maxRoomLimite;
bool roomCmp(examinationRoom A, examinationRoom B){return A.totStudents > B.totStudents;}
QVector<examinationRoom> avaliableTime[14];

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    ItemModel=new QStandardItemModel(this);
    ui->message->setWordWrap(true);
    setMessage("未选中任何文件(确保选择的文件请有姓名、学号、原教学班、课程代码、上课教师、补考确认、考试性质这几项信息且信息都在sheet1中)");
}

MainWindow::~MainWindow(){
    delete ui;
}

void MainWindow::setMessage(QString message)
{
    ItemModel->appendRow(new QStandardItem(message));
    ui->message->setModel(ItemModel);
}

//读表
void MainWindow::on_chooseFile_clicked(){
    fileName=QFileDialog::getOpenFileName(this,tr("选择文件"),"",tr("Excel文件(*.xls *.xlsx)"));
    maxRoomLimite=90;
    if(fileName.isNull()){return;}
    //清空各个容器
    classes.clear();
    courses.clear();
    classMap.clear();
    courseMap.clear();
    math.clear();
    ordinary.clear();
    computerA.clear();
    computerB.clear();
    QAxObject* excel = new QAxObject(this);    //连接Excel控件
    excel->setControl("ket.Application");  //连接Excel控件
    excel->setProperty("Visible", false);  //不显示窗体
    QAxObject* workbooks = excel->querySubObject("WorkBooks");  //获取工作簿集合
    workbooks->dynamicCall("Open(const QString&)", fileName); //打开打开已存在的工作簿
    QAxObject* workbook = excel->querySubObject("ActiveWorkBook"); //获取当前工作簿
    QAxObject* sheets = workbook->querySubObject("Sheets");  //获取工作表集合，Sheets也可换用WorkSheets
    QAxObject* sheet = workbook->querySubObject("WorkSheets(int)", 1);//获取工作表集合的工作表1，即sheet1
    QAxObject* range = sheet->querySubObject("UsedRange"); //获取该sheet的使用范围对象
    QVariant var = range->dynamicCall("Value");
    excel->dynamicCall("Quit()");
    varRows = var.toList();  //得到表格中的所有数据
    if(varRows.isEmpty())
    {
        return;
    }
    ItemModel=new QStandardItemModel(this);
    setMessage("已选择"+fileName);
    /***************数据处理代码分界线*****************/
# pragma region readData{
    const int rowCount = varRows.size();
    QVariantList names = varRows[0].toList();
    const int colCount=names.size();
    //信息的位置
    int stuId = -1,     //学生学号
        stuName = -1,   //学生姓名
        check = -1,     //学生是否确认补缓考
        college = -1,   //学生所在学院
        professionalClass = -1,//学生所在专业班级;
        examProperty = -1;
    int className = -1, //教学班班级名称
        teacher = -1;//上课教师;
    int courseId = -1,//课程ID
        courseName = -1;//补考课程详细名称
    for(int i=0;i<colCount;++i)
    {
        if(names[i].toString()=="学号")stuId = i;
        if(names[i].toString()=="姓名")stuName = i;
        if(names[i].toString()=="补考确认")check=i;
        if(names[i].toString()=="课程代码")courseId = i;
        if(names[i].toString()=="课程名称")courseName = i;
        if(names[i].toString()=="原教学班")className = i;
        if(names[i].toString()=="上课教师")teacher = i;
        if(names[i].toString()=="学院")college = i;
        if(names[i].toString()=="班级")professionalClass = i;
        if(names[i].toString()=="考试性质")examProperty = i;
    }

    if(stuId == -1
        || stuName == -1
        || check == -1
        || courseId == -1
        || courseName == -1
        || className == -1
        || teacher == -1
        || college == -1
        || professionalClass == -1)
    {
        QString x="文件缺少:\n";
        if(stuId == -1)x += "学号\n";
        if(stuName == -1)x += "姓名\n";
        if(check == -1)x += "补考确认\n";
        if(courseId == -1)x += "课程代码\n";
        if(courseName == -1)x += "课程名称\n";
        if(className == -1)x += "原教学班\n";
        if(teacher == -1)x += "上课教师\n";
        if(college == -1)x += "学院\n";
        if(professionalClass == -1)x += "班级\n";
        if(examProperty == -1)x += "考试性质\n";
        QMessageBox::information(this,"文件信息缺失",x);
        setMessage("信息处理失败:文件信息不全!!!!");
        ui->generateTable->setDisabled(true);
        excel->dynamicCall("Quit()");
        delete range;
        delete sheet;
        delete sheets;
        delete workbook;
        delete workbooks;
        delete excel;
        return;
    }
    ui->generateTable->setDisabled(false);
    for(int i = 1; i < rowCount; ++i)
    {
        QVariantList rowData=varRows[i].toList();
        student tmpStu;
        tmpStu.stuId = rowData[stuId].toString();
        tmpStu.name = rowData[stuName].toString();
        tmpStu.isChecked = (rowData[check].toString() == "是");
        tmpStu.className = rowData[className].toString();
        tmpStu.college = rowData[college].toString();
        tmpStu.professionalClass = rowData[professionalClass].toString();
        tmpStu.examProperty = rowData[examProperty].toString();
        QString tmpCourseName;
        if(rowData[className].toString().contains("高等数学"))
        {
            tmpCourseName = rowData[courseName].toString();
        }
        else
        {
            int endIndex = rowData[className].toString().length();
            while(tmpStu.className[endIndex - 1] == '-' || (tmpStu.className[endIndex - 1] >= '0' && tmpStu.className[endIndex - 1] <= '9'))endIndex--;
            tmpCourseName = rowData[className].toString().mid(0, endIndex);
        }
        if(courseMap[tmpCourseName] == 0)
        {
            course tmpCourse;
            tmpCourse.courseName = tmpCourseName;
            tmpCourse.courseId = rowData[courseId].toString();
            tmpCourse.chineseStuExist = false;
            tmpCourse.internationalStuExist = false;
            courses.push_back(tmpCourse);
            courseMap[tmpCourseName] = courses.size();
            qDebug()<<tmpCourseName<<endl;
        }
        int tmpCourseOrder = courseMap[tmpCourseName] - 1;
        QString tmpClassName = rowData[className].toString();
        QString tmpTeacher = rowData[teacher].toString();
        if(classMap[{tmpCourseName, tmpTeacher}] == 0)
        {
            Class tmpClass;
            tmpClass.className = tmpClassName;
            tmpClass.courseName = tmpCourseName;
            tmpClass.teacher = tmpTeacher;
            classes.push_back(tmpClass);
            classMap[{tmpCourseName, tmpTeacher}] = classes.size();
            courses[tmpCourseOrder].Classes.push_back(classes.size() - 1);
        }
        int tmpClassOrder = classMap[{tmpCourseName, tmpTeacher}] - 1;
        classes[tmpClassOrder].stus.push_back(tmpStu);
        int length = tmpStu.name.size();
        for(int k=0;k<length;++k)
        {
            QChar cha = tmpStu.name.at(k);
            ushort uni = cha.unicode();
            if((uni >= 'a' && uni <=' z') || (uni >= 'A' &&uni <= 'Z'))
            {
                courses[tmpCourseOrder].internationalStuExist = true;
                break;
            }
            if(uni>=0x4E00&&uni<=0x9FA5)
            {
                courses[tmpCourseOrder].chineseStuExist = true;
                break;
            }
        }
    }
    excel->dynamicCall("Quit()");
    delete range;
    delete sheet;
    delete sheets;
    delete workbook;
    delete workbooks;
    delete excel;
    int numOfCourses = courses.size();
    for(int i = 0; i < numOfCourses; ++ i){
        for(int j = 0;j < courses[i].Classes.size(); ++ j){
            qSort(classes[courses[i].Classes[j]].stus.begin(), classes[courses[i].Classes[j]].stus.end(), stuCmp);
        }
        qSort(courses[i].Classes.begin(), courses[i].Classes.end(), classCmp);
        if(courses[i].courseName.contains("高等数学"))math.push_back(courses[i]);
        else if(courses[i].courseName.contains("计算机应用"))
        {
            if(courses[i].courseName.contains("A"))
            {
                computerA.push_back(courses[i]);
            }
            else if(courses[i].courseName.contains("B"))
            {
                computerB.push_back(courses[i]);
            }
        }
        else ordinary.push_back(courses[i]);
    }
# pragma endregion}
}

void MainWindow::on_generateTable_clicked(){
    setMessage("初始化中........");
    QString filePath = QFileDialog::getSaveFileName(this, "Save Data", "Schedule","Excel文件(*.xls)");
    qDebug()<<filePath<<endl;
    if(filePath.isNull()){
        setMessage("请选择保存路径!");
        return;
    }
    QAxObject *excel = new QAxObject(this);    //连接Excel控件
    excel->setControl("ket.Application");  //连接Excel控件
    excel->setProperty("Visible", false);  //不显示窗体
    QAxObject *workbooks = excel->querySubObject("Workbooks"); //获取工作簿集合
    workbooks->dynamicCall("Add");
    QAxObject *workbook = excel->querySubObject("ActiveWorkBook");
    QAxObject *sheets;
    workbook->dynamicCall("SaveAs (const QString&,int,const QString&,const QString&,bool,bool)",filePath,56,QString(""),QString(""),false,false);
    workbook->dynamicCall("Close (Boolean)",false);
    QFile file3(filePath);
    if (file3.exists())
    {
        workbooks->dynamicCall("Open(const QString&)", filePath);
        workbook = excel->querySubObject("ActiveWorkBook");
        sheets = workbook->querySubObject("WorkSheets");
    }
    QTime time;
    time.start();
    /************************时间安排算法***************************/

    for(int i=0;i<14;++i)avaliableTime[i].clear();
    setMessage("专业课考试安排中........");
    int numOfOrdinary = ordinary.size();
    for(int i = 0;i < numOfOrdinary; ++ i)
    {
        course tmpCourse = ordinary[i];
        int classNum = ordinary[i].Classes.size();
        for(int j = 0; j < 14; ++ j)
        {
            //有留学生的考试不安排在晚上(目前不需要了)
            //if(j % 2 == 0 && tmpCourse.internationalStuExist)continue;
            bool flag = true;//标识这一个时间段有没有学生安排重复
            for(int k = 0; k < classNum && flag; ++k)
            {
                Class tmpClass = classes[tmpCourse.Classes[k]];
                int numOfRoom = avaliableTime[j].size();
                int stuNum = tmpClass.stus.size();
                for(int l = 0; l < numOfRoom && flag; ++ l)
                {
                    int numOfClass = avaliableTime[j][l].arrangedClasses.size();
                    for(int m=0; m < numOfClass && flag; ++ m)
                    {
                        int tmpArrangedClass = avaliableTime[j][l].arrangedClasses[m];
                        int tmpStuNum = classes[tmpArrangedClass].stus.size();
                        for(int n = 0; n< stuNum && flag; ++n){
                            for(int o=0;o < tmpStuNum && flag; ++ o)
                            {
                                if(tmpClass.stus[n].stuId ==
                                   classes[tmpArrangedClass].stus[o].stuId)flag = false;
                            }
                        }
                    }
                }
            }
            //如果没有学生重复，则安排进这个时间段
            if(flag){
                int classesArranged = 0,
                    totClasses = classNum;
                int isClassArranged[classNum];
                for(int k = 0; k < classNum; ++k)isClassArranged[k] = 0;
                while(classesArranged < totClasses)
                {
                    examinationRoom tmpRoom;
                    for(int k = 0; k < classNum; ++k)
                    {
                        if(!isClassArranged[k] && tmpRoom.totStudents + classes[tmpCourse.Classes[k]].stus.size() <= maxRoomLimite){
                            tmpRoom.arrangedClasses.push_back(tmpCourse.Classes[k]);
                            tmpRoom.totStudents += classes[tmpCourse.Classes[k]].stus.size();
                            isClassArranged[k] = 1;
                            classesArranged++;
                        }
                    }
                    bool arrangedFlag = false;
                    for(int k = 0; k < avaliableTime[j].size() && !arrangedFlag; ++k)
                    {
                        if(tmpRoom.totStudents + avaliableTime[j][k].totStudents <= maxRoomLimite){
                            for(int l = 0; l < tmpRoom.arrangedClasses.size(); ++l)
                            {
                                avaliableTime[j][k].arrangedClasses.push_back(tmpRoom.arrangedClasses[l]);
                            }
                            avaliableTime[j][k].totStudents += tmpRoom.totStudents;
                            arrangedFlag = true;
                        }
                    }
                    if(!arrangedFlag)
                    {
                        avaliableTime[j].push_back(tmpRoom);
                    }
                }
                break;
            }
        }
    }

    //合并处理，把一些人少的教室和当天晚上人少的教室合并（没有学生在时间段内重复的话）
    for(int i=0;i<14;i+=2)
    {
        QMap<QString,int> stuExist;
        int numOfRN = avaliableTime[i+1].size();
        for(int j = numOfRN-1; j>=0; --j)
        {
            int numOfC = avaliableTime[i+1][j].arrangedClasses.size();
            for(int k = 0; k < numOfC; ++k)
            {
                int numOfS = classes[avaliableTime[i + 1][j].arrangedClasses[k]].stus.size();
                for(int l = 0; l<numOfS; ++l)stuExist[classes[avaliableTime[i + 1][j].arrangedClasses[k]].stus[l].stuId] = 1;
            }
        }
        int numOfRD = avaliableTime[i].size();
        for(int j = numOfRD - 1; j >= 0; -- j)
        {
            if(avaliableTime[i][j].totStudents <= 50)
            {
                int numOfC = avaliableTime[i][j].arrangedClasses.size();
                bool flag = false;
                for(int k = 0; k<numOfC && !flag; ++k)
                {
                    int numOfS = classes[avaliableTime[i][j].arrangedClasses[k]].stus.size();
                    for(int l = 0;l < numOfS && !flag; ++l)flag = stuExist[classes[avaliableTime[i][j].arrangedClasses[k]].stus[l].stuId];
                }
                if(!flag)
                {
                    for(int k = numOfRN-1; k >= 0; --k)
                    {
                        if(avaliableTime[i+1][k].totStudents + avaliableTime[i][j].totStudents <= maxRoomLimite)
                        {
                            for(int l = 0;l < numOfC && !flag; ++l)
                            {
                                avaliableTime[i+1][k].arrangedClasses.push_back(avaliableTime[i][j].arrangedClasses[l]);
                            }
                            avaliableTime[i+1][k].totStudents += avaliableTime[i][j].totStudents;
                            avaliableTime[i].erase(avaliableTime[i].begin()+j);
                            break;
                        }
                    }
                }
            }
        }
    }

    //sheets->querySubObject("Add()");
    QAxObject *sheet=sheets->querySubObject("Item(int)", 1);
    sheet->setProperty("Name", "专业课安排");
    sheet=sheets->querySubObject("Item(const QString&)", "专业课安排");
    QAxObject *range = sheet->querySubObject("Cells(int,int)", 1, 1);
    range->dynamicCall("SetValue(const QString&)", "时间");
    range = sheet->querySubObject("Cells(int,int)", 1, 2);
    range->dynamicCall("SetValue(const QString&)", "考场");
    range = sheet->querySubObject("Cells(int,int)", 1, 3);
    range->dynamicCall("SetValue(const QString&)", "考场人数");
    range = sheet->querySubObject("Cells(int,int)", 1, 4);
    range->dynamicCall("SetValue(const QString&)", "课程名称");
    range = sheet->querySubObject("Cells(int,int)", 1, 5);
    range->dynamicCall("SetValue(const QString&)", "原教学班");
    range = sheet->querySubObject("Cells(int,int)", 1, 6);
    range->dynamicCall("SetValue(const QString&)", "上课教师");
    range = sheet->querySubObject("Cells(int,int)", 1, 7);
    range->dynamicCall("SetValue(const QString&)", "教学班补考人数");
    range = sheet->querySubObject("Cells(int,int)", 1, 8);
    range->dynamicCall("SetValue(const QString&)", "学生姓名");
    range = sheet->querySubObject("Cells(int,int)", 1, 9);
    range->dynamicCall("SetValue(const QString&)", "学生学号");
    range = sheet->querySubObject("Cells(int,int)", 1, 10);
    range->dynamicCall("SetValue(const QString&)", "补考确认");
    range = sheet->querySubObject("Cells(int,int)", 1, 11);
    range->dynamicCall("SetValue(const QString&)", "学院");
    range = sheet->querySubObject("Cells(int,int)", 1, 12);
    range->dynamicCall("SetValue(const QString&)", "班级");
    range = sheet->querySubObject("Cells(int,int)", 1, 13);
    range->dynamicCall("SetValue(const QString&)", "考试性质");
    int tmp = 1;
    for(int i = 0; i < 14; ++i){
        if(avaliableTime[i].size()==0)continue;
        QString date="周";
        switch(i/2){
        case 0:
            date+="一";
            break;
        case 1:
            date+="二";
            break;
        case 2:
            date+="三";
            break;
        case 3:
            date+="四";
            break;
        case 4:
            date+="五";
            break;
        case 5:
            date+="六";
            break;
        case 6:
            date+="日";
            break;
        }
        if(i&1)date+="晚上";
        else date+="中午";
        int numOfRoom = avaliableTime[i].size();
        int dateStart = tmp + 1;
        for(int j = 0; j < numOfRoom; ++j){
            QString room = "教室" + QString::number(j+1);
            int roomStart = tmp + 1;
            int numOfC = avaliableTime[i][j].arrangedClasses.size();
            for(int k = 0; k < numOfC; ++k){
                int courseStart=tmp+1,lastSame=tmp+1;
                Class tmpClass = classes[avaliableTime[i][j].arrangedClasses[k]];
                int numOfStus= tmpClass.stus.size();
                for(int l=0;l<numOfStus;++l){
                    student tmpStu = tmpClass.stus[l];
                    range = sheet->querySubObject("Cells(int,int)",tmp+1,8);
                    range->dynamicCall("SetValue(const QString&)",tmpStu.name);
                    range = sheet->querySubObject("Cells(int,int)",tmp+1,9);
                    range->setProperty("NumberFormatLocal", "@");
                    range->dynamicCall("SetValue(const QString&)",tmpStu.stuId);
                    range = sheet->querySubObject("Cells(int,int)",tmp+1,10);
                    if(tmpStu.isChecked)range->dynamicCall("SetValue(const QString&)","是");
                    else range->dynamicCall("SetValue(const QString&)","否");
                    range = sheet->querySubObject("Cells(int,int)",tmp+1,11);
                    range->dynamicCall("SetValue(const QString&)",tmpStu.college);
                    range = sheet->querySubObject("Cells(int,int)",tmp+1,12);
                    range->dynamicCall("SetValue(const QString&)",tmpStu.professionalClass);
                    range = sheet->querySubObject("Cells(int,int)",tmp+1,13);
                    range->dynamicCall("SetValue(const QString&)",tmpStu.examProperty);
                    tmp++;
                    if(l != 0 && tmpStu.className != tmpClass.stus[l - 1].className)
                    {
                        QString cell=("E"+QString::number(lastSame)+":"+"E"+QString::number(tmp-1));
                        range = sheet->querySubObject("Range(const QString&)",cell);
                        range->setProperty("VerticalAlignment", -4108);//xlCenter
                        range->setProperty("WrapText", true);
                        range->setProperty("MergeCells", true);
                        range->dynamicCall("SetValue(const QString&)",tmpClass.stus[l-1].className);
                        cell=("F"+QString::number(lastSame)+":"+"F"+QString::number(tmp-1));
                        range = sheet->querySubObject("Range(const QString&)",cell);
                        range->setProperty("VerticalAlignment", -4108);//xlCenter
                        range->setProperty("WrapText", true);
                        range->setProperty("MergeCells", true);
                        range->dynamicCall("SetValue(const QString&)",tmpClass.teacher.split('/')[1]);
                        cell=("G"+QString::number(lastSame)+":"+"G"+QString::number(tmp-1));
                        range = sheet->querySubObject("Range(const QString&)",cell);
                        range->setProperty("VerticalAlignment", -4108);//xlCenter
                        range->setProperty("WrapText", true);
                        range->setProperty("MergeCells", true);
                        range->setProperty("NumberFormatLocal", "@");
                        range->dynamicCall("SetValue(const QString&)",QString::number(tmp-lastSame));
                        lastSame=tmp;
                    }
                }
                int courseEnd=tmp;
                QString cell=("E"+QString::number(lastSame)+":"+"E"+QString::number(courseEnd));
                range = sheet->querySubObject("Range(const QString&)",cell);
                range->setProperty("VerticalAlignment", -4108);//xlCenter
                range->setProperty("WrapText", true);
                range->setProperty("MergeCells", true);
                range->dynamicCall("SetValue(const QString&)",tmpClass.stus[numOfStus-1].className);
                cell=("F"+QString::number(lastSame)+":"+"F"+QString::number(courseEnd));
                range = sheet->querySubObject("Range(const QString&)",cell);
                range->setProperty("VerticalAlignment", -4108);//xlCenter
                range->setProperty("WrapText", true);
                range->setProperty("MergeCells", true);
                range->dynamicCall("SetValue(const QString&)",tmpClass.teacher.split('/')[1]);
                cell=("G"+QString::number(lastSame)+":"+"G"+QString::number(courseEnd));
                range = sheet->querySubObject("Range(const QString&)",cell);
                range->setProperty("VerticalAlignment", -4108);//xlCenter
                range->setProperty("WrapText", true);
                range->setProperty("MergeCells", true);
                range->setProperty("NumberFormatLocal", "@");
                range->dynamicCall("SetValue(const QString&)",QString::number(tmp-lastSame+1));
                cell=("D"+QString::number(courseStart)+":"+"D"+QString::number(courseEnd));
                range = sheet->querySubObject("Range(const QString&)",cell);
                range->setProperty("VerticalAlignment", -4108);//xlCenter
                range->setProperty("WrapText", true);
                range->setProperty("MergeCells", true);
                range->dynamicCall("SetValue(const QString&)",tmpClass.courseName);
            }
            int roomEnd=tmp;
            QString cell=("C"+QString::number(roomStart)+":"+"C"+QString::number(roomEnd));
            range = sheet->querySubObject("Range(const QString&)",cell);
            range->setProperty("VerticalAlignment", -4108);//xlCenter
            range->setProperty("WrapText", true);
            range->setProperty("MergeCells", true);
            range->dynamicCall("SetValue(const QString&)",QString::number(roomEnd-roomStart+1));
            cell=("B"+QString::number(roomStart)+":"+"B"+QString::number(roomEnd));
            range = sheet->querySubObject("Range(const QString&)",cell);
            range->setProperty("VerticalAlignment", -4108);//xlCenter
            range->setProperty("WrapText", true);
            range->dynamicCall("SetValue(const QString&)",room);
        }
        int dateEnd=tmp;
        QString cell=("A"+QString::number(dateStart)+":"+"A"+QString::number(dateEnd));
        range = sheet->querySubObject("Range(const QString&)",cell);
        range->setProperty("VerticalAlignment", -4108);//xlCenter
        range->setProperty("WrapText", true);
        range->dynamicCall("SetValue(const QString&)",date);
    }
    range = sheet->querySubObject("UsedRange");
    range->querySubObject("Columns")->dynamicCall("AutoFit");
    workbooks=excel->querySubObject("ActiveWindow");
    if (workbooks!=NULL){
        workbooks->setProperty("FreezePanes",false);//首先先把冻结窗格去掉，防止Excel已经设置的冻结窗格导致后面设置的不成功。
        workbooks->setProperty("SplitColumn", 0);//，第几列，目前设置的是Excel的第0列
        workbooks->setProperty("SplitRow", 1);//第几行,目前设置的是第1行
        workbooks->setProperty("FreezePanes", true);//设置冻结窗格属性为true
    }
    workbook->dynamicCall("Save()");
    setMessage("专业课安排完成!");


    setMessage("高等数学班考试安排中........");
    QVector<examinationRoom>rooms;
    int numOfMath = math.size();
    for(int i = 0;i < numOfMath; ++i){
        int classNum = math[i].Classes.size();
        int classesArranged = 0,
            totClasses = classNum;
        course tmpCourse = math[i];
        int isClassArranged[classNum];
        for(int j = 0; j < classNum; ++j)isClassArranged[j] = 0;
        while(classesArranged < totClasses)
        {
            examinationRoom tmpRoom;
            for(int j = 0; j < classNum; ++j)
            {
                if(!isClassArranged[j] && tmpRoom.totStudents + classes[tmpCourse.Classes[j]].stus.size() <= maxRoomLimite){
                    tmpRoom.arrangedClasses.push_back(tmpCourse.Classes[j]);
                    tmpRoom.totStudents += classes[tmpCourse.Classes[j]].stus.size();
                    isClassArranged[j] = 1;
                    classesArranged++;
                }
            }
            int numOfRoom = rooms.size();
            bool done = false;
            for(int k = 0; k < numOfRoom && !done; ++k){
                if(rooms[k].totStudents + tmpRoom.totStudents <= maxRoomLimite){
                    int numOfCourse = rooms[k].arrangedClasses.size();
                    bool flag = tmpCourse.courseName.contains("英文");
                    for(int j = 0; j < numOfCourse && flag == tmpCourse.courseName.contains("英文"); ++j)flag = classes[rooms[k].arrangedClasses[j]].courseName.contains("英文");
                    if(flag == tmpCourse.courseName.contains("英文")){
                        for(int l = 0; l < tmpRoom.arrangedClasses.size(); ++l)rooms[k].arrangedClasses.push_back(tmpRoom.arrangedClasses[l]);
                        rooms[k].totStudents += tmpRoom.totStudents;
                        done=true;
                    }
                }
            }
            if(!done){
                rooms.push_back(tmpRoom);
            }
        }
    }
    sheets->querySubObject("Add()");
    sheet=sheets->querySubObject("Item(int)",1);
    sheet->setProperty("Name","高等数学安排");
    sheet=sheets->querySubObject("Item(const QString&)", "高等数学安排");
    range = sheet->querySubObject("Cells(int,int)", 1, 1);
    range->dynamicCall("SetValue(const QString&)", "时间");
    range = sheet->querySubObject("Cells(int,int)", 1, 2);
    range->dynamicCall("SetValue(const QString&)", "考场");
    range = sheet->querySubObject("Cells(int,int)", 1, 3);
    range->dynamicCall("SetValue(const QString&)", "考场人数");
    range = sheet->querySubObject("Cells(int,int)", 1, 4);
    range->dynamicCall("SetValue(const QString&)", "课程名称");
    range = sheet->querySubObject("Cells(int,int)", 1, 5);
    range->dynamicCall("SetValue(const QString&)", "原教学班");
    range = sheet->querySubObject("Cells(int,int)", 1, 6);
    range->dynamicCall("SetValue(const QString&)", "上课教师");
    range = sheet->querySubObject("Cells(int,int)", 1, 7);
    range->dynamicCall("SetValue(const QString&)", "教学班补考人数");
    range = sheet->querySubObject("Cells(int,int)", 1, 8);
    range->dynamicCall("SetValue(const QString&)", "学生姓名");
    range = sheet->querySubObject("Cells(int,int)", 1, 9);
    range->dynamicCall("SetValue(const QString&)", "学生学号");
    range = sheet->querySubObject("Cells(int,int)", 1, 10);
    range->dynamicCall("SetValue(const QString&)", "补考确认");
    range = sheet->querySubObject("Cells(int,int)", 1, 11);
    range->dynamicCall("SetValue(const QString&)", "学院");
    range = sheet->querySubObject("Cells(int,int)", 1, 12);
    range->dynamicCall("SetValue(const QString&)", "班级");
    range = sheet->querySubObject("Cells(int,int)", 1, 13);
    range->dynamicCall("SetValue(const QString&)", "考试性质");
    tmp = 1;
    QString date = "周六下午";
    int numOfRoom = rooms.size();
    for(int j = 0;j < numOfRoom; ++j){
        QString room = "教室"+QString::number(j + 1);
        int roomStart=tmp+1;
        int numOfC = rooms[j].arrangedClasses.size();
        for(int k = 0;k < numOfC; ++k){
            int courseStart = tmp+1, lastSame = tmp+1;
            Class tmpClass = classes[rooms[j].arrangedClasses[k]];
            int numOfStus = tmpClass.stus.size();
            for(int l = 0;l < numOfStus; ++l){
                student tmpStu = tmpClass.stus[l];
                range = sheet->querySubObject("Cells(int,int)",tmp+1,8);
                range->dynamicCall("SetValue(const QString&)",tmpStu.name);
                range = sheet->querySubObject("Cells(int,int)",tmp+1,9);
                range->setProperty("NumberFormatLocal", "@");
                range->dynamicCall("SetValue(const QString&)",tmpStu.stuId);
                range = sheet->querySubObject("Cells(int,int)",tmp+1,10);
                if(tmpStu.isChecked)range->dynamicCall("SetValue(const QString&)","是");
                else range->dynamicCall("SetValue(const QString&)","否");
                range = sheet->querySubObject("Cells(int,int)",tmp+1,11);
                range->dynamicCall("SetValue(const QString&)",tmpStu.college);
                range = sheet->querySubObject("Cells(int,int)",tmp+1,12);
                range->dynamicCall("SetValue(const QString&)",tmpStu.professionalClass);
                range = sheet->querySubObject("Cells(int,int)",tmp+1,13);
                range->dynamicCall("SetValue(const QString&)",tmpStu.examProperty);
                tmp++;
                if(l!=0 && tmpStu.className != tmpClass.stus[l-1].className)
                {
                    QString cell=("E"+QString::number(lastSame)+":"+"E"+QString::number(tmp-1));
                    range = sheet->querySubObject("Range(const QString&)",cell);
                    range->setProperty("VerticalAlignment", -4108);//xlCenter
                    range->setProperty("WrapText", true);
                    range->setProperty("MergeCells", true);
                    range->dynamicCall("SetValue(const QString&)",tmpClass.stus[l-1].className);
                    cell=("F"+QString::number(lastSame)+":"+"F"+QString::number(tmp-1));
                    range = sheet->querySubObject("Range(const QString&)",cell);
                    range->setProperty("VerticalAlignment", -4108);//xlCenter
                    range->setProperty("WrapText", true);
                    range->setProperty("MergeCells", true);
                    range->dynamicCall("SetValue(const QString&)",tmpClass.teacher.split('/')[1]);
                    cell=("G"+QString::number(lastSame)+":"+"G"+QString::number(tmp-1));
                    range = sheet->querySubObject("Range(const QString&)",cell);
                    range->setProperty("VerticalAlignment", -4108);//xlCenter
                    range->setProperty("WrapText", true);
                    range->setProperty("MergeCells", true);
                    range->setProperty("NumberFormatLocal", "@");
                    range->dynamicCall("SetValue(const QString&)",QString::number(tmp-lastSame));
                    lastSame=tmp;
                }
            }
            int courseEnd=tmp;
            QString cell=("D"+QString::number(courseStart)+":"+"D"+QString::number(courseEnd));
            range = sheet->querySubObject("Range(const QString&)",cell);
            range->setProperty("VerticalAlignment", -4108);//xlCenter
            range->setProperty("WrapText", true);
            range->setProperty("MergeCells", true);
            range->dynamicCall("SetValue(const QString&)",tmpClass.courseName);
            cell=("E"+QString::number(lastSame)+":"+"E"+QString::number(courseEnd));
            range = sheet->querySubObject("Range(const QString&)",cell);
            range->setProperty("VerticalAlignment", -4108);//xlCenter
            range->setProperty("WrapText", true);
            range->setProperty("MergeCells", true);
            range->dynamicCall("SetValue(const QString&)",tmpClass.stus[numOfStus-1].className);
            cell=("F"+QString::number(lastSame)+":"+"F"+QString::number(courseEnd));
            range = sheet->querySubObject("Range(const QString&)",cell);
            range->setProperty("VerticalAlignment", -4108);//xlCenter
            range->setProperty("WrapText", true);
            range->setProperty("MergeCells", true);
            range->dynamicCall("SetValue(const QString&)",tmpClass.teacher.split('/')[1]);
            cell=("G"+QString::number(lastSame)+":"+"G"+QString::number(courseEnd));
            range = sheet->querySubObject("Range(const QString&)",cell);
            range->setProperty("VerticalAlignment", -4108);//xlCenter
            range->setProperty("WrapText", true);
            range->setProperty("MergeCells", true);
            range->setProperty("NumberFormatLocal", "@");
            range->dynamicCall("SetValue(const QString&)",QString::number(tmp-lastSame+1));
        }
        int roomEnd=tmp;
        QString cell=("C"+QString::number(roomStart)+":"+"C"+QString::number(roomEnd));
        range = sheet->querySubObject("Range(const QString&)",cell);
        range->setProperty("VerticalAlignment", -4108);//xlCenter
        range->setProperty("WrapText", true);
        range->setProperty("MergeCells", true);
        range->dynamicCall("SetValue(const QString&)",QString::number(roomEnd-roomStart+1));
        cell=("B"+QString::number(roomStart)+":"+"B"+QString::number(roomEnd));
        range = sheet->querySubObject("Range(const QString&)",cell);
        range->setProperty("VerticalAlignment", -4108);//xlCenter
        range->setProperty("WrapText", true);
        range->dynamicCall("SetValue(const QString&)",room);
    }
    QString cell=("A"+QString::number(1)+":"+"A"+QString::number(tmp));
    range = sheet->querySubObject("Range(const QString&)",cell);
    range->setProperty("VerticalAlignment", -4108);//xlCenter
    range->setProperty("WrapText", true);
    range->dynamicCall("SetValue(const QString&)",date);
    range = sheet->querySubObject("UsedRange");
    range->querySubObject("Columns")->dynamicCall("AutoFit");
    workbooks=excel->querySubObject("ActiveWindow");
    if (workbooks!=NULL){
        workbooks->setProperty("FreezePanes",false);//首先先把冻结窗格去掉，防止Excel已经设置的冻结窗格导致后面设置的不成功。
        workbooks->setProperty("SplitColumn", 0);//，第几列，目前设置的是Excel的第0列
        workbooks->setProperty("SplitRow", 1);//第几行,目前设置的是第1行
        workbooks->setProperty("FreezePanes", true);//设置冻结窗格属性为true
    }
    workbook->dynamicCall("Save()");
    setMessage("高等数学安排完成!");


    setMessage("计算机应用A考试安排中........");
    rooms.clear();
    int numOfComputer=computerA.size();
    for(int i = 0; i < numOfComputer; ++i){
        course tmpCourse = computerA[i];
        int classNum = tmpCourse.Classes.size();
        int classesArranged = 0,
            totClasses = classNum;
        int isClassArranged[classNum];
        for(int j = 0; j < classNum; ++j)isClassArranged[j] = 0;
        while(classesArranged < totClasses)
        {
            examinationRoom tmpRoom;
            for(int j = 0; j < classNum; ++j)
            {
                if(!isClassArranged[j] && tmpRoom.totStudents + classes[tmpCourse.Classes[j]].stus.size() <= maxRoomLimite){
                    tmpRoom.arrangedClasses.push_back(tmpCourse.Classes[j]);
                    tmpRoom.totStudents += classes[tmpCourse.Classes[j]].stus.size();
                    isClassArranged[j] = 1;
                    classesArranged++;
                }
            }
            int numOfRoom = rooms.size();
            bool done = false;
            for(int k = 0; k < numOfRoom && !done; ++k){
                if(rooms[k].totStudents + tmpRoom.totStudents <= maxRoomLimite){
                    for(int l = 0; l < tmpRoom.arrangedClasses.size(); ++l)rooms[k].arrangedClasses.push_back(tmpRoom.arrangedClasses[l]);
                    rooms[k].totStudents += tmpRoom.totStudents;
                    done=true;
                }
            }
            if(!done){
                rooms.push_back(tmpRoom);
            }
        }
    }
    sheets->querySubObject("Add()");
    sheet=sheets->querySubObject("Item(int)",1);
    sheet->setProperty("Name","计算机应用A安排");
    sheet=sheets->querySubObject("Item(const QString&)", "计算机应用A安排");
    range = sheet->querySubObject("Cells(int,int)", 1, 1);
    range->dynamicCall("SetValue(const QString&)", "时间");
    range = sheet->querySubObject("Cells(int,int)", 1, 2);
    range->dynamicCall("SetValue(const QString&)", "考场");
    range = sheet->querySubObject("Cells(int,int)", 1, 3);
    range->dynamicCall("SetValue(const QString&)", "考场人数");
    range = sheet->querySubObject("Cells(int,int)", 1, 4);
    range->dynamicCall("SetValue(const QString&)", "课程名称");
    range = sheet->querySubObject("Cells(int,int)", 1, 5);
    range->dynamicCall("SetValue(const QString&)", "原教学班");
    range = sheet->querySubObject("Cells(int,int)", 1, 6);
    range->dynamicCall("SetValue(const QString&)", "上课教师");
    range = sheet->querySubObject("Cells(int,int)", 1, 7);
    range->dynamicCall("SetValue(const QString&)", "教学班补考人数");
    range = sheet->querySubObject("Cells(int,int)", 1, 8);
    range->dynamicCall("SetValue(const QString&)", "学生姓名");
    range = sheet->querySubObject("Cells(int,int)", 1, 9);
    range->dynamicCall("SetValue(const QString&)", "学生学号");
    range = sheet->querySubObject("Cells(int,int)", 1, 10);
    range->dynamicCall("SetValue(const QString&)", "补考确认");
    range = sheet->querySubObject("Cells(int,int)", 1, 11);
    range->dynamicCall("SetValue(const QString&)", "学院");
    range = sheet->querySubObject("Cells(int,int)", 1, 12);
    range->dynamicCall("SetValue(const QString&)", "班级");
    range = sheet->querySubObject("Cells(int,int)", 1, 13);
    range->dynamicCall("SetValue(const QString&)", "考试性质");
    tmp=1;
    date="周六下午";

    numOfRoom=rooms.size();
    for(int j=0;j<numOfRoom;++j){
        QString room="教室"+QString::number(j+1);
        int roomStart = tmp+1;
        int numOfC = rooms[j].arrangedClasses.size();
        for(int k = 0; k < numOfC; ++k){
            int courseStart = tmp+1,lastSame = tmp+1;
            Class tmpClass = classes[rooms[j].arrangedClasses[k]];
            int numOfStus = tmpClass.stus.size();
            for(int l = 0; l < numOfStus; ++l){
                student tmpStu=tmpClass.stus[l];
                range = sheet->querySubObject("Cells(int,int)",tmp+1,8);
                range->dynamicCall("SetValue(const QString&)",tmpStu.name);
                range = sheet->querySubObject("Cells(int,int)",tmp+1,9);
                range->setProperty("NumberFormatLocal", "@");
                range->dynamicCall("SetValue(const QString&)",tmpStu.stuId);
                range = sheet->querySubObject("Cells(int,int)",tmp+1,10);
                if(tmpStu.isChecked)range->dynamicCall("SetValue(const QString&)","是");
                else range->dynamicCall("SetValue(const QString&)","否");
                range = sheet->querySubObject("Cells(int,int)",tmp+1,11);
                range->dynamicCall("SetValue(const QString&)",tmpStu.college);
                range = sheet->querySubObject("Cells(int,int)",tmp+1,12);
                range->dynamicCall("SetValue(const QString&)",tmpStu.professionalClass);
                range = sheet->querySubObject("Cells(int,int)",tmp+1,13);
                range->dynamicCall("SetValue(const QString&)",tmpStu.examProperty);
                tmp++;
                if(l != 0 && tmpStu.className != tmpClass.stus[l - 1].className)
                {
                    QString cell = ("E"+QString::number(lastSame)+":"+"E"+QString::number(tmp-1));
                    range = sheet->querySubObject("Range(const QString&)",cell);
                    range->setProperty("VerticalAlignment", -4108);//xlCenter
                    range->setProperty("WrapText", true);
                    range->setProperty("MergeCells", true);
                    range->dynamicCall("SetValue(const QString&)",tmpClass.stus[l-1].className);
                    cell=("F"+QString::number(lastSame)+":"+"F"+QString::number(tmp-1));
                    range = sheet->querySubObject("Range(const QString&)",cell);
                    range->setProperty("VerticalAlignment", -4108);//xlCenter
                    range->setProperty("WrapText", true);
                    range->setProperty("MergeCells", true);
                    range->dynamicCall("SetValue(const QString&)",tmpClass.teacher.split('/')[1]);
                    cell=("G"+QString::number(lastSame)+":"+"G"+QString::number(tmp-1));
                    range = sheet->querySubObject("Range(const QString&)",cell);
                    range->setProperty("VerticalAlignment", -4108);//xlCenter
                    range->setProperty("WrapText", true);
                    range->setProperty("MergeCells", true);
                    range->setProperty("NumberFormatLocal", "@");
                    range->dynamicCall("SetValue(const QString&)",QString::number(tmp-lastSame));
                    lastSame=tmp;
                }
            }
            int courseEnd=tmp;
            QString cell=("D"+QString::number(courseStart)+":"+"D"+QString::number(courseEnd));
            range = sheet->querySubObject("Range(const QString&)",cell);
            range->setProperty("VerticalAlignment", -4108);//xlCenter
            range->setProperty("WrapText", true);
            range->setProperty("MergeCells", true);
            range->dynamicCall("SetValue(const QString&)",tmpClass.courseName);
            cell=("E"+QString::number(lastSame)+":"+"E"+QString::number(courseEnd));
            range = sheet->querySubObject("Range(const QString&)",cell);
            range->setProperty("VerticalAlignment", -4108);//xlCenter
            range->setProperty("WrapText", true);
            range->setProperty("MergeCells", true);
            range->dynamicCall("SetValue(const QString&)",tmpClass.stus[numOfStus-1].className);
            cell=("F"+QString::number(lastSame)+":"+"F"+QString::number(courseEnd));
            range = sheet->querySubObject("Range(const QString&)",cell);
            range->setProperty("VerticalAlignment", -4108);//xlCenter
            range->setProperty("WrapText", true);
            range->setProperty("MergeCells", true);
            range->dynamicCall("SetValue(const QString&)",tmpClass.teacher.split('/')[1]);
            cell=("G"+QString::number(lastSame)+":"+"G"+QString::number(courseEnd));
            range = sheet->querySubObject("Range(const QString&)",cell);
            range->setProperty("VerticalAlignment", -4108);//xlCenter
            range->setProperty("WrapText", true);
            range->setProperty("MergeCells", true);
            range->setProperty("NumberFormatLocal", "@");
            range->dynamicCall("SetValue(const QString&)",QString::number(tmp-lastSame+1));
        }
        int roomEnd=tmp;
        QString cell=("C"+QString::number(roomStart)+":"+"C"+QString::number(roomEnd));
        range = sheet->querySubObject("Range(const QString&)",cell);
        range->setProperty("VerticalAlignment", -4108);//xlCenter
        range->setProperty("WrapText", true);
        range->setProperty("MergeCells", true);
        range->dynamicCall("SetValue(const QString&)",QString::number(roomEnd-roomStart+1));
        cell=("B"+QString::number(roomStart)+":"+"B"+QString::number(roomEnd));
        range = sheet->querySubObject("Range(const QString&)",cell);
        range->setProperty("VerticalAlignment", -4108);//xlCenter
        range->setProperty("WrapText", true);
        range->dynamicCall("SetValue(const QString&)",room);
    }
    cell=("A"+QString::number(1)+":"+"A"+QString::number(tmp));
    range = sheet->querySubObject("Range(const QString&)",cell);
    range->setProperty("VerticalAlignment", -4108);//xlCenter
    range->setProperty("WrapText", true);
    range->dynamicCall("SetValue(const QString&)",date);
    range = sheet->querySubObject("UsedRange");
    range->querySubObject("Columns")->dynamicCall("AutoFit");
    workbooks=excel->querySubObject("ActiveWindow");
    if (workbooks!=NULL){
        workbooks->setProperty("FreezePanes",false);//首先先把冻结窗格去掉，防止Excel已经设置的冻结窗格导致后面设置的不成功。
        workbooks->setProperty("SplitColumn", 0);//，第几列，目前设置的是Excel的第0列
        workbooks->setProperty("SplitRow", 1);//第几行,目前设置的是第1行
        workbooks->setProperty("FreezePanes", true);//设置冻结窗格属性为true
    }
    workbook->dynamicCall("Save()");
    setMessage("计算机应用A安排完成!");

    setMessage("计算机应用B考试安排中........");
    rooms.clear();
    numOfComputer = computerB.size();
    for(int i = 0; i < numOfComputer; ++i){
        course tmpCourse = computerB[i];
        int classNum = tmpCourse.Classes.size();
        int classesArranged = 0,
            totClasses = classNum;
        int isClassArranged[classNum];
        for(int j = 0; j < classNum; ++j)isClassArranged[j] = 0;
        while(classesArranged < totClasses)
        {
            examinationRoom tmpRoom;
            for(int j = 0; j < classNum; ++j)
            {
                if(!isClassArranged[j] && tmpRoom.totStudents + classes[tmpCourse.Classes[j]].stus.size() <= maxRoomLimite){
                    tmpRoom.arrangedClasses.push_back(tmpCourse.Classes[j]);
                    tmpRoom.totStudents += classes[tmpCourse.Classes[j]].stus.size();
                    isClassArranged[j] = 1;
                    classesArranged++;
                }
            }
            int numOfRoom = rooms.size();
            bool done = false;
            for(int k = 0; k < numOfRoom && !done; ++k){
                if(rooms[k].totStudents + tmpRoom.totStudents <= maxRoomLimite){
                    for(int l = 0; l < tmpRoom.arrangedClasses.size(); ++l)rooms[k].arrangedClasses.push_back(tmpRoom.arrangedClasses[l]);
                    rooms[k].totStudents += tmpRoom.totStudents;
                    done=true;
                }
            }
            if(!done){
                rooms.push_back(tmpRoom);
            }
        }
    }
    sheets->querySubObject("Add()");
    sheet=sheets->querySubObject("Item(int)",1);
    sheet->setProperty("Name","计算机应用B安排");
    sheet=sheets->querySubObject("Item(const QString&)", "计算机应用B安排");
    range = sheet->querySubObject("Cells(int,int)", 1, 1);
    range->dynamicCall("SetValue(const QString&)", "时间");
    range = sheet->querySubObject("Cells(int,int)", 1, 2);
    range->dynamicCall("SetValue(const QString&)", "考场");
    range = sheet->querySubObject("Cells(int,int)", 1, 3);
    range->dynamicCall("SetValue(const QString&)", "考场人数");
    range = sheet->querySubObject("Cells(int,int)", 1, 4);
    range->dynamicCall("SetValue(const QString&)", "课程名称");
    range = sheet->querySubObject("Cells(int,int)", 1, 5);
    range->dynamicCall("SetValue(const QString&)", "原教学班");
    range = sheet->querySubObject("Cells(int,int)", 1, 6);
    range->dynamicCall("SetValue(const QString&)", "上课教师");
    range = sheet->querySubObject("Cells(int,int)", 1, 7);
    range->dynamicCall("SetValue(const QString&)", "教学班补考人数");
    range = sheet->querySubObject("Cells(int,int)", 1, 8);
    range->dynamicCall("SetValue(const QString&)", "学生姓名");
    range = sheet->querySubObject("Cells(int,int)", 1, 9);
    range->dynamicCall("SetValue(const QString&)", "学生学号");
    range = sheet->querySubObject("Cells(int,int)", 1, 10);
    range->dynamicCall("SetValue(const QString&)", "补考确认");
    range = sheet->querySubObject("Cells(int,int)", 1, 11);
    range->dynamicCall("SetValue(const QString&)", "学院");
    range = sheet->querySubObject("Cells(int,int)", 1, 12);
    range->dynamicCall("SetValue(const QString&)", "班级");
    range = sheet->querySubObject("Cells(int,int)", 1, 13);
    range->dynamicCall("SetValue(const QString&)", "考试性质");
    tmp=1;
    date="周六下午";

    numOfRoom=rooms.size();
    //dateStart=tmp+1;
    for(int j = 0;j < numOfRoom; ++j){
        QString room = "教室"+QString::number(j+1);
        int roomStart = tmp+1;
        int numOfC = rooms[j].arrangedClasses.size();
        for(int k = 0;k < numOfC; ++k){
            int courseStart = tmp + 1,lastSame = tmp + 1;
            Class tmpClass = classes[rooms[j].arrangedClasses[k]];
            int numOfStus = tmpClass.stus.size();
            for(int l = 0; l < numOfStus; ++l){
                student tmpStu=tmpClass.stus[l];
                range = sheet->querySubObject("Cells(int,int)",tmp+1,8);
                range->dynamicCall("SetValue(const QString&)",tmpStu.name);
                range = sheet->querySubObject("Cells(int,int)",tmp+1,9);
                range->setProperty("NumberFormatLocal", "@");
                range->dynamicCall("SetValue(const QString&)",tmpStu.stuId);
                range = sheet->querySubObject("Cells(int,int)",tmp+1,10);
                if(tmpStu.isChecked)range->dynamicCall("SetValue(const QString&)","是");
                else range->dynamicCall("SetValue(const QString&)","否");
                range = sheet->querySubObject("Cells(int,int)",tmp+1,11);
                range->dynamicCall("SetValue(const QString&)",tmpStu.college);
                range = sheet->querySubObject("Cells(int,int)",tmp+1,12);
                range->dynamicCall("SetValue(const QString&)",tmpStu.professionalClass);
                range = sheet->querySubObject("Cells(int,int)",tmp+1,13);
                range->dynamicCall("SetValue(const QString&)",tmpStu.examProperty);
                tmp++;
                if(l != 0 && tmpStu.className != tmpClass.stus[l-1].className)
                {
                    QString cell=("E"+QString::number(lastSame)+":"+"E"+QString::number(tmp-1));
                    range = sheet->querySubObject("Range(const QString&)",cell);
                    range->setProperty("VerticalAlignment", -4108);//xlCenter
                    range->setProperty("WrapText", true);
                    range->setProperty("MergeCells", true);
                    range->dynamicCall("SetValue(const QString&)",tmpClass.stus[l-1].className);
                    cell=("F"+QString::number(lastSame)+":"+"F"+QString::number(tmp-1));
                    range = sheet->querySubObject("Range(const QString&)",cell);
                    range->setProperty("VerticalAlignment", -4108);//xlCenter
                    range->setProperty("WrapText", true);
                    range->setProperty("MergeCells", true);
                    range->dynamicCall("SetValue(const QString&)",tmpClass.teacher.split('/')[1]);
                    cell=("G"+QString::number(lastSame)+":"+"G"+QString::number(tmp-1));
                    range = sheet->querySubObject("Range(const QString&)",cell);
                    range->setProperty("VerticalAlignment", -4108);//xlCenter
                    range->setProperty("WrapText", true);
                    range->setProperty("MergeCells", true);
                    range->setProperty("NumberFormatLocal", "@");
                    range->dynamicCall("SetValue(const QString&)",QString::number(tmp-lastSame));
                    lastSame=tmp;
                }
            }
            int courseEnd=tmp;
            QString cell=("D"+QString::number(courseStart)+":"+"D"+QString::number(courseEnd));
            range = sheet->querySubObject("Range(const QString&)",cell);
            range->setProperty("VerticalAlignment", -4108);//xlCenter
            range->setProperty("WrapText", true);
            range->setProperty("MergeCells", true);
            range->dynamicCall("SetValue(const QString&)",tmpClass.courseName);
            cell=("E"+QString::number(lastSame)+":"+"E"+QString::number(courseEnd));
            range = sheet->querySubObject("Range(const QString&)",cell);
            range->setProperty("VerticalAlignment", -4108);//xlCenter
            range->setProperty("WrapText", true);
            range->setProperty("MergeCells", true);
            range->dynamicCall("SetValue(const QString&)",tmpClass.stus[numOfStus-1].className);
            cell=("F"+QString::number(lastSame)+":"+"F"+QString::number(courseEnd));
            range = sheet->querySubObject("Range(const QString&)",cell);
            range->setProperty("VerticalAlignment", -4108);//xlCenter
            range->setProperty("WrapText", true);
            range->setProperty("MergeCells", true);
            range->dynamicCall("SetValue(const QString&)",tmpClass.teacher.split('/')[1]);
            cell=("G"+QString::number(lastSame)+":"+"G"+QString::number(courseEnd));
            range = sheet->querySubObject("Range(const QString&)",cell);
            range->setProperty("VerticalAlignment", -4108);//xlCenter
            range->setProperty("WrapText", true);
            range->setProperty("MergeCells", true);
            range->setProperty("NumberFormatLocal", "@");
            range->dynamicCall("SetValue(const QString&)",QString::number(tmp-lastSame+1));
        }
        int roomEnd=tmp;
        QString cell=("C"+QString::number(roomStart)+":"+"C"+QString::number(roomEnd));
        range = sheet->querySubObject("Range(const QString&)",cell);
        range->setProperty("VerticalAlignment", -4108);//xlCenter
        range->setProperty("WrapText", true);
        range->setProperty("MergeCells", true);
        range->dynamicCall("SetValue(const QString&)",QString::number(roomEnd-roomStart+1));
        cell=("B"+QString::number(roomStart)+":"+"B"+QString::number(roomEnd));
        range = sheet->querySubObject("Range(const QString&)",cell);
        range->setProperty("VerticalAlignment", -4108);//xlCenter
        range->setProperty("WrapText", true);
        range->dynamicCall("SetValue(const QString&)",room);
    }
    cell=("A"+QString::number(1)+":"+"A"+QString::number(tmp));
    range = sheet->querySubObject("Range(const QString&)",cell);
    range->setProperty("VerticalAlignment", -4108);//xlCenter
    range->setProperty("WrapText", true);
    range->dynamicCall("SetValue(const QString&)",date);
    range = sheet->querySubObject("UsedRange");
    range->querySubObject("Columns")->dynamicCall("AutoFit");
    workbooks=excel->querySubObject("ActiveWindow");
    if (workbooks!=NULL){
        workbooks->setProperty("FreezePanes",false);//首先先把冻结窗格去掉，防止Excel已经设置的冻结窗格导致后面设置的不成功。
        workbooks->setProperty("SplitColumn", 0);//，第几列，目前设置的是Excel的第0列
        workbooks->setProperty("SplitRow", 1);//第几行,目前设置的是第1行
        workbooks->setProperty("FreezePanes", true);//设置冻结窗格属性为true
    }
    workbook->dynamicCall("Save()");
    setMessage("计算机应用B安排完成!");


    excel->dynamicCall("Quit()");
    delete range;
    delete sheet;
    delete sheets;
    delete workbook;
    delete workbooks;
    delete excel;
    setMessage("考试安排完成!");

    qDebug()<<"课程总数:"<<courses.size()<<endl;
    qDebug()<<"专业课总数:"<<ordinary.size()<<endl;
    qDebug()<<"高等数学总数:"<<math.size()<<endl;
    qDebug()<<"计算机应用A总数:"<<computerA.size()<<endl;
    qDebug()<<"计算机应用B总数:"<<computerB.size()<<endl;
    qDebug()<<"消耗时间:"<<time.elapsed()/1000.0<<endl;
}
