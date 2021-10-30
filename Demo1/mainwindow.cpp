#include "mainwindow.h"
#include "ui_mainwindow.h"
#include<QFileDialog>
#include <QAxObject>
#include<QFileDialog>
#include<QDebug>
#include<QTime>
#include<QMessageBox>
#include<algorithm>
struct students{
    QString stuId=NULL,name=NULL,courseName=NULL,courseId=NULL,teacher=NULL,college=NULL,professionalClass=NULL;
    bool isChecked;
};
struct course{
    QString courseName=NULL;
    bool chineseStuExist,internationalStuExist;
    QVector<students>stus;
};
bool stuCmp(students A,students B)
{
    if(A.courseName!=B.courseName)return A.courseName<B.courseName;
    else return  A.stuId<B.stuId;
}
QVector<course>courses;
QMap<QString,int>courseExist;
QVector<course>ordinary,math,computerA,computerB;
struct examinationRoom{
    QVector<course>arrangedCourse;
    int totStudents=0;
    int start=0,end=0;
};
int maxRoomLimite;
bool roomCmp(examinationRoom A,examinationRoom B){return A.totStudents>B.totStudents;}
QVector<examinationRoom>avaliableTime[14];

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    ItemModel=new QStandardItemModel(this);
    ui->message->setWordWrap(true);
    setMessage("未选中任何文件(确保选择的文件请有姓名、学号、原教学班、课程代码、上课教师、补考确认这几项信息且信息都在sheet1中)");
}

MainWindow::~MainWindow(){
    delete ui;
}
void MainWindow::setMessage(QString message)
{
    ItemModel->appendRow(new QStandardItem(message));
    ui->message->setModel(ItemModel);
}
void MainWindow::on_chooseFile_clicked(){
    fileName=QFileDialog::getOpenFileName(this,tr("选择文件"),"",tr("Excel文件(*.xls *.xlsx)"));
    maxRoomLimite=90;
    if(fileName.isNull()){return;}
    courses.clear();
    courseExist.clear();
    math.clear();
    ordinary.clear();
    computerA.clear();
    computerB.clear();
    QAxObject *excel = new QAxObject(this);    //连接Excel控件
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
    const int rowCount = varRows.size();
    QVariantList names = varRows[0].toList();
    const int colCount=names.size();
    int stuId=-1,stuName=-1,check=-1,college=-1,professionalClass=-1;//学生学号、姓名、性别、原班级、是否确认;
    int courseId=-1,courseCName=-1,courseName=-1,teacher=-1;//课程代码、课程名称、课程详细名称、上课教师;
    for(int i=0;i<colCount;++i){
        if(names[i].toString()=="学号")stuId=i;
        if(names[i].toString()=="姓名")stuName=i;
        if(names[i].toString()=="补考确认")check=i;
        if(names[i].toString()=="课程名称")courseCName=i;
        if(names[i].toString()=="原教学班")courseName=i;
        if(names[i].toString()=="课程代码")courseId=i;
        if(names[i].toString()=="上课教师")teacher=i;
        if(names[i].toString()=="学院")college=i;
        if(names[i].toString()=="班级")professionalClass=i;
    }
    qDebug()<<"23333"<<endl;
    if(stuId==-1||stuName==-1||check==-1||courseId==-1||courseName==-1||teacher==-1){
        QString x="文件缺少:\n";
        if(stuId==-1)x+="学号\n";
        if(stuName==-1)x+="姓名\n";
        if(check==-1)x+="补考确认\n";
        if(college==-1)x+="学院\n";
        if(professionalClass==-1)x+="班级\n";
        if(courseId==-1)x+="课程名称\n";
        if(courseCName==-1)x+="原教学班\n";
        if(courseName==-1)x+="原教学班\n";
        if(teacher==-1)x+="上课教师\n";
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
        students tmpStu;
        int tmpCourse;
        tmpStu.stuId=rowData[stuId].toString();
        tmpStu.name=rowData[stuName].toString();
        tmpStu.isChecked=rowData[check].toString()=="是";
        tmpStu.college=rowData[college].toString();
        tmpStu.professionalClass=rowData[professionalClass].toString();
        tmpStu.courseName=rowData[courseName].toString();
        tmpStu.teacher=rowData[teacher].toString();
        tmpStu.courseId=rowData[courseId].toString();
        QString CName;
        CName=rowData[courseCName].toString();
        if(courseExist[CName]==0){
            course tmpC;
            tmpC.courseName=CName;
            tmpC.chineseStuExist=false;
            tmpC.internationalStuExist=false;
            courses.push_back(tmpC);
            courseExist[CName]=courses.size();
        }
        tmpCourse=courseExist[CName]-1;
        courses[tmpCourse].stus.push_back(tmpStu);
        int length=tmpStu.name.size();
        for(int k=0;k<length;++k){
            QChar cha=tmpStu.name.at(k);
            ushort uni=cha.unicode();
            if((uni>='a'&&uni<='z')||(uni>='A'&&uni<='Z')){
                courses[tmpCourse].internationalStuExist=true;
                break;
            }
            if(uni>=0x4E00&&uni<=0x9FA5){
                courses[tmpCourse].chineseStuExist=true;
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
    int numOfCourses=courses.size();
    for(int i=0;i<numOfCourses;++i){
        if(!courses[i].stus.size())continue;
        qSort(courses[i].stus.begin(),courses[i].stus.end(),stuCmp);
        if(courses[i].courseName.contains("高等数学"))math.push_back(courses[i]);
        else if(courses[i].courseName.contains("计算机应用A"))computerA.push_back(courses[i]);
        else if(courses[i].courseName.contains("计算机应用B"))computerB.push_back(courses[i]);
        else ordinary.push_back(courses[i]);
    }
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
    int numOfOrdinary=ordinary.size();
    for(int i=0;i<numOfOrdinary;++i){
        int stuNum=ordinary[i].stus.size();
        for(int j=0;j<14;++j){
            if(j%2==0&&ordinary[i].internationalStuExist)continue;
            bool flag=true;
            int numOfRoom=avaliableTime[j].size();
            int studentsLeft=ordinary[i].stus.size(),totStudents=ordinary[i].stus.size();
            for(int k=0;k<numOfRoom&&flag;++k){
                int numOfCourses=avaliableTime[j][k].arrangedCourse.size();
                for(int l=0;l<numOfCourses&&flag;++l){
                    int tmpStu=avaliableTime[j][k].arrangedCourse[l].stus.size();
                    for(int m=0;m<stuNum&&flag;++m){
                        for(int n=0;n<tmpStu&&flag;++n){
                            if(ordinary[i].stus[m].stuId==avaliableTime[j][k].arrangedCourse[l].stus[n].stuId)flag=false;
                        }
                    }
                }
            }
            if(flag){
                course tmpCourse=ordinary[i];
                while(studentsLeft>=maxRoomLimite)
                {
                    course tmpPartC;
                    for(int j=totStudents-studentsLeft;j<totStudents-studentsLeft+maxRoomLimite;++j)
                    {
                        tmpPartC.stus.push_back(tmpCourse.stus[j]);
                    }
                    tmpPartC.courseName=tmpCourse.courseName;
                    tmpPartC.chineseStuExist=tmpCourse.chineseStuExist;
                    tmpPartC.internationalStuExist=tmpCourse.internationalStuExist;
                    examinationRoom tmpRoom;
                    tmpRoom.arrangedCourse.push_back(tmpPartC);
                    tmpRoom.totStudents=maxRoomLimite;
                    avaliableTime[j].push_back(tmpRoom);
                    studentsLeft-=maxRoomLimite;
                }
                if(studentsLeft)
                {
                    bool done=false;
                    course tmpPartC;
                    for(int j=totStudents-studentsLeft;j<totStudents;++j)
                    {
                        tmpPartC.stus.push_back(tmpCourse.stus[j]);
                    }
                    tmpPartC.courseName=tmpCourse.courseName;
                    tmpPartC.chineseStuExist=tmpCourse.chineseStuExist;
                    tmpPartC.internationalStuExist=tmpCourse.internationalStuExist;
                    for(int k=0;k<numOfRoom&&!done;++k){
                        if(avaliableTime[j][k].totStudents+studentsLeft<=maxRoomLimite){
                            avaliableTime[j][k].totStudents+=studentsLeft;
                            avaliableTime[j][k].arrangedCourse.push_back(tmpPartC);
                            done=true;
                        }
                    }
                    if(!done){
                        examinationRoom tmpRoom;
                        tmpRoom.arrangedCourse.push_back(ordinary[i]);
                        tmpRoom.totStudents=studentsLeft;
                        avaliableTime[j].push_back(tmpRoom);
                    }
                    break;
                }
            }
        }
    }
    for(int i=0;i<14;i+=2)
    {
        QMap<QString,int>stuExist;
        int numOfRN=avaliableTime[i+1].size();
        for(int j=numOfRN-1;j>=0;--j)
        {
            int numOfC=avaliableTime[i+1][j].arrangedCourse.size();
            for(int k=0;k<numOfC;++k)
            {
                int numOfS=avaliableTime[i+1][j].arrangedCourse[k].stus.size();
                for(int l=0;l<numOfS;++l)stuExist[avaliableTime[i+1][j].arrangedCourse[k].stus[l].stuId]=1;
            }
        }
        int numOfRD=avaliableTime[i].size();
        for(int j=numOfRD-1;j>=0;--j)
        {
            if(avaliableTime[i][j].totStudents<=50)
            {
                int numOfC=avaliableTime[i][j].arrangedCourse.size();
                bool flag=false;
                for(int k=0;k<numOfC&&!flag;++k)
                {
                    int numOfS=avaliableTime[i][j].arrangedCourse[k].stus.size();
                    for(int l=0;l<numOfS&&!flag;++l)flag=stuExist[avaliableTime[i][j].arrangedCourse[k].stus[l].stuId];
                }
                if(!flag)
                {
                    for(int k=numOfRN-1;k>=0;--k)
                    {
                        if(avaliableTime[i+1][k].totStudents+avaliableTime[i][j].totStudents<=maxRoomLimite)
                        {
                            for(int l=0;l<numOfC&&!flag;++l)
                            {
                                avaliableTime[i+1][k].arrangedCourse.push_back(avaliableTime[i][j].arrangedCourse[l]);
                            }
                            avaliableTime[i+1][k].totStudents+=avaliableTime[i][j].totStudents;
                            avaliableTime[i].erase(avaliableTime[i].begin()+j);
                            break;
                        }
                    }
                }
            }
        }
    }

    //sheets->querySubObject("Add()");
    QAxObject *sheet=sheets->querySubObject("Item(int)",1);
    sheet->setProperty("Name","专业课安排");
    sheet=sheets->querySubObject("Item(const QString&)", "专业课安排");
    QAxObject *range = sheet->querySubObject("Cells(int,int)",1,1);
    range->dynamicCall("SetValue(const QString&)","时间");
    range = sheet->querySubObject("Cells(int,int)",1,2);
    range->dynamicCall("SetValue(const QString&)","考场");
    range = sheet->querySubObject("Cells(int,int)",1,3);
    range->dynamicCall("SetValue(const QString&)","考场人数");
    range = sheet->querySubObject("Cells(int,int)",1,4);
    range->dynamicCall("SetValue(const QString&)","课程名称");
    range = sheet->querySubObject("Cells(int,int)",1,5);
    range->dynamicCall("SetValue(const QString&)","原教学班");
    range = sheet->querySubObject("Cells(int,int)",1,6);
    range->dynamicCall("SetValue(const QString&)","上课教师");
    range = sheet->querySubObject("Cells(int,int)",1,7);
    range->dynamicCall("SetValue(const QString&)","教学班补考人数");
    range = sheet->querySubObject("Cells(int,int)",1,8);
    range->dynamicCall("SetValue(const QString&)","学生姓名");
    range = sheet->querySubObject("Cells(int,int)",1,9);
    range->dynamicCall("SetValue(const QString&)","学生学号");
    range = sheet->querySubObject("Cells(int,int)",1,10);
    range->dynamicCall("SetValue(const QString&)","补考确认");
    range = sheet->querySubObject("Cells(int,int)",1,11);
    range->dynamicCall("SetValue(const QString&)","学院");
    range = sheet->querySubObject("Cells(int,int)",1,12);
    range->dynamicCall("SetValue(const QString&)","班级");
    int tmp=1;
    for(int i=0;i<14;++i){
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
        int numOfRoom=avaliableTime[i].size();
        int dateStart=tmp+1;
        for(int j=0;j<numOfRoom;++j){
            QString room="教室"+QString::number(j+1);
            int roomStart=tmp+1;
            int numOfC=avaliableTime[i][j].arrangedCourse.size();
            for(int k=0;k<numOfC;++k){
                int courseStart=tmp+1,lastSame=tmp+1;
                course tmpCourse=avaliableTime[i][j].arrangedCourse[k];
                int numOfStus=tmpCourse.stus.size();
                for(int l=0;l<numOfStus;++l){
                    students tmpStu=tmpCourse.stus[l];
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
                    tmp++;
                    if(l!=0&&tmpStu.courseName!=tmpCourse.stus[l-1].courseName)
                    {
                        QString cell=("E"+QString::number(lastSame)+":"+"E"+QString::number(tmp-1));
                        range = sheet->querySubObject("Range(const QString&)",cell);
                        range->setProperty("VerticalAlignment", -4108);//xlCenter
                        range->setProperty("WrapText", true);
                        range->setProperty("MergeCells", true);
                        range->dynamicCall("SetValue(const QString&)",tmpCourse.stus[l-1].courseName);
                        cell=("F"+QString::number(lastSame)+":"+"F"+QString::number(tmp-1));
                        range = sheet->querySubObject("Range(const QString&)",cell);
                        range->setProperty("VerticalAlignment", -4108);//xlCenter
                        range->setProperty("WrapText", true);
                        range->setProperty("MergeCells", true);
                        range->dynamicCall("SetValue(const QString&)",tmpCourse.stus[l-1].teacher.split('/')[1]);
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
                range->dynamicCall("SetValue(const QString&)",tmpCourse.stus[numOfStus-1].courseName);
                cell=("F"+QString::number(lastSame)+":"+"F"+QString::number(courseEnd));
                range = sheet->querySubObject("Range(const QString&)",cell);
                range->setProperty("VerticalAlignment", -4108);//xlCenter
                range->setProperty("WrapText", true);
                range->setProperty("MergeCells", true);
                range->dynamicCall("SetValue(const QString&)",tmpCourse.stus[numOfStus-1].teacher.split('/')[1]);
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
                range->dynamicCall("SetValue(const QString&)",tmpCourse.courseName);
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
    int numOfMath=math.size();
    for(int i=0;i<numOfMath;++i){
        int numOfRoom=rooms.size();
        course tmpCourse=math[i];
        bool done=false;
        int totStudents=tmpCourse.stus.size(),studentsLeft=tmpCourse.stus.size();
        while(studentsLeft>=maxRoomLimite)
        {
            course tmpPartC;
            for(int j=totStudents-studentsLeft;j<totStudents-studentsLeft+maxRoomLimite;++j)
            {
                tmpPartC.stus.push_back(tmpCourse.stus[j]);
            }
            tmpPartC.courseName=tmpCourse.courseName;
            tmpPartC.chineseStuExist=tmpCourse.chineseStuExist;
            tmpPartC.internationalStuExist=tmpCourse.internationalStuExist;
            examinationRoom tmpRoom;
            tmpRoom.arrangedCourse.push_back(tmpPartC);
            tmpRoom.totStudents=maxRoomLimite;
            rooms.push_back(tmpRoom);
            studentsLeft-=maxRoomLimite;
        }
        if(studentsLeft)
        {
            course tmpPartC;
            for(int j=totStudents-studentsLeft;j<totStudents;++j)
            {
                tmpPartC.stus.push_back(tmpCourse.stus[j]);
            }
            tmpPartC.courseName=tmpCourse.courseName;
            tmpPartC.chineseStuExist=tmpCourse.chineseStuExist;
            tmpPartC.internationalStuExist=tmpCourse.internationalStuExist;
            for(int k=0;k<numOfRoom&&!done;++k){
                if(rooms[k].totStudents<100&&rooms[k].totStudents+studentsLeft<=maxRoomLimite){
                    int numOfCourse=rooms[k].arrangedCourse.size();
                    bool flag=tmpCourse.courseName.contains("英文");
                    for(int j=0;j<numOfCourse&&flag==tmpCourse.courseName.contains("英文");++j)flag=rooms[k].arrangedCourse[j].courseName.contains("英文");
                    if(flag!=tmpCourse.courseName.contains("英文"))continue;
                    rooms[k].totStudents+=studentsLeft;
                    rooms[k].arrangedCourse.push_back(tmpPartC);
                    done=true;
                }
            }
            if(!done){
                examinationRoom tmpRoom;
                tmpRoom.arrangedCourse.push_back(tmpPartC);
                tmpRoom.totStudents=studentsLeft;
                rooms.push_back(tmpRoom);
            }

        }

    }
    sheets->querySubObject("Add()");
    sheet=sheets->querySubObject("Item(int)",1);
    sheet->setProperty("Name","高等数学安排");
    sheet=sheets->querySubObject("Item(const QString&)", "高等数学安排");
    range = sheet->querySubObject("Cells(int,int)",1,1);
    range->dynamicCall("SetValue(const QString&)","时间");
    range = sheet->querySubObject("Cells(int,int)",1,2);
    range->dynamicCall("SetValue(const QString&)","考场");
    range = sheet->querySubObject("Cells(int,int)",1,3);
    range->dynamicCall("SetValue(const QString&)","考场人数");
    range = sheet->querySubObject("Cells(int,int)",1,4);
    range->dynamicCall("SetValue(const QString&)","课程名称");
    range = sheet->querySubObject("Cells(int,int)",1,5);
    range->dynamicCall("SetValue(const QString&)","原教学班");
    range = sheet->querySubObject("Cells(int,int)",1,6);
    range->dynamicCall("SetValue(const QString&)","上课教师");
    range = sheet->querySubObject("Cells(int,int)",1,7);
    range->dynamicCall("SetValue(const QString&)","教学班补考人数");
    range = sheet->querySubObject("Cells(int,int)",1,8);
    range->dynamicCall("SetValue(const QString&)","学生姓名");
    range = sheet->querySubObject("Cells(int,int)",1,9);
    range->dynamicCall("SetValue(const QString&)","学生学号");
    range = sheet->querySubObject("Cells(int,int)",1,10);
    range->dynamicCall("SetValue(const QString&)","补考确认");
    range = sheet->querySubObject("Cells(int,int)",1,11);
    range->dynamicCall("SetValue(const QString&)","学院");
    range = sheet->querySubObject("Cells(int,int)",1,10);
    range->dynamicCall("SetValue(const QString&)","班级");
    tmp=1;
    QString date="周六下午";
    int numOfRoom=rooms.size();
    for(int j=0;j<numOfRoom;++j){
        QString room="教室"+QString::number(j+1);
        int roomStart=tmp+1;
        int numOfC=rooms[j].arrangedCourse.size();
        for(int k=0;k<numOfC;++k){
            int courseStart=tmp+1,lastSame=tmp+1;
            course tmpCourse=rooms[j].arrangedCourse[k];
            int numOfStus=tmpCourse.stus.size();
            for(int l=0;l<numOfStus;++l){
                students tmpStu=tmpCourse.stus[l];
                range = sheet->querySubObject("Cells(int,int)",tmp+1,8);
                range->dynamicCall("SetValue(const QString&)",tmpStu.name);
                range = sheet->querySubObject("Cells(int,int)",tmp+1,9);
                range->setProperty("NumberFormatLocal", "@");
                range->dynamicCall("SetValue(const QString&)",tmpStu.stuId);
                range = sheet->querySubObject("Cells(int,int)",tmp+1,10);
                if(tmpStu.isChecked)range->dynamicCall("SetValue(const QString&)","是");
                else range->dynamicCall("SetValue(const QString&)","否");
                range->dynamicCall("SetValue(const QString&)",tmpStu.college);
                range = sheet->querySubObject("Cells(int,int)",tmp+1,11);
                range->dynamicCall("SetValue(const QString&)",tmpStu.professionalClass);
                range = sheet->querySubObject("Cells(int,int)",tmp+1,12);
                tmp++;
                if(l!=0&&tmpStu.courseName!=tmpCourse.stus[l-1].courseName)
                {
                    QString cell=("E"+QString::number(lastSame)+":"+"E"+QString::number(tmp-1));
                    range = sheet->querySubObject("Range(const QString&)",cell);
                    range->setProperty("VerticalAlignment", -4108);//xlCenter
                    range->setProperty("WrapText", true);
                    range->setProperty("MergeCells", true);
                    range->dynamicCall("SetValue(const QString&)",tmpCourse.stus[l-1].courseName);
                    cell=("F"+QString::number(lastSame)+":"+"F"+QString::number(tmp-1));
                    range = sheet->querySubObject("Range(const QString&)",cell);
                    range->setProperty("VerticalAlignment", -4108);//xlCenter
                    range->setProperty("WrapText", true);
                    range->setProperty("MergeCells", true);
                    range->dynamicCall("SetValue(const QString&)",tmpCourse.stus[l-1].teacher.split('/')[1]);
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
            range->dynamicCall("SetValue(const QString&)",tmpCourse.courseName);
            cell=("E"+QString::number(lastSame)+":"+"E"+QString::number(courseEnd));
            range = sheet->querySubObject("Range(const QString&)",cell);
            range->setProperty("VerticalAlignment", -4108);//xlCenter
            range->setProperty("WrapText", true);
            range->setProperty("MergeCells", true);
            range->dynamicCall("SetValue(const QString&)",tmpCourse.stus[numOfStus-1].courseName);
            cell=("F"+QString::number(lastSame)+":"+"F"+QString::number(courseEnd));
            range = sheet->querySubObject("Range(const QString&)",cell);
            range->setProperty("VerticalAlignment", -4108);//xlCenter
            range->setProperty("WrapText", true);
            range->setProperty("MergeCells", true);
            range->dynamicCall("SetValue(const QString&)",tmpCourse.stus[numOfStus-1].teacher.split('/')[1]);
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
    for(int i=0;i<numOfComputer;++i){
        int numOfRoom=rooms.size();
        course tmpCourse=courses[courseExist[computerA[i].courseName]-1];
        bool done=false;
        int totStudents=tmpCourse.stus.size(),studentsLeft=tmpCourse.stus.size();
        while(studentsLeft>=maxRoomLimite)
        {
            course tmpPartC;
            for(int j=totStudents-studentsLeft;j<totStudents-studentsLeft+maxRoomLimite;++j)
            {
                tmpPartC.stus.push_back(tmpCourse.stus[j]);
            }
            tmpPartC.courseName=tmpCourse.courseName;
            tmpPartC.chineseStuExist=tmpCourse.chineseStuExist;
            tmpPartC.internationalStuExist=tmpCourse.internationalStuExist;
            examinationRoom tmpRoom;
            tmpRoom.arrangedCourse.push_back(tmpPartC);
            tmpRoom.totStudents=maxRoomLimite;
            rooms.push_back(tmpRoom);
            studentsLeft-=maxRoomLimite;
        }
        if(studentsLeft)
        {
            course tmpPartC;
            for(int j=totStudents-studentsLeft;j<totStudents;++j)
            {
                tmpPartC.stus.push_back(tmpCourse.stus[j]);
            }
            tmpPartC.courseName=tmpCourse.courseName;
            tmpPartC.chineseStuExist=tmpCourse.chineseStuExist;
            tmpPartC.internationalStuExist=tmpCourse.internationalStuExist;
            for(int k=0;k<numOfRoom&&!done;++k){
                if(rooms[k].totStudents<100&&rooms[k].totStudents+studentsLeft<=maxRoomLimite){
                    rooms[k].totStudents+=studentsLeft;
                    rooms[k].arrangedCourse.push_back(tmpPartC);
                    done=true;
                }
            }
            if(!done){
                examinationRoom tmpRoom;
                tmpRoom.arrangedCourse.push_back(tmpPartC);
                tmpRoom.totStudents=studentsLeft;
                rooms.push_back(tmpRoom);
            }
        }
    }
    sheets->querySubObject("Add()");
    sheet=sheets->querySubObject("Item(int)",1);
    sheet->setProperty("Name","计算机应用A安排");
    sheet=sheets->querySubObject("Item(const QString&)", "计算机应用A安排");
    range = sheet->querySubObject("Cells(int,int)",1,1);
    range->dynamicCall("SetValue(const QString&)","时间");
    range = sheet->querySubObject("Cells(int,int)",1,2);
    range->dynamicCall("SetValue(const QString&)","考场");
    range = sheet->querySubObject("Cells(int,int)",1,3);
    range->dynamicCall("SetValue(const QString&)","考场人数");
    range = sheet->querySubObject("Cells(int,int)",1,4);
    range->dynamicCall("SetValue(const QString&)","课程名称");
    range = sheet->querySubObject("Cells(int,int)",1,5);
    range->dynamicCall("SetValue(const QString&)","原教学班");
    range = sheet->querySubObject("Cells(int,int)",1,6);
    range->dynamicCall("SetValue(const QString&)","上课教师");
    range = sheet->querySubObject("Cells(int,int)",1,7);
    range->dynamicCall("SetValue(const QString&)","教学班补考人数");
    range = sheet->querySubObject("Cells(int,int)",1,8);
    range->dynamicCall("SetValue(const QString&)","学生姓名");
    range = sheet->querySubObject("Cells(int,int)",1,9);
    range->dynamicCall("SetValue(const QString&)","学生学号");
    range = sheet->querySubObject("Cells(int,int)",1,10);
    range->dynamicCall("SetValue(const QString&)","补考确认");
    range = sheet->querySubObject("Cells(int,int)",1,11);
    range->dynamicCall("SetValue(const QString&)","学院");
    range = sheet->querySubObject("Cells(int,int)",1,12);
    range->dynamicCall("SetValue(const QString&)","班级");
    tmp=1;
    date="周六下午";

    numOfRoom=rooms.size();
    for(int j=0;j<numOfRoom;++j){
        QString room="教室"+QString::number(j+1);
        int roomStart=tmp+1;
        int numOfC=rooms[j].arrangedCourse.size();
        for(int k=0;k<numOfC;++k){
            int courseStart=tmp+1,lastSame=tmp+1;
            course tmpCourse=rooms[j].arrangedCourse[k];
            int numOfStus=tmpCourse.stus.size();
            for(int l=0;l<numOfStus;++l){
                students tmpStu=tmpCourse.stus[l];
                range = sheet->querySubObject("Cells(int,int)",tmp+1,8);
                range->dynamicCall("SetValue(const QString&)",tmpStu.name);
                range = sheet->querySubObject("Cells(int,int)",tmp+1,9);
                range->setProperty("NumberFormatLocal", "@");
                range->dynamicCall("SetValue(const QString&)",tmpStu.stuId);
                range = sheet->querySubObject("Cells(int,int)",tmp+1,10);
                if(tmpStu.isChecked)range->dynamicCall("SetValue(const QString&)","是");
                else range->dynamicCall("SetValue(const QString&)","否");
                range->dynamicCall("SetValue(const QString&)",tmpStu.college);
                range = sheet->querySubObject("Cells(int,int)",tmp+1,11);
                range->dynamicCall("SetValue(const QString&)",tmpStu.professionalClass);
                range = sheet->querySubObject("Cells(int,int)",tmp+1,12);
                tmp++;
                if(l!=0&&tmpStu.courseName!=tmpCourse.stus[l-1].courseName)
                {
                    QString cell=("E"+QString::number(lastSame)+":"+"E"+QString::number(tmp-1));
                    range = sheet->querySubObject("Range(const QString&)",cell);
                    range->setProperty("VerticalAlignment", -4108);//xlCenter
                    range->setProperty("WrapText", true);
                    range->setProperty("MergeCells", true);
                    range->dynamicCall("SetValue(const QString&)",tmpCourse.stus[l-1].courseName);
                    cell=("F"+QString::number(lastSame)+":"+"F"+QString::number(tmp-1));
                    range = sheet->querySubObject("Range(const QString&)",cell);
                    range->setProperty("VerticalAlignment", -4108);//xlCenter
                    range->setProperty("WrapText", true);
                    range->setProperty("MergeCells", true);
                    range->dynamicCall("SetValue(const QString&)",tmpCourse.stus[l-1].teacher.split('/')[1]);
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
            range->dynamicCall("SetValue(const QString&)",tmpCourse.courseName);
            cell=("E"+QString::number(lastSame)+":"+"E"+QString::number(courseEnd));
            range = sheet->querySubObject("Range(const QString&)",cell);
            range->setProperty("VerticalAlignment", -4108);//xlCenter
            range->setProperty("WrapText", true);
            range->setProperty("MergeCells", true);
            range->dynamicCall("SetValue(const QString&)",tmpCourse.stus[numOfStus-1].courseName);
            cell=("F"+QString::number(lastSame)+":"+"F"+QString::number(courseEnd));
            range = sheet->querySubObject("Range(const QString&)",cell);
            range->setProperty("VerticalAlignment", -4108);//xlCenter
            range->setProperty("WrapText", true);
            range->setProperty("MergeCells", true);
            range->dynamicCall("SetValue(const QString&)",tmpCourse.stus[numOfStus-1].teacher.split('/')[1]);
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
    numOfComputer=computerB.size();
    for(int i=0;i<numOfComputer;++i){
        int numOfRoom=rooms.size();
        course tmpCourse=courses[courseExist[computerB[i].courseName]-1];
        bool done=false;
        int totStudents=tmpCourse.stus.size(),studentsLeft=tmpCourse.stus.size();
        while(studentsLeft>=maxRoomLimite)
        {
            course tmpPartC;
            for(int j=totStudents-studentsLeft;j<totStudents-studentsLeft+maxRoomLimite;++j)
            {
                tmpPartC.stus.push_back(tmpCourse.stus[j]);
            }
            tmpPartC.courseName=tmpCourse.courseName;
            tmpPartC.chineseStuExist=tmpCourse.chineseStuExist;
            tmpPartC.internationalStuExist=tmpCourse.internationalStuExist;
            examinationRoom tmpRoom;
            tmpRoom.arrangedCourse.push_back(tmpPartC);
            tmpRoom.totStudents=maxRoomLimite;
            rooms.push_back(tmpRoom);
            studentsLeft-=maxRoomLimite;
        }
        if(studentsLeft)
        {
            course tmpPartC;
            for(int j=totStudents-studentsLeft;j<totStudents;++j)
            {
                tmpPartC.stus.push_back(tmpCourse.stus[j]);
            }
            tmpPartC.courseName=tmpCourse.courseName;
            tmpPartC.chineseStuExist=tmpCourse.chineseStuExist;
            tmpPartC.internationalStuExist=tmpCourse.internationalStuExist;
            for(int k=0;k<numOfRoom&&!done;++k){
                if(rooms[k].totStudents<100&&rooms[k].totStudents+studentsLeft<=maxRoomLimite){
                    rooms[k].totStudents+=studentsLeft;
                    rooms[k].arrangedCourse.push_back(tmpPartC);
                    done=true;
                }
            }
            if(!done){
                examinationRoom tmpRoom;
                tmpRoom.arrangedCourse.push_back(tmpPartC);
                tmpRoom.totStudents=studentsLeft;
                rooms.push_back(tmpRoom);
            }
        }
    }
    sheets->querySubObject("Add()");
    sheet=sheets->querySubObject("Item(int)",1);
    sheet->setProperty("Name","计算机应用B安排");
    sheet=sheets->querySubObject("Item(const QString&)", "计算机应用B安排");
    range = sheet->querySubObject("Cells(int,int)",1,1);
    range->dynamicCall("SetValue(const QString&)","时间");
    range = sheet->querySubObject("Cells(int,int)",1,2);
    range->dynamicCall("SetValue(const QString&)","考场");
    range = sheet->querySubObject("Cells(int,int)",1,3);
    range->dynamicCall("SetValue(const QString&)","考场人数");
    range = sheet->querySubObject("Cells(int,int)",1,4);
    range->dynamicCall("SetValue(const QString&)","课程名称");
    range = sheet->querySubObject("Cells(int,int)",1,5);
    range->dynamicCall("SetValue(const QString&)","原教学班");
    range = sheet->querySubObject("Cells(int,int)",1,6);
    range->dynamicCall("SetValue(const QString&)","上课教师");
    range = sheet->querySubObject("Cells(int,int)",1,7);
    range->dynamicCall("SetValue(const QString&)","教学班补考人数");
    range = sheet->querySubObject("Cells(int,int)",1,8);
    range->dynamicCall("SetValue(const QString&)","学生姓名");
    range = sheet->querySubObject("Cells(int,int)",1,9);
    range->dynamicCall("SetValue(const QString&)","学生学号");
    range = sheet->querySubObject("Cells(int,int)",1,10);
    range->dynamicCall("SetValue(const QString&)","补考确认");
    range = sheet->querySubObject("Cells(int,int)",1,11);
    range->dynamicCall("SetValue(const QString&)","学院");
    range = sheet->querySubObject("Cells(int,int)",1,12);
    range->dynamicCall("SetValue(const QString&)","班级");
    tmp=1;
    date="周六下午";

    numOfRoom=rooms.size();
    //dateStart=tmp+1;
    for(int j=0;j<numOfRoom;++j){
        QString room="教室"+QString::number(j+1);
        int roomStart=tmp+1;
        int numOfC=rooms[j].arrangedCourse.size();
        for(int k=0;k<numOfC;++k){
            int courseStart=tmp+1,lastSame=tmp+1;
            course tmpCourse=rooms[j].arrangedCourse[k];
            int numOfStus=tmpCourse.stus.size();
            for(int l=0;l<numOfStus;++l){
                students tmpStu=tmpCourse.stus[l];
                range = sheet->querySubObject("Cells(int,int)",tmp+1,8);
                range->dynamicCall("SetValue(const QString&)",tmpStu.name);
                range = sheet->querySubObject("Cells(int,int)",tmp+1,9);
                range->setProperty("NumberFormatLocal", "@");
                range->dynamicCall("SetValue(const QString&)",tmpStu.stuId);
                range = sheet->querySubObject("Cells(int,int)",tmp+1,10);
                if(tmpStu.isChecked)range->dynamicCall("SetValue(const QString&)","是");
                else range->dynamicCall("SetValue(const QString&)","否");
                range->dynamicCall("SetValue(const QString&)",tmpStu.college);
                range = sheet->querySubObject("Cells(int,int)",tmp+1,11);
                range->dynamicCall("SetValue(const QString&)",tmpStu.professionalClass);
                range = sheet->querySubObject("Cells(int,int)",tmp+1,12);
                tmp++;
                if(l!=0&&tmpStu.courseName!=tmpCourse.stus[l-1].courseName)
                {
                    QString cell=("E"+QString::number(lastSame)+":"+"E"+QString::number(tmp-1));
                    range = sheet->querySubObject("Range(const QString&)",cell);
                    range->setProperty("VerticalAlignment", -4108);//xlCenter
                    range->setProperty("WrapText", true);
                    range->setProperty("MergeCells", true);
                    range->dynamicCall("SetValue(const QString&)",tmpCourse.stus[l-1].courseName);
                    cell=("F"+QString::number(lastSame)+":"+"F"+QString::number(tmp-1));
                    range = sheet->querySubObject("Range(const QString&)",cell);
                    range->setProperty("VerticalAlignment", -4108);//xlCenter
                    range->setProperty("WrapText", true);
                    range->setProperty("MergeCells", true);
                    range->dynamicCall("SetValue(const QString&)",tmpCourse.stus[l-1].teacher.split('/')[1]);
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
            range->dynamicCall("SetValue(const QString&)",tmpCourse.courseName);
            cell=("E"+QString::number(lastSame)+":"+"E"+QString::number(courseEnd));
            range = sheet->querySubObject("Range(const QString&)",cell);
            range->setProperty("VerticalAlignment", -4108);//xlCenter
            range->setProperty("WrapText", true);
            range->setProperty("MergeCells", true);
            range->dynamicCall("SetValue(const QString&)",tmpCourse.stus[numOfStus-1].courseName);
            cell=("F"+QString::number(lastSame)+":"+"F"+QString::number(courseEnd));
            range = sheet->querySubObject("Range(const QString&)",cell);
            range->setProperty("VerticalAlignment", -4108);//xlCenter
            range->setProperty("WrapText", true);
            range->setProperty("MergeCells", true);
            range->dynamicCall("SetValue(const QString&)",tmpCourse.stus[numOfStus-1].teacher.split('/')[1]);
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
