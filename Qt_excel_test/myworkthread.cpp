#include "myworkthread.h"
#include <QDebug>
MyWorkThread::MyWorkThread(QObject *parent) : QObject(parent)
{
	
}

MyWorkThread::~MyWorkThread()
{
	OleUninitialize();
}

void MyWorkThread::startThread()
{
	HRESULT r = OleInitialize(0);
	if (r != S_OK && r != S_FALSE)
	{
		qWarning("Qt:初始化Ole失败（error %x）", (unsigned int)r);
	}
	qDebug() << "startThread" << "进入工作线程:" << QThread::currentThreadId();
	myexcel = new qtexcel;
	//向GUI线程发送 保存数据为空 信号
	connect(myexcel, &qtexcel::writeExcelisEmpty, this, &MyWorkThread::writeExcelIsEmpty_Thread);
}
void MyWorkThread::ReadDataDis_Thread(QString fileName, QList<QList<QVariant>> *x_y0)
{
	
	qDebug() << "ReadDataDis_Thread 进入工作线程:"<< QThread::currentThreadId();
	myexcel->readExcelFast(fileName, *x_y0);
	emit readExcelDown2main();
	qDebug() << "readExcelFast结束 回到工作线程";

}

void MyWorkThread::WriteDataDis_Thread(QList<QList<QVariant> > *x_y1, QString fileName)
{
	qDebug() << "WriteDataDis_Thread 进入工作线程:" << QThread::currentThreadId();
	myexcel->writeExcelFast(*x_y1, fileName);
	emit writeExcelDown2main();
	qDebug() << "writeExcelFast结束 回到工作线程";
}

void MyWorkThread::writeExcelIsEmpty_Thread()
{
	emit writeExcelIsEmpty2main();
}


