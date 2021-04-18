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
		qWarning("Qt:��ʼ��Oleʧ�ܣ�error %x��", (unsigned int)r);
	}
	qDebug() << "startThread" << "���빤���߳�:" << QThread::currentThreadId();
	myexcel = new qtexcel;
	//��GUI�̷߳��� ��������Ϊ�� �ź�
	connect(myexcel, &qtexcel::writeExcelisEmpty, this, &MyWorkThread::writeExcelIsEmpty_Thread);
}
void MyWorkThread::ReadDataDis_Thread(QString fileName, QList<QList<QVariant>> *x_y0)
{
	
	qDebug() << "ReadDataDis_Thread ���빤���߳�:"<< QThread::currentThreadId();
	myexcel->readExcelFast(fileName, *x_y0);
	emit readExcelDown2main();
	qDebug() << "readExcelFast���� �ص������߳�";

}

void MyWorkThread::WriteDataDis_Thread(QList<QList<QVariant> > *x_y1, QString fileName)
{
	qDebug() << "WriteDataDis_Thread ���빤���߳�:" << QThread::currentThreadId();
	myexcel->writeExcelFast(*x_y1, fileName);
	emit writeExcelDown2main();
	qDebug() << "writeExcelFast���� �ص������߳�";
}

void MyWorkThread::writeExcelIsEmpty_Thread()
{
	emit writeExcelIsEmpty2main();
}


