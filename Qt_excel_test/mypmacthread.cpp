#include "mypmacthread.h"

MyPmacThread::MyPmacThread(QObject *parent) : QObject(parent)
{
}

MyPmacThread::~MyPmacThread()
{
}

void MyPmacThread::startThread()
{
	qDebug() << "startThread" << "����Pmac�߳�:" << QThread::currentThreadId();
}
