#include "mypmacthread.h"

MyPmacThread::MyPmacThread(QObject *parent) : QObject(parent)
{
}

MyPmacThread::~MyPmacThread()
{
}

void MyPmacThread::startThread()
{
	qDebug() << "startThread" << "进入Pmac线程:" << QThread::currentThreadId();
}
