#pragma once

#include <QObject>
#include <QDebug>	
#include <QThread>

class MyPmacThread : public QObject
{
	Q_OBJECT

public:
	MyPmacThread(QObject *parent = nullptr);
	~MyPmacThread();
public slots:
	void startThread();
};
