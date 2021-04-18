#pragma once

#include <QObject>
#include "qtexcel.h"
#include "qtGraph.h"
#include <windows.h>
class MyWorkThread : public QObject
{
	Q_OBJECT

public:
	MyWorkThread(QObject *parent = nullptr);
	~MyWorkThread();
private:
	/*EXCEL*/
	qtexcel  *myexcel;
signals:
	void writeExcelIsEmpty2main();
	void readExcelDown2main();
	void writeExcelDown2main();
public slots:
	void startThread();
	void ReadDataDis_Thread(QString fileName, QList<QList<QVariant>> *x_y0);
	void WriteDataDis_Thread(QList<QList<QVariant> > *x_y1, QString fileName);
	void writeExcelIsEmpty_Thread();
};
