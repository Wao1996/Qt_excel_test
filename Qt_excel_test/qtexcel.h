#pragma once
#include <QAxObject>
#include <QDir>
#include <QObject>
#include "CastType.h"
class qtexcel :public QObject
{
	Q_OBJECT
public:
	qtexcel();
	~qtexcel();
	/*读取*/
	void readExcelFast(QString fileName, QList<QList<QVariant> > &x_y);//快速读取函数
	void readExcelSlow(QString fileName, QVector<double> &x, QVector<double> &y);//慢速读取函数
	/*写入*/
	void writeExcelFast(const QList<QList<QVariant> > &x_y, QString fileName);//快速写入
	void Excel_SetCell(QAxObject *worksheet, int row, int column, QString text);//按单元写入
signals:
	void writeExcelisEmpty();
private:
	CastType  * castType;//类型转换
	QAxObject * excel;
	QAxObject * workbooks;
	QAxObject * workbook;
	QAxObject * worksheets;
	QAxObject * worksheet;
	QAxObject * usedrange_Read;//读取数据矩形区域
	QAxObject * usedrange_Write;//写入数据矩形区域
	QAxObject * rows;//行数
	QAxObject * columns;//列数
	int WorkSheetCount;//Excel文件中表的个数
	int RowsCount;//行总数
	int ColumnsCount;//列总数
	int StartRow;//数据的起始行
	int StartColumn;//数据的起始列
	QVariant var;
};

