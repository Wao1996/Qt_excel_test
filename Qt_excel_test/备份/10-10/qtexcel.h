#pragma once
#include <QAxObject>
#include <QDir>

class qtexcel
{
public:
	qtexcel();
	~qtexcel();
	/*读取*/
	void readExcelFast(QString fileName, QVector<double> &x, QVector<double> &y);//快速读取函数
	void castVariant2ListListVariant(QVariant &var, QList<QList<QVariant> > &x_y);//把QVariant转为QList<QList<QVariant> >,用于快速读出的
	void castVariant2ListListVariant(QVariant &var, QVector<double> &x, QVector<double> &y);//函数的重载
	void readExcelSlow(QString fileName, QVector<double> &x, QVector<double> &y);//慢速读取函数
	/*写入*/
	void writeExcelFast(QString fileName,  QList<QList<QVariant> > &x_y);//快速写入
	void castListListVariant2Variant(QList<QList<QVariant> > &cells, QVariant &res);//把QList<QList<QVariant> > 转为QVariant,用于快速写入的
	void Excel_SetCell(QAxObject *worksheet, int row, int column, QString text);//按单元写入
	void convert2ColName(int data, QString &res);//把列数转换为excel的字母列号
	QString to26AlphabetString(int data);//数字转换为26字母
private:
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

