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
	/*��ȡ*/
	void readExcelFast(QString fileName, QList<QList<QVariant> > &x_y);//���ٶ�ȡ����
	void readExcelSlow(QString fileName, QVector<double> &x, QVector<double> &y);//���ٶ�ȡ����
	/*д��*/
	void writeExcelFast(const QList<QList<QVariant> > &x_y, QString fileName);//����д��
	void Excel_SetCell(QAxObject *worksheet, int row, int column, QString text);//����Ԫд��
signals:
	void writeExcelisEmpty();
private:
	CastType  * castType;//����ת��
	QAxObject * excel;
	QAxObject * workbooks;
	QAxObject * workbook;
	QAxObject * worksheets;
	QAxObject * worksheet;
	QAxObject * usedrange_Read;//��ȡ���ݾ�������
	QAxObject * usedrange_Write;//д�����ݾ�������
	QAxObject * rows;//����
	QAxObject * columns;//����
	int WorkSheetCount;//Excel�ļ��б�ĸ���
	int RowsCount;//������
	int ColumnsCount;//������
	int StartRow;//���ݵ���ʼ��
	int StartColumn;//���ݵ���ʼ��
	QVariant var;
};

