#pragma once
#include <QAxObject>
#include <QDir>

class qtexcel
{
public:
	qtexcel();
	~qtexcel();
	/*��ȡ*/
	void readExcelFast(QString fileName, QVector<double> &x, QVector<double> &y);//���ٶ�ȡ����
	void castVariant2ListListVariant(QVariant &var, QList<QList<QVariant> > &x_y);//��QVariantתΪQList<QList<QVariant> >,���ڿ��ٶ�����
	void castVariant2ListListVariant(QVariant &var, QVector<double> &x, QVector<double> &y);//����������
	void readExcelSlow(QString fileName, QVector<double> &x, QVector<double> &y);//���ٶ�ȡ����
	/*д��*/
	void writeExcelFast(QString fileName,  QList<QList<QVariant> > &x_y);//����д��
	void castListListVariant2Variant(QList<QList<QVariant> > &cells, QVariant &res);//��QList<QList<QVariant> > תΪQVariant,���ڿ���д���
	void Excel_SetCell(QAxObject *worksheet, int row, int column, QString text);//����Ԫд��
	void convert2ColName(int data, QString &res);//������ת��Ϊexcel����ĸ�к�
	QString to26AlphabetString(int data);//����ת��Ϊ26��ĸ
private:
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

