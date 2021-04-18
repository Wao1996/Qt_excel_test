#pragma once
#include <QObject>
#include <QVector>
#include <QVariant>
class CastType :public QObject
{
	Q_OBJECT
public:
	CastType();
	~CastType();
	// ��QVariantתΪQList<QList<QVariant> >, ���ڿ��ٶ�����
	void castVariant2ListListVariant(const QVariant &var, QList<QList<QVariant> > &x_y);
	void castVariant2VectorAndVector(const QVariant &var, QVector<double> &x, QVector<double> &y);
	void castListListVariant2Variant(const QList<QList<QVariant> > &cells, QVariant &res);//��QList<QList<QVariant> > תΪQVariant,���ڿ���д���
	void ListListVariant2VectorVector(const QList<QList<QVariant> > &x_y, QVector<double> &x, QVector<double> &y);
	void convert2ColName(int data, QString &res);//������ת��Ϊexcel����ĸ�к�
	QString to26AlphabetString(int data);//����ת��Ϊ26��ĸ
};

