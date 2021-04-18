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
	// 把QVariant转为QList<QList<QVariant> >, 用于快速读出的
	void castVariant2ListListVariant(const QVariant &var, QList<QList<QVariant> > &x_y);
	void castVariant2VectorAndVector(const QVariant &var, QVector<double> &x, QVector<double> &y);
	void castListListVariant2Variant(const QList<QList<QVariant> > &cells, QVariant &res);//把QList<QList<QVariant> > 转为QVariant,用于快速写入的
	void ListListVariant2VectorVector(const QList<QList<QVariant> > &x_y, QVector<double> &x, QVector<double> &y);
	void convert2ColName(int data, QString &res);//把列数转换为excel的字母列号
	QString to26AlphabetString(int data);//数字转换为26字母
};

