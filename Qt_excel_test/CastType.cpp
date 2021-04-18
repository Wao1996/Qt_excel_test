#include "CastType.h"

CastType::CastType()
{
}

CastType::~CastType()
{
}

void CastType::castVariant2ListListVariant(const QVariant &var, QList<QList<QVariant> > &x_y)
{
	QVariantList varRows;
	varRows = var.toList();
	if (varRows.isEmpty())
	{
		return;
	}
	const int rowCount = varRows.size();
	QVariantList rowData;
	for (int i = 0; i < rowCount; ++i)
	{
		rowData = varRows[i].toList();
		x_y.push_back(rowData);
	}
}

void CastType::castVariant2VectorAndVector(const QVariant &var, QVector<double> &x, QVector<double> &y)
{
	QVariantList varRows;
	varRows = var.toList();
	if (varRows.isEmpty())
	{
		return;
	}
	const int rowCount = varRows.size();
	QVariantList rowData;
	for (int i = 0; i < rowCount; ++i)
	{
		rowData = varRows[i].toList();
		x.push_back(rowData.value(0).toDouble());
		y.push_back(rowData.value(1).toDouble());
	}
}

void CastType::ListListVariant2VectorVector(const QList<QList<QVariant>>& x_y, QVector<double>& x, QVector<double>& y)
{
	QVariant tempData;
	castListListVariant2Variant(x_y, tempData);
	castVariant2VectorAndVector(tempData, x, y);
}

void CastType::castListListVariant2Variant(const QList<QList<QVariant>>& cells, QVariant & res)
{
	QVariantList vars;
	const int rowCount = cells.size();
	for (int i = 0; i < rowCount; ++i)
	{
		vars.append(QVariant(cells[i]));
	}
	res = QVariant(vars);
}

//
// brief ������ת��Ϊexcel����ĸ�к�
// param data ����0����
// return ��ĸ�кţ���1->A 26->Z 27 AA
//
void CastType::convert2ColName(int data, QString &res)
{
	Q_ASSERT(data > 0 && data < 65535);
	int tempData = data / 26;
	if (tempData > 0)
	{
		int mode = data % 26;
		convert2ColName(mode, res);
		convert2ColName(tempData, res);
	}
	else
	{
		res = (to26AlphabetString(data) + res);
	}
}

// brief ����ת��Ϊ26��ĸ
// 1->A 26->Z
// param data
// return
//
QString CastType::to26AlphabetString(int data)
{
	QChar ch = data + 0x40;//A��Ӧ0x41
	return QString(ch);
}

