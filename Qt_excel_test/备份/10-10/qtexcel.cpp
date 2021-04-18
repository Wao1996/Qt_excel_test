#include "qtexcel.h"
#include <QVector>

qtexcel::qtexcel()
{
	excel = new QAxObject("Excel.Application");//����Excel����
	excel->setProperty("Visible", false); //����ʾExcel���棬���Ϊtrue�ῴ��������Excel����
	workbooks = excel->querySubObject("WorkBooks");
}

qtexcel::~qtexcel()
{
	delete excel;
	excel = NULL;
}

void qtexcel::readExcelFast(QString fileName, QVector<double> &x, QVector<double> &y)
{
	workbooks->querySubObject("Open(QString&)", fileName);//���ļ�·�����Ѵ��ڵĹ�����
	workbook = excel->querySubObject("ActiveWorkBook");// ��ȡ�������     
	worksheets = workbook->querySubObject("WorkSheets");// ��ȡ�򿪵�excel�ļ������еĹ���sheet
	WorkSheetCount = worksheets->property("Count").toInt();//Excel�ļ��б�ĸ���:
	worksheet = worksheets->querySubObject("Item(int)", 1);//��ȡ��һ����������������1
	usedrange_Read = worksheet->querySubObject("UsedRange");//��ȡ��sheet�����ݷ�Χ���������Ϊ�����ݵľ�������)
	//��ȡ����
	rows = usedrange_Read->querySubObject("Rows");
	RowsCount = rows->property("Count").toInt();
	//��ȡ����
	columns = usedrange_Read->querySubObject("Columns");
	ColumnsCount = columns->property("Count").toInt();
	//���ݵ���ʼ��
	StartRow = rows->property("Row").toInt();
	//���ݵ���ʼ��
	StartColumn = columns->property("Column").toInt();

	if (worksheet != NULL && !worksheet->isNull())
	{
		if (NULL == usedrange_Read || usedrange_Read->isNull())
		{
			return;
		}
		var = usedrange_Read->dynamicCall("Value");
		castVariant2ListListVariant(var,x,y); // �˺�����varת����������Ҫ�ĸ�ʽ
	}
	//һ��Ҫ�ǵ�close����Ȼϵͳ����������n��EXCEL.EXE����
	workbook->dynamicCall("Close(bool)", false);  //�ر��ļ�
	excel->dynamicCall("Quit()");
}

void qtexcel::castVariant2ListListVariant(QVariant &var,  QList<QList<QVariant> > &x_y)
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

void qtexcel::castVariant2ListListVariant(QVariant &var, QVector<double> &x, QVector<double> &y)
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
void qtexcel::readExcelSlow(QString fileName, QVector<double> &x, QVector<double> &y)
{
	workbooks->querySubObject("Open(QString&)", fileName);//���ļ�·�����ļ�
	workbook = excel->querySubObject("ActiveWorkBook");// ���ǰ������     
	worksheets = workbook->querySubObject("WorkSheets");// ��ȡ�򿪵�excel�ļ������еĹ���sheet
	WorkSheetCount = worksheets->property("Count").toInt();//Excel�ļ��б�ĸ���:
	worksheet = worksheets->querySubObject("Item(int)", 1);//��ȡ��һ����������������1
	usedrange_Read = worksheet->querySubObject("UsedRange");//��ȡ��sheet�����ݷ�Χ���������Ϊ�����ݵľ�������)
	//��ȡ����
	rows = usedrange_Read->querySubObject("Rows");
	RowsCount = rows->property("Count").toInt();
	//��ȡ����
	columns = usedrange_Read->querySubObject("Columns");
	ColumnsCount = columns->property("Count").toInt();
	//���ݵ���ʼ��
	StartRow = rows->property("Row").toInt();
	//���ݵ���ʼ��
	StartColumn = columns->property("Column").toInt();

	//������������������������������������(����)��������������������������
	QAxObject *cell_x;        // ���ڶ�λ��ָ��
	QAxObject *cell_y;
	QVariant cell_value_x;    // �洢ֵ��Ϣ
	QVariant cell_value_y;
	int j = StartColumn;
	for (int i = StartRow; i <= RowsCount; ++i)
	{
		cell_x = worksheet->querySubObject("Cells(int, int)", i, j);
		cell_y = worksheet->querySubObject("Cells(int, int)", i, j+1);
		cell_value_x = cell_x->property("Value"); // ��ȡ��Ԫ������
		cell_value_y = cell_y->property("Value");
		//һ��Ҫע�⣬���cell_value��ֵ����Ч�Ļ���������DCOM��������û��Microsoft ExcelӦ�ó���û�еĻ����
		x.push_back(cell_value_x.toDouble());
		y.push_back(cell_value_y.toDouble());
	}
	//һ��Ҫ�ǵ�close����Ȼϵͳ����������n��EXCEL.EXE����
	workbook->dynamicCall("Close(bool)", false);  //�ر��ļ�
	excel->dynamicCall("Quit()");
}

void qtexcel::Excel_SetCell(QAxObject * worksheet, int row, int column, QString text)
{
	QAxObject *cell = worksheet->querySubObject("Cells(int,int)", row, column);
	cell->setProperty("Value", text);
}

void qtexcel::writeExcelFast(QString fileName, QList<QList<QVariant>> & x_y)
{
	workbooks->dynamicCall("Add");//�½�һ�������� �¹�������Ϊ�������
	workbook = excel->querySubObject("ActiveWorkBook");// ��ȡ������� 
	worksheet = workbook->querySubObject("Sheets(int)", 1); //��ȡ��һ����������������1
	
	int row = x_y.size();//����
	int col = x_y.at(0).size();//����
	/*������ת����EXCEL���е���ĸ��ʽ*/
	QString rangStr;
	convert2ColName(col, rangStr);
	rangStr += QString::number(row);
	rangStr = "A1:" + rangStr;
	usedrange_Write = worksheet->querySubObject("Range(const QString&)", rangStr);

	QVariant var;
	castListListVariant2Variant(x_y, var);
	usedrange_Write->setProperty("Value", var);

	workbook->dynamicCall("SaveCopyAs(QString)", QDir::toNativeSeparators(fileName));
	workbook->dynamicCall("Close(bool)", false);  //�ر��ļ�
	excel->dynamicCall("Quit()");//�ر�excel
}

void qtexcel::castListListVariant2Variant(QList<QList<QVariant>>& cells, QVariant & res)
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
void qtexcel::convert2ColName(int data, QString &res)
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
QString qtexcel::to26AlphabetString(int data)
{
	QChar ch = data + 0x40;//A��Ӧ0x41
	return QString(ch);
}
