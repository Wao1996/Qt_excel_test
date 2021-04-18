#include "qtexcel.h"
#include <QVector>

qtexcel::qtexcel()
{
	castType = new CastType;
	excel = new QAxObject("Excel.Application");//����Excel����
	excel->setProperty("Visible", false); //����ʾExcel���棬���Ϊtrue�ῴ��������Excel����
	workbooks = excel->querySubObject("WorkBooks");
}

qtexcel::~qtexcel()
{
	delete excel;
	delete castType;
	excel = nullptr;
}

void qtexcel::readExcelFast(QString fileName, QList<QList<QVariant>>& x_y)
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
		castType->castVariant2ListListVariant(var, x_y); // �˺�����varת����������Ҫ�ĸ�ʽ
	}
	//һ��Ҫ�ǵ�close����Ȼϵͳ����������n��EXCEL.EXE����
	workbook->dynamicCall("Close(bool)", false);  //�ر��ļ�
	excel->dynamicCall("Quit()");
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
		cell_y = worksheet->querySubObject("Cells(int, int)", i, j + 1);
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

void qtexcel::writeExcelFast(const QList<QList<QVariant>> & x_y, QString fileName)
{
	if (x_y.isEmpty())
	{
		//�����Ϊ������
		emit writeExcelisEmpty();
	}
	else
	{
		workbooks->dynamicCall("Add");//�½�һ�������� �¹�������Ϊ�������
		workbook = excel->querySubObject("ActiveWorkBook");// ��ȡ������� 
		worksheet = workbook->querySubObject("Sheets(int)", 1); //��ȡ��һ����������������1

		int row = x_y.size();//����
		int col = x_y.at(0).size();//����
		/*������ת����EXCEL���е���ĸ��ʽ*/
		QString rangStr;
		castType->convert2ColName(col, rangStr);
		rangStr += QString::number(row);
		rangStr = "A1:" + rangStr;
		usedrange_Write = worksheet->querySubObject("Range(const QString&)", rangStr);

		QVariant var;
		castType->castListListVariant2Variant(x_y, var);
		usedrange_Write->setProperty("Value", var);

		workbook->dynamicCall("SaveCopyAs(QString)", QDir::toNativeSeparators(fileName));
		workbook->dynamicCall("Close(bool)", false);  //�ر��ļ�
		excel->dynamicCall("Quit()");//�ر�excel
	}
}

