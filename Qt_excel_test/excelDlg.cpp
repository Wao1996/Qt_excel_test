#include "exceldlg.h"

ExcelDlg::ExcelDlg(QWidget *parent)
	: QDialog(parent)
{
	ui.setupUi(this);
	ui.progressBar->setMinimum(0);  // ��Сֵ
	ui.progressBar->setMaximum(0);  // ���ֵ
}

ExcelDlg::~ExcelDlg()
{
}
