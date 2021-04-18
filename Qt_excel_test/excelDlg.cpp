#include "exceldlg.h"

ExcelDlg::ExcelDlg(QWidget *parent)
	: QDialog(parent)
{
	ui.setupUi(this);
	ui.progressBar->setMinimum(0);  // 最小值
	ui.progressBar->setMaximum(0);  // 最大值
}

ExcelDlg::~ExcelDlg()
{
}
