#pragma once

#include <QDialog>
#include "ui_exceldlg.h"

class ExcelDlg : public QDialog
{
	Q_OBJECT

public:
	ExcelDlg(QWidget *parent = Q_NULLPTR);
	~ExcelDlg();

private:
	Ui::ExcelDlg ui;
};
