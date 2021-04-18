#pragma once

#include <QtWidgets/QMainWindow>
#include "ui_qt_excel_test.h"
#include "qtGraph.h"
#include "myworkthread.h"
#include "mypmacthread.h"
#include "exceldlg.h"
#include "pcommDevice.h"

using namespace PCOMMSERVERLib;
class Qt_excel_test : public QMainWindow
{
    Q_OBJECT

public:
    Qt_excel_test(QWidget *parent = Q_NULLPTR);
	~Qt_excel_test();

private:
    Ui::Qt_excel_testClass ui;
	/***********PMAC通讯************/
	PCOMMSERVERLib::PmacDevice* Pmac;
	QString pAnswer ="default";
	int hWindow = 0;//父窗口
	int pdwDevice;//设备号
	int dwDevice;
	bool pbSuccess = false;//是否成功标志
	double displacement;
	double force;
	double force_unfilter;
	int p100 = 0;
	int p100_last = 0;
	int pIncrease = 0;
	//download
	QString PmacfilePath;
	bool pbDownLoadSuccess;
	//力传感器1个值等于0.15258789KG
	/***********多线程*************/
	QThread * firstThread;
	QThread * secondThread;
	MyWorkThread *myTask;
	MyPmacThread *myPmacTask;
	/************定时器、计时器***********/
	QTimer * m_Timer;//定时器
	double timecount;//定时器触发时间
	double time_curr;
	int btnTimerOnCount;

	QTime  BaseTime;
	QTime  CurrTime;
	double time_ms;//现在时间距基准时间之差 单位毫秒
	double time_s;//现在时间距基准时间之差 单位秒

	QCPGraph *tracerGraph = nullptr;// 跟踪的曲线
	/*类型转换类*/
	CastType  * castType;
	/*位移*/
	QCustomPlot *widgetDisplacement;
	qtGraph *GraphDisplacement;
	/*力*/
	QCustomPlot *widgetForce;
	qtGraph *GraphForce;
	/*正在读写对话框*/
	ExcelDlg * excelProssessDlg;

	void addPoint(qtGraph *graph,const double &x, const double &y);//添加数据点
	void clearGraph_curve1(qtGraph * myGraph);
	
signals:
	void sig_ReadDataDis(QString fileName, QList<QList<QVariant>> *x_y0);
	void sig_WriteDataDis(QList<QList<QVariant> > *x_y, QString fileName);
private slots:
	void on_timer_timeout(); //定时溢出处理槽函数
	void on_btnTimerOn_Clicked();//计时器开始计时
	/*位移*/
	void on_btnReadDataDis_Clicked();//导入位移0号曲线数据
	void on_btnPlotDis_Clicked();//绘制位移0号曲线
	void on_btnSaveDataDis_Clicked();//保存位移1号曲线数据
	void on_btnReceiveDis_Clicked();//接收位移1号曲线写入的数据
	void DoubleClickedPlotDis(QMouseEvent *event);//双击添加位移1号曲线数据
	void on_btnClearPlotDis_Clicked();//清除位移1号曲线
	void contextMenuRequestDis(QPoint pos);//右键图像菜单
	/*力*/
	void on_btnReadDataFor_Clicked();//导入力0号曲线数据
	void on_btnPlotFor_Clicked();//绘制力0号曲线
	void on_btnSaveDataFor_Clicked();//保存力1号曲线数据
	void DoubleClickedPlotFor(QMouseEvent *event);//双击添加力1号曲线数据
	void on_btnClearPlotFor_Clicked();//清除力1号曲线
	void contextMenuRequestFor(QPoint pos);//右键图像菜单

	void on_act_zoomIn_toggled_(bool arg1);//左键框选放大
	void on_act_move_toggled_(bool arg1);//左键拖动
	void traceGraphSelect();//选择哪条曲线进行游标

	void warnWriteExcelIsEmpty();
	void writeExcelDown();
	void readExcelDown();
};

