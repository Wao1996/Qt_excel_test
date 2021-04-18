#pragma once

#include <QtWidgets/QMainWindow>
#include "ui_qt_excel_test.h"
#include "qtexcel.h"
#include "qtGraph.h"
#include "myworkthread.h"

class Qt_excel_test : public QMainWindow
{
    Q_OBJECT

public:
    Qt_excel_test(QWidget *parent = Q_NULLPTR);
	~Qt_excel_test();

private:
    Ui::Qt_excel_testClass ui;
	//工作线程
	QThread *firstThread;
	MyWorkThread *myTask;

	/************定时器***********/
	QTimer * m_Timer;
	double timecount;//定时器触发计数
	int btnTimerOnCount;

	QCPGraph *tracerGraph = nullptr;// 跟踪的曲线
	/*EXCEL*/
	qtexcel  *myexcel;
	/*位移*/
	QCustomPlot *widgetDisplacement;
	qtGraph *GraphDisplacement;
	/*力*/
	QCustomPlot *widgetForce;
	qtGraph *GraphForce;

	void addPoint(qtGraph *graph,const double &x, const double &y);
	


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
};

