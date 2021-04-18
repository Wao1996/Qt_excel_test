#pragma once

#include <QtWidgets/QMainWindow>
#include "ui_qt_excel_test.h"
#include "qtexcel.h"

class Qt_excel_test : public QMainWindow
{
    Q_OBJECT

public:
    Qt_excel_test(QWidget *parent = Q_NULLPTR);
	~Qt_excel_test();

private:
    Ui::Qt_excel_testClass ui;
	qtexcel  myexcel;//操作EXCEL的类

	QVector<double> x;//存储x坐标的向量
	QVector<double> y;//存储y坐标的向量
	QList<QList<QVariant>> x_y;
	
	QTimer * m_Timer;
	double timecount;//定时器触发计数
	int btnTimerOnCount;

	/*游标*/
	QCPItemTracer *tracer = nullptr; // 0曲线游标
	QCPItemTracer *tracer1 = nullptr; // 1曲线游标
	QCPItemText *tracerXLabel = nullptr; // X游标标签
	QCPItemText *tracerYLabel = nullptr;;// 0曲线Y轴游标标签
	QCPItemText *tracerY1Label = nullptr;;// 1曲线Y轴游标标签
	QCPGraph *tracerGraph = nullptr;// 跟踪的曲线
	bool tracerEnable;//游标使能
	//QCPItemLine *arrow;  	 // 箭头

	/*曲线X轴范围*/
	QCPRange QrangeX;
	QCPRange QrangeX1;
	double  QrangeX_lower;
	double  QrangeX_upper;
	double  QrangeX1_lower;
	double  QrangeX1_upper;
	bool foundRange;
	

	void addPoint(const double &x, const double &y);
	void contextMenuRequest(QPoint pos);//右键图像菜单
	void rescaleAxes();//调整视图显示
	void setVisibleTracerFalse();
	void setVisibleTracerTrue();
	

private slots:
	void on_btnReadData_Clicked();//导入数据
	void on_btnPlot_Clicked();//开始绘图
	void on_timer_timeout(); //定时溢出处理槽函数
	void on_btnSaveData_Clicked();//保存数据
	void on_btnReceive_Clicked();//接收写入的数据
	void on_btnTimerOn_Clicked();//计时器开始计时
	void DoubleClickedPlot(QMouseEvent *event);//双击鼠标
	void on_btnClearPlot1_Clicked();//清除曲线1
	void on_act_zoomIn_toggled_(bool arg1);//左键框选放大
	void on_act_move_toggled_(bool arg1);//左键拖动
	void on_act_tracer_toggled_(bool arg1);//开启游标
	void traceGraphSelect();//选择哪条曲线进行游标
	void myMouseMoveEvent(QMouseEvent *event);
	
};

