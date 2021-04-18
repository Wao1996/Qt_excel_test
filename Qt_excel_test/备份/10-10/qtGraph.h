#pragma once
#include <QAxObject>
#include "qcustomplot.h"

class qtGraph : public QObject
{
	Q_OBJECT

public:
	qtGraph(QCustomPlot * Plot);
	~qtGraph();

	void setxAxisName(QString name);
	void setyAxisName(QString name);
	void setxAxisRange(double lower, double upper);
	void setyAxisRange(double lower, double upper);
	void setGraphName(QString name0, QString name1);


public:
	/*************曲线数据*************/
	//0号曲线
	QVector<double> x0;//存储x坐标的向量
	QVector<double> y0;//存储y坐标的向量
	QList<QList<QVariant>> x_y0;
	//1号曲线
	QVector<double> x1;//存储x坐标的向量
	QVector<double> y1;//存储y坐标的向量
	QList<QList<QVariant>> x_y1;

private:
	//绘图窗口
	QCustomPlot * myPlot = nullptr;
	/*************游标*************/
	bool tracerEnable;//游标使能
	
	QCPItemTracer *tracer0 = nullptr; // 0号曲线游标
	QCPItemTracer *tracer1 = nullptr; // 1号曲线游标
	QCPItemText *tracerX0Label = nullptr; // 0号曲线X游标标签
	QCPItemText *tracerY0Label = nullptr;;// 0号曲线Y轴游标标签
	QCPItemText *tracerY1Label = nullptr;;// 1号曲线Y轴游标标签
	//QCPItemLine *arrow;  	 // 箭头

	/*********曲线X轴范围**********/
	QCPRange QrangeX0;//0号曲线X轴范围
	double  QrangeX0_lower;
	double  QrangeX0_upper;
	QCPRange QrangeX1;//1号曲线X轴范围
	double  QrangeX1_lower;
	double  QrangeX1_upper;
	bool foundRange;

	void setVisibleTracerFalse();//游标不可见
	void setVisibleTracerTrue();//游标可见
public slots:
	void on_act_tracer_toggled_(bool arg1);
	void myMouseMoveEvent(QMouseEvent* event);
	void rescaleAxes();
};

