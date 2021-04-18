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
	/*************��������*************/
	//0������
	QVector<double> x0;//�洢x���������
	QVector<double> y0;//�洢y���������
	QList<QList<QVariant>> x_y0;
	//1������
	QVector<double> x1;//�洢x���������
	QVector<double> y1;//�洢y���������
	QList<QList<QVariant>> x_y1;

private:
	//��ͼ����
	QCustomPlot * myPlot = nullptr;
	/*************�α�*************/
	bool tracerEnable;//�α�ʹ��
	
	QCPItemTracer *tracer0 = nullptr; // 0�������α�
	QCPItemTracer *tracer1 = nullptr; // 1�������α�
	QCPItemText *tracer0Label = nullptr; // 0������X�α��ǩ
	QCPItemText *tracer1Label = nullptr;;// 0������Y���α��ǩ
	//QCPItemLine *arrow;  	 // ��ͷ

	/*********����X�᷶Χ**********/
	QCPRange QrangeX0;//0������X�᷶Χ
	double  QrangeX0_lower;
	double  QrangeX0_upper;
	QCPRange QrangeX1;//1������X�᷶Χ
	double  QrangeX1_lower;
	double  QrangeX1_upper;
	bool foundRange;

	void setVisibleTracerFalse();//�α겻�ɼ�
	void setVisibleTracerTrue();//�α�ɼ�
public slots:
	void on_act_tracer_toggled_(bool arg1);
	void myMouseMoveEvent(QMouseEvent* event);
	void rescaleAxes();
};

