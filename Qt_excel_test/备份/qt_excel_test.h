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
	qtexcel  myexcel;//����EXCEL����

	QVector<double> x;//�洢x���������
	QVector<double> y;//�洢y���������
	QList<QList<QVariant>> x_y;
	
	QTimer * m_Timer;
	double timecount;//��ʱ����������
	int btnTimerOnCount;

	/*�α�*/
	QCPItemTracer *tracer = nullptr; // 0�����α�
	QCPItemTracer *tracer1 = nullptr; // 1�����α�
	QCPItemText *tracerXLabel = nullptr; // X�α��ǩ
	QCPItemText *tracerYLabel = nullptr;;// 0����Y���α��ǩ
	QCPItemText *tracerY1Label = nullptr;;// 1����Y���α��ǩ
	QCPGraph *tracerGraph = nullptr;// ���ٵ�����
	bool tracerEnable;//�α�ʹ��
	//QCPItemLine *arrow;  	 // ��ͷ

	/*����X�᷶Χ*/
	QCPRange QrangeX;
	QCPRange QrangeX1;
	double  QrangeX_lower;
	double  QrangeX_upper;
	double  QrangeX1_lower;
	double  QrangeX1_upper;
	bool foundRange;
	

	void addPoint(const double &x, const double &y);
	void contextMenuRequest(QPoint pos);//�Ҽ�ͼ��˵�
	void rescaleAxes();//������ͼ��ʾ
	void setVisibleTracerFalse();
	void setVisibleTracerTrue();
	

private slots:
	void on_btnReadData_Clicked();//��������
	void on_btnPlot_Clicked();//��ʼ��ͼ
	void on_timer_timeout(); //��ʱ�������ۺ���
	void on_btnSaveData_Clicked();//��������
	void on_btnReceive_Clicked();//����д�������
	void on_btnTimerOn_Clicked();//��ʱ����ʼ��ʱ
	void DoubleClickedPlot(QMouseEvent *event);//˫�����
	void on_btnClearPlot1_Clicked();//�������1
	void on_act_zoomIn_toggled_(bool arg1);//�����ѡ�Ŵ�
	void on_act_move_toggled_(bool arg1);//����϶�
	void on_act_tracer_toggled_(bool arg1);//�����α�
	void traceGraphSelect();//ѡ���������߽����α�
	void myMouseMoveEvent(QMouseEvent *event);
	
};

