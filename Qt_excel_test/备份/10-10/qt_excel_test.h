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
	//�����߳�
	QThread *firstThread;
	MyWorkThread *myTask;

	/************��ʱ��***********/
	QTimer * m_Timer;
	double timecount;//��ʱ����������
	int btnTimerOnCount;

	QCPGraph *tracerGraph = nullptr;// ���ٵ�����
	/*EXCEL*/
	qtexcel  *myexcel;
	/*λ��*/
	QCustomPlot *widgetDisplacement;
	qtGraph *GraphDisplacement;
	/*��*/
	QCustomPlot *widgetForce;
	qtGraph *GraphForce;

	void addPoint(qtGraph *graph,const double &x, const double &y);
	


private slots:
	void on_timer_timeout(); //��ʱ�������ۺ���
	void on_btnTimerOn_Clicked();//��ʱ����ʼ��ʱ
	/*λ��*/
	void on_btnReadDataDis_Clicked();//����λ��0����������
	void on_btnPlotDis_Clicked();//����λ��0������
	void on_btnSaveDataDis_Clicked();//����λ��1����������
	void on_btnReceiveDis_Clicked();//����λ��1������д�������
	void DoubleClickedPlotDis(QMouseEvent *event);//˫�����λ��1����������
	void on_btnClearPlotDis_Clicked();//���λ��1������
	void contextMenuRequestDis(QPoint pos);//�Ҽ�ͼ��˵�
	/*��*/
	void on_btnReadDataFor_Clicked();//������0����������
	void on_btnPlotFor_Clicked();//������0������
	void on_btnSaveDataFor_Clicked();//������1����������
	void DoubleClickedPlotFor(QMouseEvent *event);//˫�������1����������
	void on_btnClearPlotFor_Clicked();//�����1������
	void contextMenuRequestFor(QPoint pos);//�Ҽ�ͼ��˵�

	void on_act_zoomIn_toggled_(bool arg1);//�����ѡ�Ŵ�
	void on_act_move_toggled_(bool arg1);//����϶�
	void traceGraphSelect();//ѡ���������߽����α�
};

