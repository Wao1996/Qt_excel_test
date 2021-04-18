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
	/***********PMACͨѶ************/
	PCOMMSERVERLib::PmacDevice* Pmac;
	QString pAnswer ="default";
	int hWindow = 0;//������
	int pdwDevice;//�豸��
	int dwDevice;
	bool pbSuccess = false;//�Ƿ�ɹ���־
	double displacement;
	double force;
	double force_unfilter;
	int p100 = 0;
	int p100_last = 0;
	int pIncrease = 0;
	//download
	QString PmacfilePath;
	bool pbDownLoadSuccess;
	//��������1��ֵ����0.15258789KG
	/***********���߳�*************/
	QThread * firstThread;
	QThread * secondThread;
	MyWorkThread *myTask;
	MyPmacThread *myPmacTask;
	/************��ʱ������ʱ��***********/
	QTimer * m_Timer;//��ʱ��
	double timecount;//��ʱ������ʱ��
	double time_curr;
	int btnTimerOnCount;

	QTime  BaseTime;
	QTime  CurrTime;
	double time_ms;//����ʱ����׼ʱ��֮�� ��λ����
	double time_s;//����ʱ����׼ʱ��֮�� ��λ��

	QCPGraph *tracerGraph = nullptr;// ���ٵ�����
	/*����ת����*/
	CastType  * castType;
	/*λ��*/
	QCustomPlot *widgetDisplacement;
	qtGraph *GraphDisplacement;
	/*��*/
	QCustomPlot *widgetForce;
	qtGraph *GraphForce;
	/*���ڶ�д�Ի���*/
	ExcelDlg * excelProssessDlg;

	void addPoint(qtGraph *graph,const double &x, const double &y);//������ݵ�
	void clearGraph_curve1(qtGraph * myGraph);
	
signals:
	void sig_ReadDataDis(QString fileName, QList<QList<QVariant>> *x_y0);
	void sig_WriteDataDis(QList<QList<QVariant> > *x_y, QString fileName);
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

	void warnWriteExcelIsEmpty();
	void writeExcelDown();
	void readExcelDown();
};

