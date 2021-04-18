#include "qt_excel_test.h"
#include "qcustomplot.h"
#include "qtexcel.h"
//#include <QAxObject>
#include <QDebug>
#include <iostream>
#include <QMetaType>
#include <QThread>


Qt_excel_test::Qt_excel_test(QWidget *parent)
    : QMainWindow(parent)
{
    ui.setupUi(this);

	this->setWindowTitle("QT��ȡEXCEL��ͼ����");//���ô��ڱ���
	/************������************/
	ui.act_move->setChecked(true);//Ĭ��������Ϊ�϶�
	/************�����ź�ģ��************/
	btnTimerOnCount = 0;
	/*************��ʱ������ʱ��ģ��***************/
	m_Timer = new QTimer(this);//��ʱ��
	timecount = 50;
	time_curr = 0;
	m_Timer->setTimerType(Qt::PreciseTimer);
	m_Timer->setInterval(timecount);//���ö�ʱ���ڣ���λ������
	
	/*************PMACģ��************************/
	Pmac = new PmacDevice;
	pdwDevice = 0;
	Pmac->SelectDevice(hWindow, pdwDevice, pbSuccess);
	qDebug() << "Select:" << pbSuccess;

	dwDevice = 0;
	Pmac->Open(dwDevice, pbSuccess);
	qDebug() << "open:" << pbSuccess;

	/*����ת��*/
	castType = new CastType;
	/*exceld��д�Ի���*/
	excelProssessDlg = new ExcelDlg(this);
	/*λ�����ߴ��ڳ�ʼ��*/
	widgetDisplacement = ui.widget_displacement;
	GraphDisplacement = new qtGraph(widgetDisplacement);
	GraphDisplacement->setxAxisRange(0, 210);
	GraphDisplacement->setyAxisRange(-100, 100);
	GraphDisplacement->setxAxisName("ʱ��/s");
	GraphDisplacement->setyAxisName("λ��/mm");
	GraphDisplacement->setGraphName("Ŀ������", "��������");
	/*�����ߴ��ڳ�ʼ��*/
	widgetForce = ui.widget_force;
	GraphForce = new qtGraph(widgetForce);
	GraphForce->setxAxisRange(0, 210);
	GraphForce->setyAxisRange(2000, -2000);
	GraphForce->setxAxisName("ʱ��/s");
	GraphForce->setyAxisName("��/kg");
	GraphForce->setGraphName("Ŀ������", "��������");
	/*ע���������� ���̴߳���*/
	qRegisterMetaType<QList<QList<QVariant>>>("QList<QList<QVariant>>");
	qRegisterMetaType<QList<QList<QVariant>>>("QList<QList<QVariant>>*");
	/*********************************�źŲ�����*********************************/
	//qRegisterMetaType<TextAndNumber>("TextAndNumber");
	//��ʱ������
	connect(m_Timer, &QTimer::timeout, this, &Qt_excel_test::on_timer_timeout);
	connect(ui.btnTimerOn, &QPushButton::clicked, this, &Qt_excel_test::on_btnTimerOn_Clicked);
	
	/************λ������******************/
	//����λ������ 0������
	connect(ui.btnOpenDataDis, &QPushButton::clicked, this, &Qt_excel_test::on_btnReadDataDis_Clicked);
	//����λ������ 0������
	connect(ui.btnPlotDis, &QPushButton::clicked, this, &Qt_excel_test::on_btnPlotDis_Clicked);
	//����λ������ 1������
	connect(ui.btnSaveDataDis, &QPushButton::clicked, this, &Qt_excel_test::on_btnSaveDataDis_Clicked);
	//������������� 1������
	connect(ui.btnReceiveDis, &QPushButton::clicked, this, &Qt_excel_test::on_btnReceiveDis_Clicked);
	//˫��������ݵ�
	//connect(widgetDisplacement, &QCustomPlot::mouseDoubleClick, this, &Qt_excel_test::DoubleClickedPlotDis);
	// ���λ��1������
	connect(ui.btnClearPlotDis, &QPushButton::clicked, this, &Qt_excel_test::on_btnClearPlotDis_Clicked);
	//λ������ �Ҽ� �˵���
	connect(widgetDisplacement, &QCustomPlot::customContextMenuRequested, this, &Qt_excel_test::contextMenuRequestDis);
	//������ �α�
	connect(ui.act_tracer, &QAction::toggled, GraphDisplacement, &qtGraph::on_act_tracer_toggled_);
	connect(widgetDisplacement, &QCustomPlot::mouseMove, GraphDisplacement, &qtGraph::myMouseMoveEvent);
	/************������******************/
	//����λ������ 0������
	connect(ui.btnOpenDataForce, &QPushButton::clicked, this, &Qt_excel_test::on_btnReadDataFor_Clicked);
	//����λ������ 0������
	connect(ui.btnPlotForce, &QPushButton::clicked, this, &Qt_excel_test::on_btnPlotFor_Clicked);
	//����λ������ 1������
	connect(ui.btnSaveDataForce, &QPushButton::clicked, this, &Qt_excel_test::on_btnSaveDataFor_Clicked);
	//˫��������ݵ�
	//connect(widgetForce, &QCustomPlot::mouseDoubleClick, this, &Qt_excel_test::DoubleClickedPlotFor);
	//���λ��1������
	connect(ui.btnClearPlotForce, &QPushButton::clicked, this, &Qt_excel_test::on_btnClearPlotFor_Clicked);
	//λ������ �Ҽ� �˵���
	connect(widgetForce, &QCustomPlot::customContextMenuRequested, this, &Qt_excel_test::contextMenuRequestFor);
	//������ �α�
	connect(ui.act_tracer, &QAction::toggled, GraphForce, &qtGraph::on_act_tracer_toggled_);
	connect(widgetForce, &QCustomPlot::mouseMove, GraphForce, &qtGraph::myMouseMoveEvent);
	/***********************************************************************/
	//������ ��ѡ�Ŵ�
	connect(ui.act_zoomIn, &QAction::toggled, this, &Qt_excel_test::on_act_zoomIn_toggled_);
	//������ �϶�
	connect(ui.act_move, &QAction::toggled, this, &Qt_excel_test::on_act_move_toggled_);
	//˫��ͼ��
	connect(widgetDisplacement, &QCustomPlot::selectionChangedByUser, this, &Qt_excel_test::traceGraphSelect);

	/*********************���߳�****************/
	//excel�����߳�
	firstThread = new QThread();
	myTask = new MyWorkThread();
	myTask->moveToThread(firstThread);
	connect(firstThread, &QThread::finished, myTask, &QObject::deleteLater);//��ֹ�߳�ʱҪ����deleteLater�ۺ���
	connect(firstThread, &QThread::finished, firstThread, &QObject::deleteLater);
	//��ȡλ��0����������
	connect(this, &Qt_excel_test::sig_ReadDataDis, myTask, &MyWorkThread::ReadDataDis_Thread);
	//����λ��1����������
	connect(this, &Qt_excel_test::sig_WriteDataDis, myTask, &MyWorkThread::WriteDataDis_Thread);
	//�߳̿�ʼ
	connect(firstThread, &QThread::started, myTask, &MyWorkThread::startThread);
	firstThread->start();
	qDebug() <<"main:"<< QThread::currentThreadId();
	//��������Ϊ�� ����
	connect(myTask, &MyWorkThread::writeExcelIsEmpty2main, this, &Qt_excel_test::warnWriteExcelIsEmpty);
	//��ȡ�������
	connect(myTask, &MyWorkThread::writeExcelDown2main, this, &Qt_excel_test::writeExcelDown);
	//�����������
	connect(myTask, &MyWorkThread::readExcelDown2main, this, &Qt_excel_test::readExcelDown);

	//Pmac�߳�
	secondThread = new QThread();
	myPmacTask = new MyPmacThread();
	myPmacTask->moveToThread(secondThread);
	connect(secondThread, &QThread::finished, myPmacTask, &QObject::deleteLater);//��ֹ�߳�ʱҪ����deleteLater�ۺ���
	connect(secondThread, &QThread::finished, secondThread, &QObject::deleteLater);
	//�߳̿�ʼ
	connect(secondThread, &QThread::started, myPmacTask, &MyPmacThread::startThread);
	secondThread->start();
	qDebug() << "secondThread:" << QThread::currentThreadId();
}	

//����
 Qt_excel_test::~Qt_excel_test() 
 {
	 firstThread->quit();
	 firstThread->wait();
	 secondThread->quit();
	 secondThread->wait();
	 Pmac->Close(dwDevice);
 }
 //���1���������ݵ�
 void Qt_excel_test::addPoint(qtGraph *graph, const double &x_temp, const double &y_temp)
 {
	 graph->x1.append(x_temp);
	 graph->y1.append(y_temp);
	 QList<QVariant> xy;
	 xy.append(x_temp);
	 xy.append(y_temp);
	 graph->x_y1.append(xy);
 }

 void Qt_excel_test::clearGraph_curve1(qtGraph *myGraph)
 {
	 myGraph->x1.clear();
	 myGraph->y1.clear();
	 myGraph->x_y1.clear();
 }


 void Qt_excel_test::on_timer_timeout()
 {
	 // Pmac->GetResponse(dwDevice, "M105", pAnswer);//����������ֵ
	 //qDebug() << "i105=" << pAnswer;
	 /*1����λ��*/
	 Pmac->GetResponse(dwDevice, "M162", pAnswer);//λ��
	 displacement = pAnswer.left(pAnswer.length() - 1).toDouble() / 3072 / 8192 * 16;
	 /*2����λ��*/
	 //Pmac->GetResponse(dwDevice, "M262", pAnswer);//λ��
	 //displacement = pAnswer.left(pAnswer.length() - 1).toDouble() / 3072 * 0.005;
	 ///*��*/
	 Pmac->GetResponse(dwDevice, "P15", pAnswer);//��
	 force = pAnswer.left(pAnswer.length() - 1).toDouble();

	 CurrTime = QTime::currentTime();
	 time_ms = BaseTime.msecsTo(CurrTime);
	 time_s = time_ms / 1000;
	
	 //addPoint(GraphDisplacement, time_curr, displacement);
	 addPoint(GraphDisplacement, time_s, displacement);
	 widgetDisplacement->graph(1)->setData(GraphDisplacement->x1, GraphDisplacement->y1);
	 widgetDisplacement->replot(QCustomPlot::rpQueuedReplot);

	 addPoint(GraphForce, time_s, force);
	 widgetForce->graph(1)->setData(GraphForce->x1, GraphForce->y1);
	 widgetForce->replot(QCustomPlot::rpQueuedReplot);
	 
	 /*
	 p100_last = p100;
	 Pmac->GetResponse(dwDevice, "P100", pAnswer);
	 p100 = pAnswer.left(pAnswer.length() - 1).toInt();
	 pIncrease = p100 - p100_last;
	 qDebug() << "20ms p100=" << pIncrease;

	 Pmac->GetResponse(dwDevice, "m105", pAnswer);//δ�˲�����
	 displacement = pAnswer.left(pAnswer.length() - 1).toDouble() *0.15258789;
	 //qDebug() << "m105=" << displacement;
	 addPoint(GraphDisplacement, timeNumCount,displacement);
	 widgetDisplacement->graph(1)->setData(GraphDisplacement->x1, GraphDisplacement->y1);
	 widgetDisplacement->replot(QCustomPlot::rpQueuedReplot);
	 */


	 //QList<QVariant> xy;
	 //double x_temp = timecount / 1000;
	 //double y_temp = (timecount / 1000)*(timecount / 1000);
	 //addPoint(GraphDisplacement,x_temp, y_temp);
	 //widgetDisplacement->graph(1)->setData(GraphForce->x1, GraphForce->y1);
	 //widgetDisplacement->replot(QCustomPlot::rpQueuedReplot);
	 //timecount++;

	 //time_curr = time_curr + timecount / 1000;
	 //qDebug() << "time_curr" << time_curr;
 }

 //��ʱ���źŴ���
 void Qt_excel_test::on_btnTimerOn_Clicked()
 {
	 if (btnTimerOnCount % 2 == 0)
	 {
		 m_Timer->start();
		 
	
		// PmacfilePath = "E:\\MyPmacSpace\\MotorTest\\MotorTest.pmc";
		 //Pmac->Download(dwDevice, PmacfilePath, true, true, true, true, pbDownLoadSuccess);
		// qDebug() << "pbDownLoadSuccess:" << pbDownLoadSuccess;
		 
		 Pmac->GetResponse(dwDevice, "enable plc 4", pAnswer);

		 //Pmac->GetResponse(dwDevice, "&1B6R", pAnswer);
		// Pmac->GetResponse(dwDevice, "m114=1", pAnswer);//1�ŵ��ʹ��
		// Pmac->GetResponse(dwDevice, "&1b4r", pAnswer);
		 BaseTime = QTime::currentTime();

		 //Pmac->GetResponse(dwDevice, "#1J:8192", pAnswer);
		 
		
	 }
	 else
	 {
		 m_Timer->stop();
		// Pmac->GetResponse(dwDevice, "m114=0", pAnswer);//1�ŵ��ʧȥʹ��
		 //Pmac->GetResponse(dwDevice, "disable plc 1", pAnswer);
	 }
	 btnTimerOnCount++;
 }

 /***************λ��******************/
//��ȡλ��0����������
void Qt_excel_test::on_btnReadDataDis_Clicked()
{
	GraphDisplacement->x0.clear();
	GraphDisplacement->y0.clear();
	GraphDisplacement->x_y0.clear();
	ui.labelDataStatus->setText("-----���ڶ�ȡ����-----");
	QString strFile = QFileDialog::getOpenFileName(this, "ѡ���ļ�", "./", "�ı��ļ�(*.xls;*.xlsx;)");
	if (!strFile.isEmpty())
	{

		QElapsedTimer mstimer;
		mstimer.start();//��ʱ��ʼ
		qDebug() << "emit";
		emit sig_ReadDataDis(strFile, &GraphDisplacement->x_y0);
		qDebug() << "emit��";
		excelProssessDlg->exec();//��ʾ���ڶ�д�Ի���
		float time = (double)mstimer.nsecsElapsed() / (double)1000000;
		QString strTime = QString("%1").arg(time);
		ui.label_time->setText(strTime);
	}
}

//����λ��0������
void Qt_excel_test::on_btnPlotDis_Clicked()
{
	castType->ListListVariant2VectorVector(GraphDisplacement->x_y0, GraphDisplacement->x0, GraphDisplacement->y0);
	widgetDisplacement->graph(0)->setData(GraphDisplacement->x0, GraphDisplacement->y0);
	widgetDisplacement->replot();
	ui.labelDataStatus->setText("��ͼ���");
}

//����λ��1����������
void Qt_excel_test::on_btnSaveDataDis_Clicked()
{
	ui.labelDataStatus->setText("���ڱ�������");
	QString strFile = QFileDialog::getSaveFileName(this, "���Ϊ", "./", "�ı��ļ�(*.xlsx;)");
	if (!strFile.isEmpty())
	{
		QElapsedTimer mstimer;
		mstimer.start();//��ʱ��ʼ

		emit sig_WriteDataDis(&GraphDisplacement->x_y1, strFile);
		excelProssessDlg->exec();//��ʾ���ڶ�д�Ի���

		float time = (double)mstimer.nsecsElapsed() / (double)1000000;
		QString strTime = QString("%1").arg(time);
		ui.label_time->setText(strTime);//��ʾʱ��
	}
}

//����д��λ��1�����ߵ�����
void Qt_excel_test::on_btnReceiveDis_Clicked()
{
	double x_temp = ui.lineEdit_X->text().toDouble();
	double y_temp = ui.lineEdit_Y->text().toDouble();
	addPoint(GraphDisplacement, x_temp, y_temp);
	widgetDisplacement->graph(1)->setData(GraphDisplacement->x1, GraphDisplacement->y1);
	widgetDisplacement->replot(QCustomPlot::rpQueuedReplot);
}

//˫�����λ��1���������ݵ�
void Qt_excel_test::DoubleClickedPlotDis(QMouseEvent *event)
{
	QPoint point = event->pos();
	//��ȡ��������
	double x_pos = point.x();
	double y_pos = point.y();
	//���������� ת��Ϊ QCustomPlot �ڲ�����ֵ ��pixelToCoord ������
	//coordToPixel ������֮�෴ �ǰ��ڲ�����ֵ ת��Ϊ�ⲿ�����
	double x_val = widgetDisplacement->xAxis->pixelToCoord(x_pos);
	double y_val = widgetDisplacement->yAxis->pixelToCoord(y_pos);
	addPoint(GraphDisplacement,x_val, y_val);
	widgetDisplacement->graph(1)->setData(GraphDisplacement->x1, GraphDisplacement->y1);
	widgetDisplacement->replot(QCustomPlot::rpQueuedReplot);
}

//���λ��1������
void Qt_excel_test::on_btnClearPlotDis_Clicked()
{
	time_curr = 0;
	clearGraph_curve1(GraphDisplacement);
	widgetDisplacement->graph(1)->setData(GraphDisplacement->x1, GraphDisplacement->y1);
	widgetDisplacement->replot(QCustomPlot::rpQueuedReplot);
}

//λ������ �Ҽ��˵� ������Χ
void Qt_excel_test::contextMenuRequestDis(QPoint pos)
{
	QMenu *menu = new QMenu(this);
	menu->setAttribute(Qt::WA_DeleteOnClose);
	menu->popup(widgetDisplacement->mapToGlobal(pos));
	menu->addAction("������Χ", GraphDisplacement, &qtGraph::rescaleAxes);
}


/***************��******************/
//��ȡ��0����������
void Qt_excel_test::on_btnReadDataFor_Clicked()
{
	GraphForce->x0.clear();
	GraphForce->y0.clear();
	GraphForce->x_y0.clear();
	ui.labelDataStatus->setText("���ڶ�ȡ����");
	QString strFile = QFileDialog::getOpenFileName(this, "ѡ���ļ�", "./", "�ı��ļ�(*.xls;*.xlsx;)");
	if (!strFile.isEmpty())
	{
		QElapsedTimer mstimer;
		mstimer.start();//��ʱ��ʼ

		emit sig_ReadDataDis(strFile, &GraphForce->x_y0);
		excelProssessDlg->exec();//��ʾ���ڶ�д�Ի���

		float time = (double)mstimer.nsecsElapsed() / (double)1000000;
		QString strTime = QString("%1").arg(time);
		ui.label_time->setText(strTime);
	}
}

//������0������
void Qt_excel_test::on_btnPlotFor_Clicked()
{
	castType->ListListVariant2VectorVector(GraphForce->x_y0, GraphForce->x0, GraphForce->y0);
	widgetForce->graph(0)->setData(GraphForce->x0, GraphForce->y0);
	widgetForce->replot();
	ui.labelDataStatus->setText("��ͼ���");
}

//������1����������
void Qt_excel_test::on_btnSaveDataFor_Clicked()
{
	ui.labelDataStatus->setText("���ڱ�������");
	QString strFile = QFileDialog::getSaveFileName(this, "���Ϊ", "./", "�ı��ļ�(*.xlsx;)");
	if (!strFile.isEmpty())
	{
		QElapsedTimer mstimer;
		mstimer.start();//��ʱ��ʼ

		emit sig_WriteDataDis(&GraphForce->x_y1, strFile);
		excelProssessDlg->exec();//��ʾ���ڶ�д�Ի���

		float time = (double)mstimer.nsecsElapsed() / (double)1000000;
		QString strTime = QString("%1").arg(time);
		ui.label_time->setText(strTime);//��ʾʱ��
	}
}

//˫�������1���������ݵ�
void Qt_excel_test::DoubleClickedPlotFor(QMouseEvent *event)
{
	QPoint point = event->pos();
	//��ȡ��������
	double x_pos = point.x();
	double y_pos = point.y();
	//���������� ת��Ϊ QCustomPlot �ڲ�����ֵ ��pixelToCoord ������
	//coordToPixel ������֮�෴ �ǰ��ڲ�����ֵ ת��Ϊ�ⲿ�����
	double x_val = widgetForce->xAxis->pixelToCoord(x_pos);
	double y_val = widgetForce->yAxis->pixelToCoord(y_pos);
	addPoint(GraphForce, x_val, y_val);
	widgetForce->graph(1)->setData(GraphForce->x1, GraphForce->y1);
	widgetForce->replot(QCustomPlot::rpQueuedReplot);
}

//�����1������
void Qt_excel_test::on_btnClearPlotFor_Clicked()
{
	time_curr = 0;
	clearGraph_curve1(GraphForce);
	widgetForce->graph(1)->setData(GraphForce->x1, GraphForce->y1);
	widgetForce->replot(QCustomPlot::rpQueuedReplot);
}

//������ �Ҽ��˵� ������Χ
void Qt_excel_test::contextMenuRequestFor(QPoint pos)
{
	QMenu *menu = new QMenu(this);
	menu->setAttribute(Qt::WA_DeleteOnClose);
	menu->popup(widgetForce->mapToGlobal(pos));
	menu->addAction("������Χ", GraphForce, &qtGraph::rescaleAxes);
}

/***************/
//�����ѡ�Ŵ�
void Qt_excel_test::on_act_zoomIn_toggled_(bool arg1)
{
	if (arg1)
	{
		ui.act_move->setChecked(false);//ȡ���϶�ѡ��
		widgetDisplacement->setInteraction(QCP::iRangeDrag, false);//ȡ���϶�
		widgetDisplacement->setSelectionRectMode(QCP::SelectionRectMode::srmZoom);
		widgetForce->setInteraction(QCP::iRangeDrag, false);//ȡ���϶�
		widgetForce->setSelectionRectMode(QCP::SelectionRectMode::srmZoom);
	}
	else
	{
		widgetDisplacement->setSelectionRectMode(QCP::SelectionRectMode::srmNone);
		widgetForce->setSelectionRectMode(QCP::SelectionRectMode::srmNone);
	}
}

//����϶�
void Qt_excel_test::on_act_move_toggled_(bool arg1)
{
	if (arg1)
	{
		
		ui.act_zoomIn->setChecked(false);//ȡ����ѡ�Ŵ�ѡ��
		widgetDisplacement->setSelectionRectMode(QCP::SelectionRectMode::srmNone);
		widgetDisplacement->setInteraction(QCP::iRangeDrag, true);//ʹ���϶�
		widgetForce->setSelectionRectMode(QCP::SelectionRectMode::srmNone);
		widgetForce->setInteraction(QCP::iRangeDrag, true);//ʹ���϶�
	}
	else
	{
		widgetDisplacement->setInteraction(QCP::iRangeDrag, false);
		widgetForce->setInteraction(QCP::iRangeDrag, false);
	}
}


void Qt_excel_test::traceGraphSelect()
{
	
	for (int i = 0; i < widgetDisplacement->graphCount(); ++i)
	{
		QCPGraph *graph = widgetDisplacement->graph(i);
		QCPPlottableLegendItem *item = widgetDisplacement->legend->itemWithPlottable(graph);
		if (item->selected() || graph->selected())//ѡ�����������߻������ߵ�ͼ��
		{
			tracerGraph = graph;
		}
		
	}
}

void Qt_excel_test::warnWriteExcelIsEmpty()
{
	QMessageBox::warning(this, "����", "��������Ϊ�գ��޷����棡");
}

void Qt_excel_test::readExcelDown()
{
	ui.labelDataStatus->setText("*****��ȡ�������*****");
	excelProssessDlg->close();
}
void Qt_excel_test::writeExcelDown()
{
	ui.labelDataStatus->setText("*****�����������*****");
	excelProssessDlg->close();
}





