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

	this->setWindowTitle("QT读取EXCEL绘图测试");//设置窗口标题
	/************工具栏************/
	ui.act_move->setChecked(true);//默认鼠标左键为拖动
	/************接收信号模块************/
	btnTimerOnCount = 0;
	/*************定时器、计时器模块***************/
	m_Timer = new QTimer(this);//计时器
	timecount = 50;
	time_curr = 0;
	m_Timer->setTimerType(Qt::PreciseTimer);
	m_Timer->setInterval(timecount);//设置定时周期，单位：毫秒
	
	/*************PMAC模块************************/
	Pmac = new PmacDevice;
	pdwDevice = 0;
	Pmac->SelectDevice(hWindow, pdwDevice, pbSuccess);
	qDebug() << "Select:" << pbSuccess;

	dwDevice = 0;
	Pmac->Open(dwDevice, pbSuccess);
	qDebug() << "open:" << pbSuccess;

	/*类型转换*/
	castType = new CastType;
	/*exceld读写对话框*/
	excelProssessDlg = new ExcelDlg(this);
	/*位移曲线窗口初始化*/
	widgetDisplacement = ui.widget_displacement;
	GraphDisplacement = new qtGraph(widgetDisplacement);
	GraphDisplacement->setxAxisRange(0, 210);
	GraphDisplacement->setyAxisRange(-100, 100);
	GraphDisplacement->setxAxisName("时间/s");
	GraphDisplacement->setyAxisName("位移/mm");
	GraphDisplacement->setGraphName("目标曲线", "接收曲线");
	/*力曲线窗口初始化*/
	widgetForce = ui.widget_force;
	GraphForce = new qtGraph(widgetForce);
	GraphForce->setxAxisRange(0, 210);
	GraphForce->setyAxisRange(2000, -2000);
	GraphForce->setxAxisName("时间/s");
	GraphForce->setyAxisName("力/kg");
	GraphForce->setGraphName("目标曲线", "接收曲线");
	/*注册数据类型 跨线程传递*/
	qRegisterMetaType<QList<QList<QVariant>>>("QList<QList<QVariant>>");
	qRegisterMetaType<QList<QList<QVariant>>>("QList<QList<QVariant>>*");
	/*********************************信号槽连接*********************************/
	//qRegisterMetaType<TextAndNumber>("TextAndNumber");
	//定时器触发
	connect(m_Timer, &QTimer::timeout, this, &Qt_excel_test::on_timer_timeout);
	connect(ui.btnTimerOn, &QPushButton::clicked, this, &Qt_excel_test::on_btnTimerOn_Clicked);
	
	/************位移曲线******************/
	//导入位移数据 0号曲线
	connect(ui.btnOpenDataDis, &QPushButton::clicked, this, &Qt_excel_test::on_btnReadDataDis_Clicked);
	//绘制位移曲线 0号曲线
	connect(ui.btnPlotDis, &QPushButton::clicked, this, &Qt_excel_test::on_btnPlotDis_Clicked);
	//保存位移曲线 1号曲线
	connect(ui.btnSaveDataDis, &QPushButton::clicked, this, &Qt_excel_test::on_btnSaveDataDis_Clicked);
	//接收输入的数据 1号曲线
	connect(ui.btnReceiveDis, &QPushButton::clicked, this, &Qt_excel_test::on_btnReceiveDis_Clicked);
	//双击添加数据点
	//connect(widgetDisplacement, &QCustomPlot::mouseDoubleClick, this, &Qt_excel_test::DoubleClickedPlotDis);
	// 清除位移1号曲线
	connect(ui.btnClearPlotDis, &QPushButton::clicked, this, &Qt_excel_test::on_btnClearPlotDis_Clicked);
	//位移曲线 右键 菜单栏
	connect(widgetDisplacement, &QCustomPlot::customContextMenuRequested, this, &Qt_excel_test::contextMenuRequestDis);
	//工具栏 游标
	connect(ui.act_tracer, &QAction::toggled, GraphDisplacement, &qtGraph::on_act_tracer_toggled_);
	connect(widgetDisplacement, &QCustomPlot::mouseMove, GraphDisplacement, &qtGraph::myMouseMoveEvent);
	/************力曲线******************/
	//导入位移数据 0号曲线
	connect(ui.btnOpenDataForce, &QPushButton::clicked, this, &Qt_excel_test::on_btnReadDataFor_Clicked);
	//绘制位移曲线 0号曲线
	connect(ui.btnPlotForce, &QPushButton::clicked, this, &Qt_excel_test::on_btnPlotFor_Clicked);
	//保存位移曲线 1号曲线
	connect(ui.btnSaveDataForce, &QPushButton::clicked, this, &Qt_excel_test::on_btnSaveDataFor_Clicked);
	//双击添加数据点
	//connect(widgetForce, &QCustomPlot::mouseDoubleClick, this, &Qt_excel_test::DoubleClickedPlotFor);
	//清除位移1号曲线
	connect(ui.btnClearPlotForce, &QPushButton::clicked, this, &Qt_excel_test::on_btnClearPlotFor_Clicked);
	//位移曲线 右键 菜单栏
	connect(widgetForce, &QCustomPlot::customContextMenuRequested, this, &Qt_excel_test::contextMenuRequestFor);
	//工具栏 游标
	connect(ui.act_tracer, &QAction::toggled, GraphForce, &qtGraph::on_act_tracer_toggled_);
	connect(widgetForce, &QCustomPlot::mouseMove, GraphForce, &qtGraph::myMouseMoveEvent);
	/***********************************************************************/
	//工具栏 框选放大
	connect(ui.act_zoomIn, &QAction::toggled, this, &Qt_excel_test::on_act_zoomIn_toggled_);
	//工具栏 拖动
	connect(ui.act_move, &QAction::toggled, this, &Qt_excel_test::on_act_move_toggled_);
	//双击图例
	connect(widgetDisplacement, &QCustomPlot::selectionChangedByUser, this, &Qt_excel_test::traceGraphSelect);

	/*********************多线程****************/
	//excel工作线程
	firstThread = new QThread();
	myTask = new MyWorkThread();
	myTask->moveToThread(firstThread);
	connect(firstThread, &QThread::finished, myTask, &QObject::deleteLater);//终止线程时要调用deleteLater槽函数
	connect(firstThread, &QThread::finished, firstThread, &QObject::deleteLater);
	//读取位移0号曲线数据
	connect(this, &Qt_excel_test::sig_ReadDataDis, myTask, &MyWorkThread::ReadDataDis_Thread);
	//保存位移1号曲线数据
	connect(this, &Qt_excel_test::sig_WriteDataDis, myTask, &MyWorkThread::WriteDataDis_Thread);
	//线程开始
	connect(firstThread, &QThread::started, myTask, &MyWorkThread::startThread);
	firstThread->start();
	qDebug() <<"main:"<< QThread::currentThreadId();
	//保存数据为空 警告
	connect(myTask, &MyWorkThread::writeExcelIsEmpty2main, this, &Qt_excel_test::warnWriteExcelIsEmpty);
	//读取数据完成
	connect(myTask, &MyWorkThread::writeExcelDown2main, this, &Qt_excel_test::writeExcelDown);
	//保存数据完成
	connect(myTask, &MyWorkThread::readExcelDown2main, this, &Qt_excel_test::readExcelDown);

	//Pmac线程
	secondThread = new QThread();
	myPmacTask = new MyPmacThread();
	myPmacTask->moveToThread(secondThread);
	connect(secondThread, &QThread::finished, myPmacTask, &QObject::deleteLater);//终止线程时要调用deleteLater槽函数
	connect(secondThread, &QThread::finished, secondThread, &QObject::deleteLater);
	//线程开始
	connect(secondThread, &QThread::started, myPmacTask, &MyPmacThread::startThread);
	secondThread->start();
	qDebug() << "secondThread:" << QThread::currentThreadId();
}	

//析构
 Qt_excel_test::~Qt_excel_test() 
 {
	 firstThread->quit();
	 firstThread->wait();
	 secondThread->quit();
	 secondThread->wait();
	 Pmac->Close(dwDevice);
 }
 //添加1号曲线数据点
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
	 // Pmac->GetResponse(dwDevice, "M105", pAnswer);//力传感器的值
	 //qDebug() << "i105=" << pAnswer;
	 /*1号轴位移*/
	 Pmac->GetResponse(dwDevice, "M162", pAnswer);//位移
	 displacement = pAnswer.left(pAnswer.length() - 1).toDouble() / 3072 / 8192 * 16;
	 /*2号轴位移*/
	 //Pmac->GetResponse(dwDevice, "M262", pAnswer);//位移
	 //displacement = pAnswer.left(pAnswer.length() - 1).toDouble() / 3072 * 0.005;
	 ///*力*/
	 Pmac->GetResponse(dwDevice, "P15", pAnswer);//力
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

	 Pmac->GetResponse(dwDevice, "m105", pAnswer);//未滤波的力
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

 //计时器信号触发
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
		// Pmac->GetResponse(dwDevice, "m114=1", pAnswer);//1号电机使能
		// Pmac->GetResponse(dwDevice, "&1b4r", pAnswer);
		 BaseTime = QTime::currentTime();

		 //Pmac->GetResponse(dwDevice, "#1J:8192", pAnswer);
		 
		
	 }
	 else
	 {
		 m_Timer->stop();
		// Pmac->GetResponse(dwDevice, "m114=0", pAnswer);//1号电机失去使能
		 //Pmac->GetResponse(dwDevice, "disable plc 1", pAnswer);
	 }
	 btnTimerOnCount++;
 }

 /***************位移******************/
//读取位移0号曲线数据
void Qt_excel_test::on_btnReadDataDis_Clicked()
{
	GraphDisplacement->x0.clear();
	GraphDisplacement->y0.clear();
	GraphDisplacement->x_y0.clear();
	ui.labelDataStatus->setText("-----正在读取数据-----");
	QString strFile = QFileDialog::getOpenFileName(this, "选择文件", "./", "文本文件(*.xls;*.xlsx;)");
	if (!strFile.isEmpty())
	{

		QElapsedTimer mstimer;
		mstimer.start();//计时开始
		qDebug() << "emit";
		emit sig_ReadDataDis(strFile, &GraphDisplacement->x_y0);
		qDebug() << "emit完";
		excelProssessDlg->exec();//显示正在读写对话框
		float time = (double)mstimer.nsecsElapsed() / (double)1000000;
		QString strTime = QString("%1").arg(time);
		ui.label_time->setText(strTime);
	}
}

//绘制位移0号曲线
void Qt_excel_test::on_btnPlotDis_Clicked()
{
	castType->ListListVariant2VectorVector(GraphDisplacement->x_y0, GraphDisplacement->x0, GraphDisplacement->y0);
	widgetDisplacement->graph(0)->setData(GraphDisplacement->x0, GraphDisplacement->y0);
	widgetDisplacement->replot();
	ui.labelDataStatus->setText("绘图完成");
}

//保存位移1号曲线数据
void Qt_excel_test::on_btnSaveDataDis_Clicked()
{
	ui.labelDataStatus->setText("正在保存数据");
	QString strFile = QFileDialog::getSaveFileName(this, "另存为", "./", "文本文件(*.xlsx;)");
	if (!strFile.isEmpty())
	{
		QElapsedTimer mstimer;
		mstimer.start();//计时开始

		emit sig_WriteDataDis(&GraphDisplacement->x_y1, strFile);
		excelProssessDlg->exec();//显示正在读写对话框

		float time = (double)mstimer.nsecsElapsed() / (double)1000000;
		QString strTime = QString("%1").arg(time);
		ui.label_time->setText(strTime);//显示时间
	}
}

//接收写入位移1号曲线的数据
void Qt_excel_test::on_btnReceiveDis_Clicked()
{
	double x_temp = ui.lineEdit_X->text().toDouble();
	double y_temp = ui.lineEdit_Y->text().toDouble();
	addPoint(GraphDisplacement, x_temp, y_temp);
	widgetDisplacement->graph(1)->setData(GraphDisplacement->x1, GraphDisplacement->y1);
	widgetDisplacement->replot(QCustomPlot::rpQueuedReplot);
}

//双击添加位移1号曲线数据点
void Qt_excel_test::DoubleClickedPlotDis(QMouseEvent *event)
{
	QPoint point = event->pos();
	//获取鼠标坐标点
	double x_pos = point.x();
	double y_pos = point.y();
	//把鼠标坐标点 转换为 QCustomPlot 内部坐标值 （pixelToCoord 函数）
	//coordToPixel 函数与之相反 是把内部坐标值 转换为外部坐标点
	double x_val = widgetDisplacement->xAxis->pixelToCoord(x_pos);
	double y_val = widgetDisplacement->yAxis->pixelToCoord(y_pos);
	addPoint(GraphDisplacement,x_val, y_val);
	widgetDisplacement->graph(1)->setData(GraphDisplacement->x1, GraphDisplacement->y1);
	widgetDisplacement->replot(QCustomPlot::rpQueuedReplot);
}

//清空位移1号曲线
void Qt_excel_test::on_btnClearPlotDis_Clicked()
{
	time_curr = 0;
	clearGraph_curve1(GraphDisplacement);
	widgetDisplacement->graph(1)->setData(GraphDisplacement->x1, GraphDisplacement->y1);
	widgetDisplacement->replot(QCustomPlot::rpQueuedReplot);
}

//位移曲线 右键菜单 调整范围
void Qt_excel_test::contextMenuRequestDis(QPoint pos)
{
	QMenu *menu = new QMenu(this);
	menu->setAttribute(Qt::WA_DeleteOnClose);
	menu->popup(widgetDisplacement->mapToGlobal(pos));
	menu->addAction("调整范围", GraphDisplacement, &qtGraph::rescaleAxes);
}


/***************力******************/
//读取力0号曲线数据
void Qt_excel_test::on_btnReadDataFor_Clicked()
{
	GraphForce->x0.clear();
	GraphForce->y0.clear();
	GraphForce->x_y0.clear();
	ui.labelDataStatus->setText("正在读取数据");
	QString strFile = QFileDialog::getOpenFileName(this, "选择文件", "./", "文本文件(*.xls;*.xlsx;)");
	if (!strFile.isEmpty())
	{
		QElapsedTimer mstimer;
		mstimer.start();//计时开始

		emit sig_ReadDataDis(strFile, &GraphForce->x_y0);
		excelProssessDlg->exec();//显示正在读写对话框

		float time = (double)mstimer.nsecsElapsed() / (double)1000000;
		QString strTime = QString("%1").arg(time);
		ui.label_time->setText(strTime);
	}
}

//绘制力0号曲线
void Qt_excel_test::on_btnPlotFor_Clicked()
{
	castType->ListListVariant2VectorVector(GraphForce->x_y0, GraphForce->x0, GraphForce->y0);
	widgetForce->graph(0)->setData(GraphForce->x0, GraphForce->y0);
	widgetForce->replot();
	ui.labelDataStatus->setText("绘图完成");
}

//保存力1号曲线数据
void Qt_excel_test::on_btnSaveDataFor_Clicked()
{
	ui.labelDataStatus->setText("正在保存数据");
	QString strFile = QFileDialog::getSaveFileName(this, "另存为", "./", "文本文件(*.xlsx;)");
	if (!strFile.isEmpty())
	{
		QElapsedTimer mstimer;
		mstimer.start();//计时开始

		emit sig_WriteDataDis(&GraphForce->x_y1, strFile);
		excelProssessDlg->exec();//显示正在读写对话框

		float time = (double)mstimer.nsecsElapsed() / (double)1000000;
		QString strTime = QString("%1").arg(time);
		ui.label_time->setText(strTime);//显示时间
	}
}

//双击添加力1号曲线数据点
void Qt_excel_test::DoubleClickedPlotFor(QMouseEvent *event)
{
	QPoint point = event->pos();
	//获取鼠标坐标点
	double x_pos = point.x();
	double y_pos = point.y();
	//把鼠标坐标点 转换为 QCustomPlot 内部坐标值 （pixelToCoord 函数）
	//coordToPixel 函数与之相反 是把内部坐标值 转换为外部坐标点
	double x_val = widgetForce->xAxis->pixelToCoord(x_pos);
	double y_val = widgetForce->yAxis->pixelToCoord(y_pos);
	addPoint(GraphForce, x_val, y_val);
	widgetForce->graph(1)->setData(GraphForce->x1, GraphForce->y1);
	widgetForce->replot(QCustomPlot::rpQueuedReplot);
}

//清空力1号曲线
void Qt_excel_test::on_btnClearPlotFor_Clicked()
{
	time_curr = 0;
	clearGraph_curve1(GraphForce);
	widgetForce->graph(1)->setData(GraphForce->x1, GraphForce->y1);
	widgetForce->replot(QCustomPlot::rpQueuedReplot);
}

//力曲线 右键菜单 调整范围
void Qt_excel_test::contextMenuRequestFor(QPoint pos)
{
	QMenu *menu = new QMenu(this);
	menu->setAttribute(Qt::WA_DeleteOnClose);
	menu->popup(widgetForce->mapToGlobal(pos));
	menu->addAction("调整范围", GraphForce, &qtGraph::rescaleAxes);
}

/***************/
//左键框选放大
void Qt_excel_test::on_act_zoomIn_toggled_(bool arg1)
{
	if (arg1)
	{
		ui.act_move->setChecked(false);//取消拖动选项
		widgetDisplacement->setInteraction(QCP::iRangeDrag, false);//取消拖动
		widgetDisplacement->setSelectionRectMode(QCP::SelectionRectMode::srmZoom);
		widgetForce->setInteraction(QCP::iRangeDrag, false);//取消拖动
		widgetForce->setSelectionRectMode(QCP::SelectionRectMode::srmZoom);
	}
	else
	{
		widgetDisplacement->setSelectionRectMode(QCP::SelectionRectMode::srmNone);
		widgetForce->setSelectionRectMode(QCP::SelectionRectMode::srmNone);
	}
}

//左键拖动
void Qt_excel_test::on_act_move_toggled_(bool arg1)
{
	if (arg1)
	{
		
		ui.act_zoomIn->setChecked(false);//取消框选放大选项
		widgetDisplacement->setSelectionRectMode(QCP::SelectionRectMode::srmNone);
		widgetDisplacement->setInteraction(QCP::iRangeDrag, true);//使能拖动
		widgetForce->setSelectionRectMode(QCP::SelectionRectMode::srmNone);
		widgetForce->setInteraction(QCP::iRangeDrag, true);//使能拖动
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
		if (item->selected() || graph->selected())//选中了哪条曲线或者曲线的图例
		{
			tracerGraph = graph;
		}
		
	}
}

void Qt_excel_test::warnWriteExcelIsEmpty()
{
	QMessageBox::warning(this, "警告", "保存曲线为空，无法保存！");
}

void Qt_excel_test::readExcelDown()
{
	ui.labelDataStatus->setText("*****读取数据完成*****");
	excelProssessDlg->close();
}
void Qt_excel_test::writeExcelDown()
{
	ui.labelDataStatus->setText("*****保存数据完成*****");
	excelProssessDlg->close();
}





