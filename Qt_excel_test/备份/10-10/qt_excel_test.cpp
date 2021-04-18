#include "qt_excel_test.h"
#include "qcustomplot.h"
#include <QAxObject>
#include <QDebug>
#include <iostream>
#include "exceldlg.h"

Qt_excel_test::Qt_excel_test(QWidget *parent)
    : QMainWindow(parent)
{
    ui.setupUi(this);


	this->setWindowTitle("QT读取EXCEL绘图测试");//设置窗口标题

	/************工具栏************/
	ui.act_move->setChecked(true);//默认鼠标左键为拖动

	/************接收信号模块************/
	btnTimerOnCount = 0;

	/*************计时器模块***************/
	m_Timer = new QTimer(this);//计时器
	m_Timer->setInterval(1);//设置定时周期，单位：毫秒
	timecount = 0;
	//QTime m_TimeCounter;//时钟

	/*EXCEL*/
	myexcel = new qtexcel;
	/*位移曲线窗口初始化*/
	widgetDisplacement = ui.widget_displacement;
	GraphDisplacement = new qtGraph(widgetDisplacement);
	GraphDisplacement->setxAxisRange(0, 210);
	GraphDisplacement->setyAxisRange(-0.15, 0.15);
	GraphDisplacement->setxAxisName("时间/s");
	GraphDisplacement->setyAxisName("位移/m");
	GraphDisplacement->setGraphName("目标曲线", "接收曲线");
	/*力曲线窗口初始化*/
	widgetForce = ui.widget_force;
	GraphForce = new qtGraph(widgetForce);
	GraphForce->setxAxisRange(0, 210);
	GraphForce->setyAxisRange(-0.15, 0.15);
	GraphForce->setxAxisName("时间/s");
	GraphForce->setyAxisName("力/N");
	GraphForce->setGraphName("目标曲线", "接收曲线");

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
	connect(widgetDisplacement, &QCustomPlot::mouseDoubleClick, this, &Qt_excel_test::DoubleClickedPlotDis);
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
	connect(widgetForce, &QCustomPlot::mouseDoubleClick, this, &Qt_excel_test::DoubleClickedPlotFor);
	//清除位移1号曲线
	connect(ui.btnClearPlotForce, &QPushButton::clicked, this, &Qt_excel_test::on_btnClearPlotFor_Clicked);
	//位移曲线 右键 菜单栏
	connect(widgetForce, &QCustomPlot::customContextMenuRequested, this, &Qt_excel_test::contextMenuRequestFor);
	//工具栏 游标
	connect(ui.act_tracer, &QAction::toggled, GraphForce, &qtGraph::on_act_tracer_toggled_);
	connect(widgetForce, &QCustomPlot::mouseMove, GraphForce, &qtGraph::myMouseMoveEvent);

	//工具栏 框选放大
	connect(ui.act_zoomIn, &QAction::toggled, this, &Qt_excel_test::on_act_zoomIn_toggled_);
	//工具栏 拖动
	connect(ui.act_move, &QAction::toggled, this, &Qt_excel_test::on_act_move_toggled_);

	connect(widgetDisplacement, &QCustomPlot::selectionChangedByUser, this, &Qt_excel_test::traceGraphSelect);//双击图例
}

//析构
 Qt_excel_test::~Qt_excel_test() 
 {
	 delete myexcel;
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


 void Qt_excel_test::on_timer_timeout()
 {
	 QList<QVariant> xy;
	 double x_temp = timecount / 1000;
	 double y_temp = (timecount / 1000)*(timecount / 1000);
	 addPoint(GraphDisplacement,x_temp, y_temp);
	 widgetDisplacement->graph(1)->setData(GraphForce->x1, GraphForce->y1);
	 widgetDisplacement->replot(QCustomPlot::rpQueuedReplot);
	 timecount++;
 }

 //计时器信号触发
 void Qt_excel_test::on_btnTimerOn_Clicked()
 {
	 if (btnTimerOnCount % 2 == 0)
	 {
		 m_Timer->start();
	 }
	 else
	 {
		 m_Timer->stop();
	 }
	 btnTimerOnCount++;
 }

 /***************位移******************/
//读取位移0号曲线数据
void Qt_excel_test::on_btnReadDataDis_Clicked()
{
	/*开启一条多线程*/
	qDebug() << tr("开启线程");
	firstThread = new QThread;//线程容器
	myTask = new MyWorkThread;
	myTask->moveToThread(firstThread); //将创建的对象移到线程容器中
	connect(firstThread, &QThread::finished, firstThread, &QObject::deleteLater); //终止线程时要调用deleteLater槽函数
	connect(firstThread, &QThread::finished, myTask, &QObject::deleteLater); //终止线程时要调用deleteLater槽函数
	connect(firstThread, &QThread::started, myTask, &MyWorkThread::startReadDataDis);  //开启线程槽函数
	connect(firstThread, SIGNAL(finished()), this, SLOT(finishedThreadSlot()));
	firstThread->start();                                                           //开启多线程槽函数
	qDebug() << "mainWidget QThread::currentThreadId()==" << QThread::currentThreadId();


	/*
	//ExcelDlg dlg(this);
	//dlg.exec();
	GraphDisplacement->x0.clear();
	GraphDisplacement->y0.clear();
	GraphDisplacement->x_y0.clear();
	ui.labelDataStatus->setText("正在读取数据");
	QString strFile = QFileDialog::getOpenFileName(this, "选择文件", "./", "文本文件(*.xls;*.xlsx;)");
	if (!strFile.isEmpty())
	{
		QElapsedTimer mstimer;
		mstimer.start();//计时开始

		myexcel->readExcelFast(strFile, GraphDisplacement->x0, GraphDisplacement->y0);

		float time = (double)mstimer.nsecsElapsed() / (double)1000000;
		QString strTime = QString("%1").arg(time);
		ui.label_time->setText(strTime);
	*/

	GraphDisplacement->x0.clear();
	GraphDisplacement->y0.clear();
	GraphDisplacement->x_y0.clear();
	ui.labelDataStatus->setText("正在读取数据");
	QString strFile = QFileDialog::getOpenFileName(this, "选择文件", "./", "文本文件(*.xls;*.xlsx;)");
	if (!strFile.isEmpty())
	{
		QElapsedTimer mstimer;
		mstimer.start();//计时开始

		myexcel->readExcelFast(strFile, GraphDisplacement->x0, GraphDisplacement->y0);

		float time = (double)mstimer.nsecsElapsed() / (double)1000000;
		QString strTime = QString("%1").arg(time);
		ui.label_time->setText(strTime);
	}
}

//绘制位移0号曲线
void Qt_excel_test::on_btnPlotDis_Clicked()
{
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

		myexcel->writeExcelFast(strFile, GraphDisplacement->x_y1);

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
	GraphDisplacement->x1.clear();
	GraphDisplacement->y1.clear();
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

		myexcel->readExcelFast(strFile, GraphForce->x0, GraphForce->y0);

		float time = (double)mstimer.nsecsElapsed() / (double)1000000;
		QString strTime = QString("%1").arg(time);
		ui.label_time->setText(strTime);
	}
}

//绘制力0号曲线
void Qt_excel_test::on_btnPlotFor_Clicked()
{
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

		myexcel->writeExcelFast(strFile, GraphForce->x_y1);

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
	GraphForce->x1.clear();
	GraphForce->y1.clear();
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



