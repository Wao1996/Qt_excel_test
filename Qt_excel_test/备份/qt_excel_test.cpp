#include "qt_excel_test.h"
#include "qcustomplot.h"
#include <QAxObject>

Qt_excel_test::Qt_excel_test(QWidget *parent)
    : QMainWindow(parent)
{
    ui.setupUi(this);
	this->setWindowTitle("QT读取EXCEL绘图测试");//设置窗口标题

	/************接收信号模块************/
	btnTimerOnCount = 0;
	/*************计时器模块***************/
	m_Timer = new QTimer(this);//计时器
	m_Timer->setInterval(1);//设置定时周期，单位：毫秒
	timecount = 0;
	//QTime m_TimeCounter;//时钟

	/**********鼠标操作图像模块************/
	ui.widget_my->selectionRect()->setPen(QPen(Qt::black, 1, Qt::DashLine));//设置选框的样式：虚线
	ui.widget_my->selectionRect()->setBrush(QBrush(QColor(0, 0, 100, 50)));//设置选框的样式：半透明浅蓝
	//ui.widget_my->setSelectionRectMode(QCP::SelectionRectMode::srmZoom);
	ui.act_move->setChecked(true);//默认鼠标左键为拖动
	ui.widget_my->setInteraction(QCP::iRangeDrag, true); //鼠标单击拖动 QCPAxisRect::mousePressEvent() 左键拖动
	ui.widget_my->setInteraction(QCP::iRangeZoom, true); //滚轮滑动缩放
	ui.widget_my->setInteraction(QCP::iSelectAxes, false); //曲线可选
	ui.widget_my->setInteraction(QCP::iSelectLegend, true); //图例可选

	tracerEnable = false;
	/*************绘图模块***************/
	//设置坐标轴参数
	ui.widget_my->xAxis->setLabel("x轴");
	ui.widget_my->xAxis->setRange(0, 210);
	ui.widget_my->yAxis->setLabel("y轴");
	ui.widget_my->yAxis->setRange(-0.15, 0.15);
	//设定右上角图形标注可见
	ui.widget_my->legend->setVisible(true);
	//设定右上角图形标注的字体
	ui.widget_my->legend->setFont(QFont("Helvetica", 9));
	//添加图形
	ui.widget_my->addGraph();
	ui.widget_my->addGraph();
	//曲线全部可见
	ui.widget_my->graph(0)->rescaleAxes();
	ui.widget_my->graph(1)->rescaleAxes(true);
	//设置画笔
	QPen pen1(Qt::red, 1, Qt::SolidLine, Qt::RoundCap, Qt::RoundJoin);
	ui.widget_my->graph(0)->setPen(QPen(Qt::blue));
	ui.widget_my->graph(1)->setPen(pen1);
	//设置线型
	ui.widget_my->graph(0)->setLineStyle(QCPGraph::lsLine);
	ui.widget_my->graph(1)->setLineStyle(QCPGraph::lsLine);
	//设置线上点的风格
	ui.widget_my->graph(0)->setScatterStyle(QCPScatterStyle::ssNone);
	ui.widget_my->graph(1)->setScatterStyle(QCPScatterStyle::ssCircle); 
	//设置右上角图形标注名称
	ui.widget_my->graph(0)->setName("目标曲线");
	ui.widget_my->graph(1)->setName("接收曲线");

	//ui.widget_my->graph(1)->setAdaptiveSampling(true);//自适应采样
	//ui.widget_my->graph(0)->rescaleAxes(true);//若没设置坐标显示范围 则自动调整范围是图像恰好全部显示出来
	//右键菜单自定义
	ui.widget_my->setContextMenuPolicy(Qt::CustomContextMenu);
	/*************信号槽连接***************/
	connect(ui.btnOpenData, &QPushButton::clicked, this, &Qt_excel_test::on_btnReadData_Clicked);
	connect(ui.btnPlot, &QPushButton::clicked, this, &Qt_excel_test::on_btnPlot_Clicked);
	connect(m_Timer, &QTimer::timeout, this, &Qt_excel_test::on_timer_timeout);
	connect(ui.btnSaveData, &QPushButton::clicked, this, &Qt_excel_test::on_btnSaveData_Clicked);
	connect(ui.btnReceive, &QPushButton::clicked, this, &Qt_excel_test::on_btnReceive_Clicked);
	connect(ui.btnTimerOn, &QPushButton::clicked, this, &Qt_excel_test::on_btnTimerOn_Clicked);
	connect(ui.btnClearPlot1, &QPushButton::clicked, this, &Qt_excel_test::on_btnClearPlot1_Clicked);//清除曲线1
	connect(ui.widget_my, &QCustomPlot::mouseDoubleClick, this, &Qt_excel_test::DoubleClickedPlot);//双击添加数据点
	connect(ui.widget_my, &QCustomPlot::customContextMenuRequested, this, &Qt_excel_test::contextMenuRequest);//右键 菜单栏
	connect(ui.act_zoomIn, &QAction::toggled, this, &Qt_excel_test::on_act_zoomIn_toggled_);//工具栏 框选放大
	connect(ui.act_move, &QAction::toggled, this, &Qt_excel_test::on_act_move_toggled_);//工具栏 拖动
	connect(ui.act_tracer, &QAction::toggled, this, &Qt_excel_test::on_act_tracer_toggled_);//工具栏 游标
	connect(ui.widget_my, &QCustomPlot::mouseMove, this, &Qt_excel_test::myMouseMoveEvent);
	connect(ui.widget_my, &QCustomPlot::selectionChangedByUser, this, &Qt_excel_test::traceGraphSelect);//双击图例
}

 Qt_excel_test::~Qt_excel_test() 
 {
	 delete m_Timer;
 }

//读取数据
void Qt_excel_test::on_btnReadData_Clicked()
{
	x.clear();
	y.clear();
	x_y.clear();
	ui.labelDataStatus->setText("正在读取数据");
	QString strFile = QFileDialog::getOpenFileName(this, "选择文件", "./", "文本文件(*.xls;*.xlsx;)");
	if (!strFile.isEmpty())
	{
		QElapsedTimer mstimer;
		mstimer.start();//计时开始

		myexcel.readExcelFast(strFile,x,y);

		float time = (double)mstimer.nsecsElapsed() / (double)1000000;
		QString strTime = QString("%1").arg(time);
		ui.label_time->setText(strTime);
	}
}

//绘制曲线
void Qt_excel_test::on_btnPlot_Clicked()
{
	//m_Timer->start();
	ui.widget_my->graph(0)->setData(x, y);
	ui.widget_my->replot();
	ui.labelDataStatus->setText("绘图完成");
}

void Qt_excel_test::on_timer_timeout()
{
	QList<QVariant> xy;
	double x_temp = timecount/1000;
	double y_temp = (timecount/1000)*(timecount / 1000);
	addPoint(x_temp, y_temp);
	ui.widget_my->graph(1)->setData(x, y);
	ui.widget_my->replot(QCustomPlot::rpQueuedReplot);
	timecount++;
}

//保存数据
void Qt_excel_test::on_btnSaveData_Clicked()
{
	ui.labelDataStatus->setText("正在保存数据");
	QString strFile = QFileDialog::getSaveFileName(this, "另存为", "./", "文本文件(*.xlsx;)");
	if (!strFile.isEmpty())
	{
		QElapsedTimer mstimer;
		mstimer.start();//计时开始

		myexcel.writeExcelFast(strFile, x_y);

		float time = (double)mstimer.nsecsElapsed() / (double)1000000;
		QString strTime = QString("%1").arg(time);
		ui.label_time->setText(strTime);//显示时间
	}
}

//接收写入的信号
void Qt_excel_test::on_btnReceive_Clicked()
{
	double x_temp = ui.lineEdit_X->text().toDouble();
	double y_temp = ui.lineEdit_Y->text().toDouble();
	addPoint(x_temp, y_temp);
	ui.widget_my->graph(1)->setData(x, y);
	ui.widget_my->replot(QCustomPlot::rpQueuedReplot);
}

void Qt_excel_test::on_btnTimerOn_Clicked()
{
	if (btnTimerOnCount%2 == 0)
	{
		m_Timer->start();
	} 
	else
	{
		m_Timer->stop();
	}
	btnTimerOnCount++;
}

//清空1号曲线
void Qt_excel_test::on_btnClearPlot1_Clicked()
{
	x.clear();
	y.clear();
	ui.widget_my->graph(1)->setData(x, y);
	ui.widget_my->replot(QCustomPlot::rpQueuedReplot);
}

//双击添加数据点
void Qt_excel_test::DoubleClickedPlot(QMouseEvent *event)
{
	QPoint point = event->pos();
	//获取鼠标坐标点
	double x_pos = point.x();
	double y_pos = point.y();
	//把鼠标坐标点 转换为 QCustomPlot 内部坐标值 （pixelToCoord 函数）
	//coordToPixel 函数与之相反 是把内部坐标值 转换为外部坐标点
	double x_val = ui.widget_my->xAxis->pixelToCoord(x_pos);
	double y_val = ui.widget_my->yAxis->pixelToCoord(y_pos);
	addPoint(x_val, y_val);
	ui.widget_my->graph(1)->setData(x, y);
	ui.widget_my->replot(QCustomPlot::rpQueuedReplot);
}

//添加数据点
void Qt_excel_test::addPoint(const double &x_temp, const double &y_temp)
{
	QList<QVariant> xy;
	x.append(x_temp);
	y.append(y_temp);
	
	xy.append(x_temp);
	xy.append(y_temp);
	x_y.append(xy);
}

//右键菜单 调整范围
void Qt_excel_test::contextMenuRequest(QPoint pos)
{
	QMenu *menu = new QMenu(this);
	menu->setAttribute(Qt::WA_DeleteOnClose);
	menu->popup(ui.widget_my->mapToGlobal(pos));
	menu->addAction("调整范围", this, &Qt_excel_test::rescaleAxes);
}

//曲线全部显示
void Qt_excel_test::rescaleAxes()
{
	//给第一个graph设置rescaleAxes()，后续所有graph都设置rescaleAxes(true)即可实现显示所有曲线。
	/*rescaleAxes(true)时如果plot的X或Y轴本来能容纳下本graph的X或Y数据点，
	那么plot的X或Y轴的可视范围就无需调整，只有plot容纳不下本graph时，才扩展plot两个轴的显示范围。*/
	ui.widget_my->graph(0)->rescaleAxes();
	ui.widget_my->graph(1)->rescaleAxes(true);
	ui.widget_my->replot();
}

void Qt_excel_test::setVisibleTracerFalse()
{
	tracer->setVisible(false);
	tracer1->setVisible(false);
	tracerXLabel->setVisible(false);
	tracerYLabel->setVisible(false);
	tracerY1Label->setVisible(false);
	ui.widget_my->replot(QCustomPlot::rpQueuedReplot); //刷新图标，不能省略
}

void Qt_excel_test::setVisibleTracerTrue()
{
	tracer->setVisible(true);
	tracer1->setVisible(true);
	tracerXLabel->setVisible(true);
	tracerYLabel->setVisible(true);
	tracerY1Label->setVisible(true);
	ui.widget_my->replot(QCustomPlot::rpQueuedReplot); //刷新图标，不能省略
}

//左键框选放大
void Qt_excel_test::on_act_zoomIn_toggled_(bool arg1)
{
	if (arg1)
	{
		ui.act_move->setChecked(false);//取消拖动选项
		ui.widget_my->setInteraction(QCP::iRangeDrag, false);//取消拖动
		ui.widget_my->setSelectionRectMode(QCP::SelectionRectMode::srmZoom);
	}
	else
	{
		ui.widget_my->setSelectionRectMode(QCP::SelectionRectMode::srmNone);
	}
}

//左键拖动
void Qt_excel_test::on_act_move_toggled_(bool arg1)
{
	if (arg1)
	{
		ui.act_zoomIn->setChecked(false);//取消框选放大选项
		ui.widget_my->setSelectionRectMode(QCP::SelectionRectMode::srmNone);
		ui.widget_my->setInteraction(QCP::iRangeDrag, true);//使能拖动
	}
	else
	{
		ui.widget_my->setInteraction(QCP::iRangeDrag, false);
	}
}

//开启游标
void Qt_excel_test::on_act_tracer_toggled_(bool arg1)
{
	if (arg1)
	{
		tracerEnable = true;
		//this->setMouseTracking(true);
		tracer = new QCPItemTracer(ui.widget_my);
		tracer->setStyle(QCPItemTracer::tsCrosshair);//游标样式：十字星、圆圈、方框
		tracer->setPen(QPen(Qt::green));//设置tracer的颜色绿色
		tracer->setPen(QPen(Qt::DashLine));//虚线游标
		tracer->setBrush(QBrush(Qt::red));
		tracer->setSize(10);
		tracer->setInterpolating(true);//false禁用插值

		tracer1 = new QCPItemTracer(ui.widget_my);
		tracer1->setStyle(QCPItemTracer::tsCrosshair);//游标样式：十字星、圆圈、方框
		tracer1->setPen(QPen(Qt::green));//设置tracer的颜色绿色
		tracer1->setPen(QPen(Qt::DashLine));//虚线游标
		tracer1->setBrush(QBrush(Qt::red));
		tracer1->setSize(10);
		tracer1->setInterpolating(true);//false禁用插值


		tracerXLabel = new QCPItemText(ui.widget_my);
		tracerXLabel->setClipToAxisRect(false);
		tracerXLabel->setLayer("overlay");
		tracerXLabel->setPen(QPen(Qt::green));
		tracerXLabel->setFont(QFont("宋体", 10));
		tracerXLabel->setPadding(QMargins(5, 5, 5, 5));
		tracerXLabel->setPositionAlignment(Qt::AlignLeft | Qt::AlignTop);
		//下面这个语句很重要，它将游标说明锚固在tracer位置处，实现自动跟随
		tracerXLabel->position->setType(QCPItemPosition::ptAxisRectRatio);//位置类型（当前轴范围的比例为单位/实际坐标为单位）
		tracerXLabel->position->setParentAnchorX(tracer->position);
		
		tracerYLabel = new QCPItemText(ui.widget_my);
		tracerYLabel->setClipToAxisRect(false);
		tracerYLabel->setLayer("overlay");
		tracerYLabel->setPen(QPen(Qt::red));
		tracerYLabel->setFont(QFont("宋体", 10));
		tracerYLabel->setPadding(QMargins(5, 5, 5, 5));
		tracerYLabel->setPositionAlignment(Qt::AlignLeft | Qt::AlignTop);
		//下面这个语句很重要，它将游标说明锚固在tracer位置处，实现自动跟随
		tracerYLabel->position->setType(QCPItemPosition::ptAxisRectRatio);//位置类型（当前轴范围的比例为单位/实际坐标为单位）
		tracerYLabel->position->setParentAnchorY(tracer->position);

		tracerY1Label = new QCPItemText(ui.widget_my);
		tracerY1Label->setClipToAxisRect(false);
		tracerY1Label->setLayer("overlay");
		tracerY1Label->setPen(QPen(Qt::red));
		tracerY1Label->setFont(QFont("宋体", 10));
		tracerY1Label->setPadding(QMargins(5, 5, 5, 5));
		tracerY1Label->setPositionAlignment(Qt::AlignLeft | Qt::AlignTop);
		//下面这个语句很重要，它将游标说明锚固在tracer位置处，实现自动跟随
		tracerY1Label->position->setType(QCPItemPosition::ptAxisRectRatio);//位置类型（当前轴范围的比例为单位/实际坐标为单位）
		tracerY1Label->position->setParentAnchorY(tracer1->position);
	} 
	else
	{
		tracerEnable = false;
		setVisibleTracerFalse();
		
	}
}

void Qt_excel_test::traceGraphSelect()
{
	
	for (int i = 0; i < ui.widget_my->graphCount(); ++i)
	{
		QCPGraph *graph = ui.widget_my->graph(i);
		QCPPlottableLegendItem *item = ui.widget_my->legend->itemWithPlottable(graph);
		if (item->selected() || graph->selected())//选中了哪条曲线或者曲线的图例
		{
			tracerGraph = graph;
		}
		
	}
}


void Qt_excel_test::myMouseMoveEvent(QMouseEvent* event)
{
	if (tracerEnable)//游标使能
	{
		//if (tracerGraph != nullptr)
		//{
		//QCustomPlot::mouseMoveEvent(event);//避免让父类的移动事件失效

		double x = ui.widget_my->xAxis->pixelToCoord(event->pos().x());//鼠标点的像素坐标转plot坐标

		foundRange = true;
		QrangeX = ui.widget_my->graph(0)->getKeyRange(foundRange,QCP::sdBoth);
		QrangeX_lower = QrangeX.lower;
		QrangeX_upper = QrangeX.upper;
		if (x < QrangeX_upper && x > QrangeX_lower)
		{
			tracer->setGraph(ui.widget_my->graph(0));//设置游标吸附在traceGraph这条曲线上
			tracer->setGraphKey(x);//将游标横坐标（key）设置成刚获得的横坐标数据x （这就是游标随动的关键代码）
			tracer->updatePosition(); //使得刚设置游标的横纵坐标位置生效
			double traceX = tracer->position->key();
			double traceY = tracer->position->value();
			double traceY1 = tracer1->position->value();
			tracerXLabel->setText(QString::number(traceX, 'f', 3));//游标文本框，指示游标的X值
			tracerYLabel->setText(QString::number(traceY, 'f', 3));//游标文本框，指示游标的Y值
			tracer1->setGraph(ui.widget_my->graph(1));//设置游标吸附在traceGraph这条曲线上
			tracer1->setGraphKey(x);
			tracer1->updatePosition(); //使得刚设置游标的横纵坐标位置生效
			tracerXLabel->setText(QString::number(traceX, 'f', 3));//游标文本框，指示游标的X值
			tracerYLabel->setText(QString::number(traceY, 'f', 3));//游标文本框，指示游标的Y值
			tracerY1Label->setText(QString::number(traceY1, 'f', 3));//游标文本框，指示游标的Y值
			setVisibleTracerTrue();
		}
		else
		{
			setVisibleTracerFalse();
		}
		//}
	}
}
//
//void Qt_excel_test::mouseMoveEvent(QMouseEvent *event)
//{
//	QCustomPlot::mouseMoveEvent(event);//避免让父类的移动事件失效
//
//	if (tracerEnable)//游标使能
//	{
//		if (tracerGraph != NULL)//检查一下游标附着的graph是否还存在
//		{
//			double x = ui.widget_my->xAxis->pixelToCoord(event->pos().x());//鼠标点的像素坐标转plot坐标
//			tracer->setGraph(tracerGraph);//设置游标吸附在traceGraph这条曲线上
//			tracer->setGraphKey(x);//将游标横坐标（key）设置成刚获得的横坐标数据x （这就是游标随动的关键代码）
//			double traceX = tracer->position->key();
//			double traceY = tracer->position->value();
//
//			tracerXLabel->setText(QString::number(traceX));//游标文本框，指示游标的X值
//			//tracerYLabel->setText(QString::number(traceY));//游标文本框，指示游标的Y值
//
//			//计算游标X值对应的所有曲线的Y值
//			/*
//			for (int i = 0; i < ui.widget_my->graphCount(); i++)
//			{
//				//搜索左边离traceX最近的key对应的点，详情参考findBegin函数的帮助
//				QCPDataContainer<QCPGraphData>::const_iterator coorPoint = ui.widget_my->graph(i)->data().data()->findBegin(traceX, true);//true代表向左搜索
//				qDebug() << QString("graph%1对应的Y值是").arg(i) << coorPoint->value;
//			}
//			*/
//			ui.widget_my->replot(QCustomPlot::rpQueuedReplot); //刷新图标，不能省略
//		}
//	}
//}

