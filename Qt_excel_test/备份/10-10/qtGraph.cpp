#include "qtGraph.h"

qtGraph::qtGraph(QCustomPlot * Plot) 
{
	myPlot = Plot;
	tracerEnable = false;
	/*************绘图模块***************/
	//设定右上角图例标注可见
	myPlot->legend->setVisible(true);
	//设定右上角图例标注的字体
	myPlot->legend->setFont(QFont("Helvetica", 9));
	//添加图形
	myPlot->addGraph();
	myPlot->addGraph();
	//曲线全部可见
	myPlot->graph(0)->rescaleAxes();
	myPlot->graph(1)->rescaleAxes(true);
	//设置画笔
	QPen pen1(Qt::red, 1, Qt::SolidLine, Qt::RoundCap, Qt::RoundJoin);
	myPlot->graph(0)->setPen(QPen(Qt::blue));
	myPlot->graph(1)->setPen(pen1);
	//设置线型
	myPlot->graph(0)->setLineStyle(QCPGraph::lsLine);
	myPlot->graph(1)->setLineStyle(QCPGraph::lsLine);
	//设置线上点的风格
	myPlot->graph(0)->setScatterStyle(QCPScatterStyle::ssNone);
	myPlot->graph(1)->setScatterStyle(QCPScatterStyle::ssCircle);
	//myPlot->graph(1)->setAdaptiveSampling(true);//自适应采样
	//myPlot->graph(0)->rescaleAxes(true);//若没设置坐标显示范围 则自动调整范围是图像恰好全部显示出来
	//右键菜单自定义
	myPlot->setContextMenuPolicy(Qt::CustomContextMenu);
	/**********鼠标操作图像模块************/
	myPlot->selectionRect()->setPen(QPen(Qt::black, 1, Qt::DashLine));//设置选框的样式：虚线
	myPlot->selectionRect()->setBrush(QBrush(QColor(0, 0, 100, 50)));//设置选框的样式：半透明浅蓝
	myPlot->setInteraction(QCP::iRangeDrag, true); //鼠标单击拖动 QCPAxisRect::mousePressEvent() 左键拖动
	myPlot->setInteraction(QCP::iRangeZoom, true); //滚轮滑动缩放
	myPlot->setInteraction(QCP::iSelectAxes, false); //true曲线可选
	myPlot->setInteraction(QCP::iSelectLegend, true); //图例可选
}

qtGraph::~qtGraph()
{
}

//X轴名称
void qtGraph::setxAxisName(QString  name)
{
	myPlot->xAxis->setLabel(name);
}
//Y轴名称
void qtGraph::setyAxisName(QString  name)
{
	myPlot->yAxis->setLabel(name);
}
//X轴范围
void qtGraph::setxAxisRange(double lower, double upper)
{
	myPlot->xAxis->setRange(lower, upper);
}
//Y轴范围
void qtGraph::setyAxisRange(double lower, double upper)
{
	myPlot->yAxis->setRange(lower, upper);
}
//曲线图例名称
void qtGraph::setGraphName(QString name0, QString name1)
{
	//设置右上角图形标注名称
	myPlot->graph(0)->setName(name0);
	myPlot->graph(1)->setName(name1);
}


void qtGraph::setVisibleTracerFalse()
{
	tracer0->setVisible(false);
	tracer1->setVisible(false);
	tracerX0Label->setVisible(false);
	tracerY0Label->setVisible(false);
	tracerY1Label->setVisible(false);
	myPlot->replot(QCustomPlot::rpQueuedReplot); //刷新图标，不能省略
}

void qtGraph::setVisibleTracerTrue()
{
	tracer0->setVisible(true);
	tracer1->setVisible(true);
	tracerX0Label->setVisible(true);
	tracerY0Label->setVisible(true);
	tracerY1Label->setVisible(true);
	myPlot->replot(QCustomPlot::rpQueuedReplot); //刷新图标，不能省略
}


/***************槽函数*****************/
//开启/关闭游标
void qtGraph::on_act_tracer_toggled_(bool arg1)
{
	if (arg1)
	{
		tracerEnable = true;
		tracer0 = new QCPItemTracer(myPlot);
		tracer0->setStyle(QCPItemTracer::tsCrosshair);//游标样式：十字星、圆圈、方框
		tracer0->setPen(QPen(Qt::green));//设置tracer的颜色绿色
		tracer0->setPen(QPen(Qt::DashLine));//虚线游标
		tracer0->setBrush(QBrush(Qt::red));
		tracer0->setSize(10);
		tracer0->setInterpolating(true);//false禁用插值

		tracer1 = new QCPItemTracer(myPlot);
		tracer1->setStyle(QCPItemTracer::tsCrosshair);//游标样式：十字星、圆圈、方框
		tracer1->setPen(QPen(Qt::green));//设置tracer的颜色绿色
		tracer1->setPen(QPen(Qt::DashLine));//虚线游标
		tracer1->setBrush(QBrush(Qt::red));
		tracer1->setSize(10);
		tracer1->setInterpolating(true);//false禁用插值


		tracerX0Label = new QCPItemText(myPlot);
		tracerX0Label->setClipToAxisRect(false);
		tracerX0Label->setLayer("overlay");
		tracerX0Label->setPen(QPen(Qt::green));
		tracerX0Label->setFont(QFont("宋体", 10));
		tracerX0Label->setPadding(QMargins(5, 5, 5, 5));
		tracerX0Label->setPositionAlignment(Qt::AlignLeft | Qt::AlignTop);
		//下面这个语句很重要，它将游标说明锚固在tracer位置处，实现自动跟随
		tracerX0Label->position->setType(QCPItemPosition::ptAxisRectRatio);//位置类型（当前轴范围的比例为单位/实际坐标为单位）
		tracerX0Label->position->setParentAnchorX(tracer0->position);

		tracerY0Label = new QCPItemText(myPlot);
		tracerY0Label->setClipToAxisRect(false);
		tracerY0Label->setLayer("overlay");
		tracerY0Label->setPen(QPen(Qt::red));
		tracerY0Label->setFont(QFont("宋体", 10));
		tracerY0Label->setPadding(QMargins(5, 5, 5, 5));
		tracerY0Label->setPositionAlignment(Qt::AlignLeft | Qt::AlignTop);
		//下面这个语句很重要，它将游标说明锚固在tracer位置处，实现自动跟随
		tracerY0Label->position->setType(QCPItemPosition::ptAxisRectRatio);//位置类型（当前轴范围的比例为单位/实际坐标为单位）
		tracerY0Label->position->setParentAnchorY(tracer0->position);

		tracerY1Label = new QCPItemText(myPlot);
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
		//delete tracer0;
		//delete tracer1;
		//delete tracerX0Label;
		//delete tracerY0Label;
		//delete tracerY1Label;
		setVisibleTracerFalse();

	}
}
//鼠标移动事件
void qtGraph::myMouseMoveEvent(QMouseEvent* event)
{
	if (tracerEnable)//游标使能判断
	{
		double x = myPlot->xAxis->pixelToCoord(event->pos().x());//鼠标点的像素坐标转plot坐标
		
		foundRange = true;
		QrangeX0 = myPlot->graph(0)->getKeyRange(foundRange, QCP::sdBoth);//获取0号曲线X轴坐标范围
		QrangeX0_lower = QrangeX0.lower;
		QrangeX0_upper = QrangeX0.upper;

		QrangeX1 = myPlot->graph(1)->getKeyRange(foundRange, QCP::sdBoth);//获取1号曲线X轴坐标范围
		QrangeX1_lower = QrangeX1.lower;
		QrangeX1_upper = QrangeX1.upper;
		//如果鼠标移动超出0号曲线X轴范围，则0号曲线隐藏游标
		if (x < QrangeX0_upper && x > QrangeX0_lower)
		{
			tracer0->setGraph(myPlot->graph(0));//设置游标吸附在traceGraph这条曲线上
			tracer0->setGraphKey(x);//将游标横坐标（key）设置成刚获得的横坐标数据x （这就是游标随动的关键代码）
			tracer0->updatePosition(); //使得刚设置游标的横纵坐标位置生效
			double traceX0 = tracer0->position->key();
			double traceY0 = tracer0->position->value();
			
			tracerX0Label->setText(QString::number(traceX0, 'f', 3));//游标文本框，指示游标的X值
			tracerY0Label->setText(QString::number(traceY0, 'f', 3));//游标文本框，指示游标的Y值
			
			tracer0->setVisible(true);
			tracerX0Label->setVisible(true);
			tracerY0Label->setVisible(true);	
		}
		else
		{
			tracer0->setVisible(false);
			tracerX0Label->setVisible(false);
			tracerY0Label->setVisible(false);
		}
		//如果鼠标移动超出1号曲线X轴范围，则1号曲线隐藏游标
		if (x < QrangeX1_upper && x > QrangeX1_lower)
		{
			double traceY1 = tracer1->position->value();
			tracer1->setGraph(myPlot->graph(1));//设置游标吸附在traceGraph这条曲线上
			tracer1->setGraphKey(x);
			tracer1->updatePosition(); //使得刚设置游标的横纵坐标位置生效

			tracerY1Label->setText(QString::number(traceY1, 'f', 3));//游标文本框，指示游标的Y值

			tracer1->setVisible(true);
			tracerY1Label->setVisible(true);
		}
		else
		{
			tracer1->setVisible(false);
			tracerY1Label->setVisible(false);
		}
		myPlot->replot(QCustomPlot::rpQueuedReplot); //刷新图标，不能省略
	}
}
////曲线全部显示
void qtGraph::rescaleAxes()
{
	//给第一个graph设置rescaleAxes()，后续所有graph都设置rescaleAxes(true)即可实现显示所有曲线。
	/*rescaleAxes(true)时如果plot的X或Y轴本来能容纳下本graph的X或Y数据点，
	那么plot的X或Y轴的可视范围就无需调整，只有plot容纳不下本graph时，才扩展plot两个轴的显示范围。*/
	myPlot->graph(0)->rescaleAxes();
	myPlot->graph(1)->rescaleAxes(true);
	myPlot->replot();
}
