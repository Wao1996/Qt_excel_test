#include "qt_excel_test.h"
#include "qcustomplot.h"
#include <QAxObject>

Qt_excel_test::Qt_excel_test(QWidget *parent)
    : QMainWindow(parent)
{
    ui.setupUi(this);
	this->setWindowTitle("QT��ȡEXCEL��ͼ����");//���ô��ڱ���

	/************�����ź�ģ��************/
	btnTimerOnCount = 0;
	/*************��ʱ��ģ��***************/
	m_Timer = new QTimer(this);//��ʱ��
	m_Timer->setInterval(1);//���ö�ʱ���ڣ���λ������
	timecount = 0;
	//QTime m_TimeCounter;//ʱ��

	/**********������ͼ��ģ��************/
	ui.widget_my->selectionRect()->setPen(QPen(Qt::black, 1, Qt::DashLine));//����ѡ�����ʽ������
	ui.widget_my->selectionRect()->setBrush(QBrush(QColor(0, 0, 100, 50)));//����ѡ�����ʽ����͸��ǳ��
	//ui.widget_my->setSelectionRectMode(QCP::SelectionRectMode::srmZoom);
	ui.act_move->setChecked(true);//Ĭ��������Ϊ�϶�
	ui.widget_my->setInteraction(QCP::iRangeDrag, true); //��굥���϶� QCPAxisRect::mousePressEvent() ����϶�
	ui.widget_my->setInteraction(QCP::iRangeZoom, true); //���ֻ�������
	ui.widget_my->setInteraction(QCP::iSelectAxes, false); //���߿�ѡ
	ui.widget_my->setInteraction(QCP::iSelectLegend, true); //ͼ����ѡ

	tracerEnable = false;
	/*************��ͼģ��***************/
	//�������������
	ui.widget_my->xAxis->setLabel("x��");
	ui.widget_my->xAxis->setRange(0, 210);
	ui.widget_my->yAxis->setLabel("y��");
	ui.widget_my->yAxis->setRange(-0.15, 0.15);
	//�趨���Ͻ�ͼ�α�ע�ɼ�
	ui.widget_my->legend->setVisible(true);
	//�趨���Ͻ�ͼ�α�ע������
	ui.widget_my->legend->setFont(QFont("Helvetica", 9));
	//���ͼ��
	ui.widget_my->addGraph();
	ui.widget_my->addGraph();
	//����ȫ���ɼ�
	ui.widget_my->graph(0)->rescaleAxes();
	ui.widget_my->graph(1)->rescaleAxes(true);
	//���û���
	QPen pen1(Qt::red, 1, Qt::SolidLine, Qt::RoundCap, Qt::RoundJoin);
	ui.widget_my->graph(0)->setPen(QPen(Qt::blue));
	ui.widget_my->graph(1)->setPen(pen1);
	//��������
	ui.widget_my->graph(0)->setLineStyle(QCPGraph::lsLine);
	ui.widget_my->graph(1)->setLineStyle(QCPGraph::lsLine);
	//�������ϵ�ķ��
	ui.widget_my->graph(0)->setScatterStyle(QCPScatterStyle::ssNone);
	ui.widget_my->graph(1)->setScatterStyle(QCPScatterStyle::ssCircle); 
	//�������Ͻ�ͼ�α�ע����
	ui.widget_my->graph(0)->setName("Ŀ������");
	ui.widget_my->graph(1)->setName("��������");

	//ui.widget_my->graph(1)->setAdaptiveSampling(true);//����Ӧ����
	//ui.widget_my->graph(0)->rescaleAxes(true);//��û����������ʾ��Χ ���Զ�������Χ��ͼ��ǡ��ȫ����ʾ����
	//�Ҽ��˵��Զ���
	ui.widget_my->setContextMenuPolicy(Qt::CustomContextMenu);
	/*************�źŲ�����***************/
	connect(ui.btnOpenData, &QPushButton::clicked, this, &Qt_excel_test::on_btnReadData_Clicked);
	connect(ui.btnPlot, &QPushButton::clicked, this, &Qt_excel_test::on_btnPlot_Clicked);
	connect(m_Timer, &QTimer::timeout, this, &Qt_excel_test::on_timer_timeout);
	connect(ui.btnSaveData, &QPushButton::clicked, this, &Qt_excel_test::on_btnSaveData_Clicked);
	connect(ui.btnReceive, &QPushButton::clicked, this, &Qt_excel_test::on_btnReceive_Clicked);
	connect(ui.btnTimerOn, &QPushButton::clicked, this, &Qt_excel_test::on_btnTimerOn_Clicked);
	connect(ui.btnClearPlot1, &QPushButton::clicked, this, &Qt_excel_test::on_btnClearPlot1_Clicked);//�������1
	connect(ui.widget_my, &QCustomPlot::mouseDoubleClick, this, &Qt_excel_test::DoubleClickedPlot);//˫��������ݵ�
	connect(ui.widget_my, &QCustomPlot::customContextMenuRequested, this, &Qt_excel_test::contextMenuRequest);//�Ҽ� �˵���
	connect(ui.act_zoomIn, &QAction::toggled, this, &Qt_excel_test::on_act_zoomIn_toggled_);//������ ��ѡ�Ŵ�
	connect(ui.act_move, &QAction::toggled, this, &Qt_excel_test::on_act_move_toggled_);//������ �϶�
	connect(ui.act_tracer, &QAction::toggled, this, &Qt_excel_test::on_act_tracer_toggled_);//������ �α�
	connect(ui.widget_my, &QCustomPlot::mouseMove, this, &Qt_excel_test::myMouseMoveEvent);
	connect(ui.widget_my, &QCustomPlot::selectionChangedByUser, this, &Qt_excel_test::traceGraphSelect);//˫��ͼ��
}

 Qt_excel_test::~Qt_excel_test() 
 {
	 delete m_Timer;
 }

//��ȡ����
void Qt_excel_test::on_btnReadData_Clicked()
{
	x.clear();
	y.clear();
	x_y.clear();
	ui.labelDataStatus->setText("���ڶ�ȡ����");
	QString strFile = QFileDialog::getOpenFileName(this, "ѡ���ļ�", "./", "�ı��ļ�(*.xls;*.xlsx;)");
	if (!strFile.isEmpty())
	{
		QElapsedTimer mstimer;
		mstimer.start();//��ʱ��ʼ

		myexcel.readExcelFast(strFile,x,y);

		float time = (double)mstimer.nsecsElapsed() / (double)1000000;
		QString strTime = QString("%1").arg(time);
		ui.label_time->setText(strTime);
	}
}

//��������
void Qt_excel_test::on_btnPlot_Clicked()
{
	//m_Timer->start();
	ui.widget_my->graph(0)->setData(x, y);
	ui.widget_my->replot();
	ui.labelDataStatus->setText("��ͼ���");
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

//��������
void Qt_excel_test::on_btnSaveData_Clicked()
{
	ui.labelDataStatus->setText("���ڱ�������");
	QString strFile = QFileDialog::getSaveFileName(this, "���Ϊ", "./", "�ı��ļ�(*.xlsx;)");
	if (!strFile.isEmpty())
	{
		QElapsedTimer mstimer;
		mstimer.start();//��ʱ��ʼ

		myexcel.writeExcelFast(strFile, x_y);

		float time = (double)mstimer.nsecsElapsed() / (double)1000000;
		QString strTime = QString("%1").arg(time);
		ui.label_time->setText(strTime);//��ʾʱ��
	}
}

//����д����ź�
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

//���1������
void Qt_excel_test::on_btnClearPlot1_Clicked()
{
	x.clear();
	y.clear();
	ui.widget_my->graph(1)->setData(x, y);
	ui.widget_my->replot(QCustomPlot::rpQueuedReplot);
}

//˫��������ݵ�
void Qt_excel_test::DoubleClickedPlot(QMouseEvent *event)
{
	QPoint point = event->pos();
	//��ȡ��������
	double x_pos = point.x();
	double y_pos = point.y();
	//���������� ת��Ϊ QCustomPlot �ڲ�����ֵ ��pixelToCoord ������
	//coordToPixel ������֮�෴ �ǰ��ڲ�����ֵ ת��Ϊ�ⲿ�����
	double x_val = ui.widget_my->xAxis->pixelToCoord(x_pos);
	double y_val = ui.widget_my->yAxis->pixelToCoord(y_pos);
	addPoint(x_val, y_val);
	ui.widget_my->graph(1)->setData(x, y);
	ui.widget_my->replot(QCustomPlot::rpQueuedReplot);
}

//������ݵ�
void Qt_excel_test::addPoint(const double &x_temp, const double &y_temp)
{
	QList<QVariant> xy;
	x.append(x_temp);
	y.append(y_temp);
	
	xy.append(x_temp);
	xy.append(y_temp);
	x_y.append(xy);
}

//�Ҽ��˵� ������Χ
void Qt_excel_test::contextMenuRequest(QPoint pos)
{
	QMenu *menu = new QMenu(this);
	menu->setAttribute(Qt::WA_DeleteOnClose);
	menu->popup(ui.widget_my->mapToGlobal(pos));
	menu->addAction("������Χ", this, &Qt_excel_test::rescaleAxes);
}

//����ȫ����ʾ
void Qt_excel_test::rescaleAxes()
{
	//����һ��graph����rescaleAxes()����������graph������rescaleAxes(true)����ʵ����ʾ�������ߡ�
	/*rescaleAxes(true)ʱ���plot��X��Y�᱾���������±�graph��X��Y���ݵ㣬
	��ôplot��X��Y��Ŀ��ӷ�Χ�����������ֻ��plot���ɲ��±�graphʱ������չplot���������ʾ��Χ��*/
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
	ui.widget_my->replot(QCustomPlot::rpQueuedReplot); //ˢ��ͼ�꣬����ʡ��
}

void Qt_excel_test::setVisibleTracerTrue()
{
	tracer->setVisible(true);
	tracer1->setVisible(true);
	tracerXLabel->setVisible(true);
	tracerYLabel->setVisible(true);
	tracerY1Label->setVisible(true);
	ui.widget_my->replot(QCustomPlot::rpQueuedReplot); //ˢ��ͼ�꣬����ʡ��
}

//�����ѡ�Ŵ�
void Qt_excel_test::on_act_zoomIn_toggled_(bool arg1)
{
	if (arg1)
	{
		ui.act_move->setChecked(false);//ȡ���϶�ѡ��
		ui.widget_my->setInteraction(QCP::iRangeDrag, false);//ȡ���϶�
		ui.widget_my->setSelectionRectMode(QCP::SelectionRectMode::srmZoom);
	}
	else
	{
		ui.widget_my->setSelectionRectMode(QCP::SelectionRectMode::srmNone);
	}
}

//����϶�
void Qt_excel_test::on_act_move_toggled_(bool arg1)
{
	if (arg1)
	{
		ui.act_zoomIn->setChecked(false);//ȡ����ѡ�Ŵ�ѡ��
		ui.widget_my->setSelectionRectMode(QCP::SelectionRectMode::srmNone);
		ui.widget_my->setInteraction(QCP::iRangeDrag, true);//ʹ���϶�
	}
	else
	{
		ui.widget_my->setInteraction(QCP::iRangeDrag, false);
	}
}

//�����α�
void Qt_excel_test::on_act_tracer_toggled_(bool arg1)
{
	if (arg1)
	{
		tracerEnable = true;
		//this->setMouseTracking(true);
		tracer = new QCPItemTracer(ui.widget_my);
		tracer->setStyle(QCPItemTracer::tsCrosshair);//�α���ʽ��ʮ���ǡ�ԲȦ������
		tracer->setPen(QPen(Qt::green));//����tracer����ɫ��ɫ
		tracer->setPen(QPen(Qt::DashLine));//�����α�
		tracer->setBrush(QBrush(Qt::red));
		tracer->setSize(10);
		tracer->setInterpolating(true);//false���ò�ֵ

		tracer1 = new QCPItemTracer(ui.widget_my);
		tracer1->setStyle(QCPItemTracer::tsCrosshair);//�α���ʽ��ʮ���ǡ�ԲȦ������
		tracer1->setPen(QPen(Qt::green));//����tracer����ɫ��ɫ
		tracer1->setPen(QPen(Qt::DashLine));//�����α�
		tracer1->setBrush(QBrush(Qt::red));
		tracer1->setSize(10);
		tracer1->setInterpolating(true);//false���ò�ֵ


		tracerXLabel = new QCPItemText(ui.widget_my);
		tracerXLabel->setClipToAxisRect(false);
		tracerXLabel->setLayer("overlay");
		tracerXLabel->setPen(QPen(Qt::green));
		tracerXLabel->setFont(QFont("����", 10));
		tracerXLabel->setPadding(QMargins(5, 5, 5, 5));
		tracerXLabel->setPositionAlignment(Qt::AlignLeft | Qt::AlignTop);
		//�������������Ҫ�������α�˵��ê����tracerλ�ô���ʵ���Զ�����
		tracerXLabel->position->setType(QCPItemPosition::ptAxisRectRatio);//λ�����ͣ���ǰ�᷶Χ�ı���Ϊ��λ/ʵ������Ϊ��λ��
		tracerXLabel->position->setParentAnchorX(tracer->position);
		
		tracerYLabel = new QCPItemText(ui.widget_my);
		tracerYLabel->setClipToAxisRect(false);
		tracerYLabel->setLayer("overlay");
		tracerYLabel->setPen(QPen(Qt::red));
		tracerYLabel->setFont(QFont("����", 10));
		tracerYLabel->setPadding(QMargins(5, 5, 5, 5));
		tracerYLabel->setPositionAlignment(Qt::AlignLeft | Qt::AlignTop);
		//�������������Ҫ�������α�˵��ê����tracerλ�ô���ʵ���Զ�����
		tracerYLabel->position->setType(QCPItemPosition::ptAxisRectRatio);//λ�����ͣ���ǰ�᷶Χ�ı���Ϊ��λ/ʵ������Ϊ��λ��
		tracerYLabel->position->setParentAnchorY(tracer->position);

		tracerY1Label = new QCPItemText(ui.widget_my);
		tracerY1Label->setClipToAxisRect(false);
		tracerY1Label->setLayer("overlay");
		tracerY1Label->setPen(QPen(Qt::red));
		tracerY1Label->setFont(QFont("����", 10));
		tracerY1Label->setPadding(QMargins(5, 5, 5, 5));
		tracerY1Label->setPositionAlignment(Qt::AlignLeft | Qt::AlignTop);
		//�������������Ҫ�������α�˵��ê����tracerλ�ô���ʵ���Զ�����
		tracerY1Label->position->setType(QCPItemPosition::ptAxisRectRatio);//λ�����ͣ���ǰ�᷶Χ�ı���Ϊ��λ/ʵ������Ϊ��λ��
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
		if (item->selected() || graph->selected())//ѡ�����������߻������ߵ�ͼ��
		{
			tracerGraph = graph;
		}
		
	}
}


void Qt_excel_test::myMouseMoveEvent(QMouseEvent* event)
{
	if (tracerEnable)//�α�ʹ��
	{
		//if (tracerGraph != nullptr)
		//{
		//QCustomPlot::mouseMoveEvent(event);//�����ø�����ƶ��¼�ʧЧ

		double x = ui.widget_my->xAxis->pixelToCoord(event->pos().x());//�������������תplot����

		foundRange = true;
		QrangeX = ui.widget_my->graph(0)->getKeyRange(foundRange,QCP::sdBoth);
		QrangeX_lower = QrangeX.lower;
		QrangeX_upper = QrangeX.upper;
		if (x < QrangeX_upper && x > QrangeX_lower)
		{
			tracer->setGraph(ui.widget_my->graph(0));//�����α�������traceGraph����������
			tracer->setGraphKey(x);//���α�����꣨key�����óɸջ�õĺ���������x ��������α��涯�Ĺؼ����룩
			tracer->updatePosition(); //ʹ�ø������α�ĺ�������λ����Ч
			double traceX = tracer->position->key();
			double traceY = tracer->position->value();
			double traceY1 = tracer1->position->value();
			tracerXLabel->setText(QString::number(traceX, 'f', 3));//�α��ı���ָʾ�α��Xֵ
			tracerYLabel->setText(QString::number(traceY, 'f', 3));//�α��ı���ָʾ�α��Yֵ
			tracer1->setGraph(ui.widget_my->graph(1));//�����α�������traceGraph����������
			tracer1->setGraphKey(x);
			tracer1->updatePosition(); //ʹ�ø������α�ĺ�������λ����Ч
			tracerXLabel->setText(QString::number(traceX, 'f', 3));//�α��ı���ָʾ�α��Xֵ
			tracerYLabel->setText(QString::number(traceY, 'f', 3));//�α��ı���ָʾ�α��Yֵ
			tracerY1Label->setText(QString::number(traceY1, 'f', 3));//�α��ı���ָʾ�α��Yֵ
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
//	QCustomPlot::mouseMoveEvent(event);//�����ø�����ƶ��¼�ʧЧ
//
//	if (tracerEnable)//�α�ʹ��
//	{
//		if (tracerGraph != NULL)//���һ���α긽�ŵ�graph�Ƿ񻹴���
//		{
//			double x = ui.widget_my->xAxis->pixelToCoord(event->pos().x());//�������������תplot����
//			tracer->setGraph(tracerGraph);//�����α�������traceGraph����������
//			tracer->setGraphKey(x);//���α�����꣨key�����óɸջ�õĺ���������x ��������α��涯�Ĺؼ����룩
//			double traceX = tracer->position->key();
//			double traceY = tracer->position->value();
//
//			tracerXLabel->setText(QString::number(traceX));//�α��ı���ָʾ�α��Xֵ
//			//tracerYLabel->setText(QString::number(traceY));//�α��ı���ָʾ�α��Yֵ
//
//			//�����α�Xֵ��Ӧ���������ߵ�Yֵ
//			/*
//			for (int i = 0; i < ui.widget_my->graphCount(); i++)
//			{
//				//���������traceX�����key��Ӧ�ĵ㣬����ο�findBegin�����İ���
//				QCPDataContainer<QCPGraphData>::const_iterator coorPoint = ui.widget_my->graph(i)->data().data()->findBegin(traceX, true);//true������������
//				qDebug() << QString("graph%1��Ӧ��Yֵ��").arg(i) << coorPoint->value;
//			}
//			*/
//			ui.widget_my->replot(QCustomPlot::rpQueuedReplot); //ˢ��ͼ�꣬����ʡ��
//		}
//	}
//}

