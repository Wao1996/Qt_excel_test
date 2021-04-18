#include "qtGraph.h"

qtGraph::qtGraph(QCustomPlot * Plot) 
{
	myPlot = Plot;
	tracerEnable = false;
	/*************��ͼģ��***************/
	//�趨���Ͻ�ͼ����ע�ɼ�
	myPlot->legend->setVisible(true);
	//�趨���Ͻ�ͼ����ע������
	myPlot->legend->setFont(QFont("Helvetica", 9));
	//���ͼ��
	myPlot->addGraph();
	myPlot->addGraph();
	//����ȫ���ɼ�
	myPlot->graph(0)->rescaleAxes();
	myPlot->graph(1)->rescaleAxes(true);
	//���û���
	QPen pen1(Qt::red, 1, Qt::SolidLine, Qt::RoundCap, Qt::RoundJoin);
	myPlot->graph(0)->setPen(QPen(Qt::blue));
	myPlot->graph(1)->setPen(pen1);
	//��������
	myPlot->graph(0)->setLineStyle(QCPGraph::lsLine);
	myPlot->graph(1)->setLineStyle(QCPGraph::lsLine);
	//�������ϵ�ķ��
	myPlot->graph(0)->setScatterStyle(QCPScatterStyle::ssNone);
	myPlot->graph(1)->setScatterStyle(QCPScatterStyle::ssNone); //ssCircle
	//myPlot->graph(1)->setAdaptiveSampling(true);//����Ӧ����
	//myPlot->graph(0)->rescaleAxes(true);//��û����������ʾ��Χ ���Զ�������Χ��ͼ��ǡ��ȫ����ʾ����
	//�Ҽ��˵��Զ���
	myPlot->setContextMenuPolicy(Qt::CustomContextMenu);
	/**********������ͼ��ģ��************/
	myPlot->selectionRect()->setPen(QPen(Qt::black, 1, Qt::DashLine));//����ѡ�����ʽ������
	myPlot->selectionRect()->setBrush(QBrush(QColor(0, 0, 100, 50)));//����ѡ�����ʽ����͸��ǳ��
	myPlot->setInteraction(QCP::iRangeDrag, true); //��굥���϶� QCPAxisRect::mousePressEvent() ����϶�
	myPlot->setInteraction(QCP::iRangeZoom, true); //���ֻ�������
	myPlot->setInteraction(QCP::iSelectAxes, false); //true���߿�ѡ
	myPlot->setInteraction(QCP::iSelectLegend, true); //ͼ����ѡ
}

qtGraph::~qtGraph()
{
}

//X������
void qtGraph::setxAxisName(QString  name)
{
	myPlot->xAxis->setLabel(name);
}
//Y������
void qtGraph::setyAxisName(QString  name)
{
	myPlot->yAxis->setLabel(name);
}
//X�᷶Χ
void qtGraph::setxAxisRange(double lower, double upper)
{
	myPlot->xAxis->setRange(lower, upper);
}
//Y�᷶Χ
void qtGraph::setyAxisRange(double lower, double upper)
{
	myPlot->yAxis->setRange(lower, upper);
}
//����ͼ������
void qtGraph::setGraphName(QString name0, QString name1)
{
	//�������Ͻ�ͼ�α�ע����
	myPlot->graph(0)->setName(name0);
	myPlot->graph(1)->setName(name1);
}


void qtGraph::setVisibleTracerFalse()
{
	tracer0->setVisible(false);
	tracer1->setVisible(false);
	tracer0Label->setVisible(false);
	tracer1Label->setVisible(false);
	myPlot->replot(QCustomPlot::rpQueuedReplot); //ˢ��ͼ�꣬����ʡ��
}

void qtGraph::setVisibleTracerTrue()
{
	tracer0->setVisible(true);
	tracer1->setVisible(true);
	tracer0Label->setVisible(true);
	tracer1Label->setVisible(true);
	myPlot->replot(QCustomPlot::rpQueuedReplot); //ˢ��ͼ�꣬����ʡ��
}


/***************�ۺ���*****************/
//����/�ر��α�
void qtGraph::on_act_tracer_toggled_(bool arg1)
{
	if (arg1)
	{
		tracerEnable = true;
		tracer0 = new QCPItemTracer(myPlot);
		tracer0->setStyle(QCPItemTracer::tsCrosshair);//�α���ʽ��ʮ���ǡ�ԲȦ������
		tracer0->setPen(QPen(Qt::green));//����tracer����ɫ��ɫ
		tracer0->setPen(QPen(Qt::DashLine));//�����α�
		tracer0->setBrush(QBrush(Qt::red));
		tracer0->setSize(10);
		tracer0->setInterpolating(true);//false���ò�ֵ

		tracer1 = new QCPItemTracer(myPlot);
		tracer1->setStyle(QCPItemTracer::tsCrosshair);//�α���ʽ��ʮ���ǡ�ԲȦ������
		tracer1->setPen(QPen(Qt::green));//����tracer����ɫ��ɫ
		tracer1->setPen(QPen(Qt::DashLine));//�����α�
		tracer1->setBrush(QBrush(Qt::red));
		tracer1->setSize(10);
		tracer1->setInterpolating(true);//false���ò�ֵ


		tracer0Label = new QCPItemText(myPlot);
		tracer0Label->setClipToAxisRect(false);
		tracer0Label->setLayer("overlay");
		tracer0Label->setPen(QPen(Qt::green));
		tracer0Label->setFont(QFont("����", 10));
		tracer0Label->setPadding(QMargins(2, 2, 2, 2));
		tracer0Label->setPositionAlignment(Qt::AlignLeft | Qt::AlignTop);
		//�������������Ҫ�������α�˵��ê����tracerλ�ô���ʵ���Զ�����
		tracer0Label->position->setType(QCPItemPosition::ptAxisRectRatio);//λ�����ͣ���ǰ�᷶Χ�ı���Ϊ��λ/ʵ������Ϊ��λ��
		tracer0Label->position->setParentAnchor(tracer0->position);


		tracer1Label = new QCPItemText(myPlot);
		tracer1Label->setClipToAxisRect(false);
		tracer1Label->setLayer("overlay");
		tracer1Label->setPen(QPen(Qt::red));
		tracer1Label->setFont(QFont("����", 10));
		tracer1Label->setPadding(QMargins(2, 2, 2, 2));
		tracer1Label->setPositionAlignment(Qt::AlignLeft | Qt::AlignTop);
		//�������������Ҫ�������α�˵��ê����tracerλ�ô���ʵ���Զ�����
		tracer1Label->position->setType(QCPItemPosition::ptAxisRectRatio);//λ�����ͣ���ǰ�᷶Χ�ı���Ϊ��λ/ʵ������Ϊ��λ��
		tracer1Label->position->setParentAnchor(tracer1->position);
	}
	else
	{
		tracerEnable = false;
		setVisibleTracerFalse();

	}
}
//����ƶ��¼�
void qtGraph::myMouseMoveEvent(QMouseEvent* event)
{
	if (tracerEnable)//�α�ʹ���ж�
	{
		double x = myPlot->xAxis->pixelToCoord(event->pos().x());//�������������תplot����
		
		foundRange = true;
		QrangeX0 = myPlot->graph(0)->getKeyRange(foundRange, QCP::sdBoth);//��ȡ0������X�����귶Χ
		QrangeX0_lower = QrangeX0.lower;
		QrangeX0_upper = QrangeX0.upper;

		QrangeX1 = myPlot->graph(1)->getKeyRange(foundRange, QCP::sdBoth);//��ȡ1������X�����귶Χ
		QrangeX1_lower = QrangeX1.lower;
		QrangeX1_upper = QrangeX1.upper;
		//�������ƶ�����0������X�᷶Χ����0�����������α�
		if (x < QrangeX0_upper && x > QrangeX0_lower)
		{
			tracer0->setGraph(myPlot->graph(0));//�����α�������traceGraph����������
			tracer0->setGraphKey(x);//���α�����꣨key�����óɸջ�õĺ���������x ��������α��涯�Ĺؼ����룩
			tracer0->updatePosition(); //ʹ�ø������α�ĺ�������λ����Ч
			double traceX0 = tracer0->position->key();
			double traceY0 = tracer0->position->value();
			
			tracer0Label->setText(QString::number(traceX0, 'f', 3));//�α��ı���ָʾ�α��Xֵ
			tracer0Label->setText(QString("x = %1\ny = %2").arg(QString::number(traceX0, 'f', 3)).arg(QString::number(traceY0, 'f', 3)));
			tracer0->setVisible(true);
			tracer0Label->setVisible(true);	
		}
		else
		{
			tracer0->setVisible(false);
			tracer0Label->setVisible(false);
			
		}
		//�������ƶ�����1������X�᷶Χ����1�����������α�
		if (x < QrangeX1_upper && x > QrangeX1_lower)
		{
			double traceX1 = tracer1->position->key();
			double traceY1 = tracer1->position->value();
			tracer1->setGraph(myPlot->graph(1));//�����α�������traceGraph����������
			tracer1->setGraphKey(x);
			tracer1->updatePosition(); //ʹ�ø������α�ĺ�������λ����Ч

			tracer1Label->setText(QString("x = %1\ny = %2").arg(QString::number(traceX1, 'f', 3)).arg(QString::number(traceY1, 'f', 3)));

			tracer1->setVisible(true);
			tracer1Label->setVisible(true);
		}
		else
		{
			tracer1->setVisible(false);
			tracer1Label->setVisible(false);
		}
		myPlot->replot(QCustomPlot::rpQueuedReplot); //ˢ��ͼ�꣬����ʡ��
	}
}
////����ȫ����ʾ
void qtGraph::rescaleAxes()
{
	//����һ��graph����rescaleAxes()����������graph������rescaleAxes(true)����ʵ����ʾ�������ߡ�
	/*rescaleAxes(true)ʱ���plot��X��Y�᱾���������±�graph��X��Y���ݵ㣬
	��ôplot��X��Y��Ŀ��ӷ�Χ�����������ֻ��plot���ɲ��±�graphʱ������չplot���������ʾ��Χ��*/
	myPlot->graph(0)->rescaleAxes();
	myPlot->graph(1)->rescaleAxes(true);
	myPlot->replot();
}
