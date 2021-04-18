#include "qt_excel_test.h"
#include <QtWidgets/QApplication>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    Qt_excel_test w;
    w.show();
    return a.exec();
}
