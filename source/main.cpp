#include <QApplication>
#include "widget.h"

// ----------------------------------------------------------------------
int main (int argc, char** argv)
{
    QApplication app(argc, argv);

    FileFinder fileFinder;

    fileFinder.resize(400, 240);
    fileFinder.show();

    return app.exec();
}
