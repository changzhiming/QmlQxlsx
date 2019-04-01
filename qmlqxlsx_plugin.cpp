#include "qmlqxlsx_plugin.h"
#include "qmlqxlsx.h"

#include <qqml.h>

void QmlQXlsxPlugin::registerTypes(const char *uri)
{
    // @uri QmlQXlsx
    qmlRegisterType<QmlQXlsx>(uri, 1, 0, "QmlQXlsx");
}

