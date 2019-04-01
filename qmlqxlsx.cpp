#include "qmlqxlsx.h"
#include <QDebug>

QmlQXlsx::QmlQXlsx(QQuickItem *parent):
    QQuickItem(parent)
{
    // By default, QQuickItem does not draw anything. If you subclass
    // QQuickItem to create a visual item, you will need to uncomment the
    // following line and re-implement updatePaintNode()

    // setFlag(ItemHasContents, true);
    xlsx = new QXlsx::Document(this);
}

QmlQXlsx::~QmlQXlsx()
{
    qDebug()<<"delete QmlQxlsx";
}

bool QmlQXlsx::write(int row, int col, const QVariant &value)
{
    return xlsx->write(row, col, value);
}
QVariant QmlQXlsx::read(int row, int col) const
{
    return xlsx->read(row, col);
}

bool QmlQXlsx::setColumnWidth(int column, double width)
{
    return xlsx->setColumnWidth(column, width);
}
bool QmlQXlsx::setColumnHidden(int column, bool hidden)
{
    return xlsx->setColumnHidden(column, hidden);
}
bool QmlQXlsx::setColumnWidth(int colFirst, int colLast, double width)
{
    return xlsx->setColumnWidth(colFirst, colLast, width);
}
bool QmlQXlsx::setColumnHidden(int colFirst, int colLast, bool hidden)
{
    return xlsx->setColumnHidden(colFirst, colLast, hidden);
}
double QmlQXlsx::columnWidth(int column)
{
    return xlsx->columnWidth(column);
}
bool QmlQXlsx::isColumnHidden(int column)
{
    return xlsx->isColumnHidden(column);
}

bool QmlQXlsx::setRowHeight(int row, double height)
{
    return xlsx->setRowHeight(row, height);
}
bool QmlQXlsx::setRowHidden(int row, bool hidden)
{
    return xlsx->setRowHidden(row, hidden);
}
bool QmlQXlsx::setRowHeight(int rowFirst, int rowLast, double height)
{
    return xlsx->setRowHeight(rowFirst, rowLast, height);
}
bool QmlQXlsx::setRowHidden(int rowFirst, int rowLast, bool hidden)
{
    return xlsx->setRowHidden(rowFirst, rowLast, hidden);
}

double QmlQXlsx::rowHeight(int row)
{
    return xlsx->rowHeight(row);
}
bool QmlQXlsx::isRowHidden(int row)
{
    return xlsx->isRowHidden(row);
}

bool QmlQXlsx::groupRows(int rowFirst, int rowLast, bool collapsed )
{
    return xlsx->groupRows(rowFirst, rowLast, collapsed);
}
bool QmlQXlsx::groupColumns(int colFirst, int colLast, bool collapsed )
{
    return xlsx->groupColumns(colFirst, colLast, collapsed);
}


bool QmlQXlsx::defineName(const QString &name, const QString &formula, const QString &comment, const QString &scope)
{
    return xlsx->defineName(name, formula, comment, scope);
}

QString QmlQXlsx::documentProperty(const QString &name) const
{
    return xlsx->documentProperty(name);
}
void QmlQXlsx::setDocumentProperty(const QString &name, const QString &property)
{
    return xlsx->setDocumentProperty(name, property);
}
QStringList QmlQXlsx::documentPropertyNames() const
{
    return  xlsx->documentPropertyNames();
}

QStringList QmlQXlsx::sheetNames() const
{
    return xlsx->sheetNames();
}
bool QmlQXlsx::addSheet(const QString &name )
{
    return xlsx->addSheet(name);
}
bool QmlQXlsx::insertSheet(int index, const QString &name)
{
    return xlsx->insertSheet(index, name);
}
bool QmlQXlsx::selectSheet(const QString &name)
{
    return xlsx->selectSheet(name);
}
bool QmlQXlsx::renameSheet(const QString &oldName, const QString &newName)
{
    return xlsx->renameSheet(oldName, newName);
}
bool QmlQXlsx::copySheet(const QString &srcName, const QString &distName)
{
    return xlsx->copySheet(srcName, distName);
}
bool QmlQXlsx::moveSheet(const QString &srcName, int distIndex)
{
    return xlsx->moveSheet(srcName, distIndex);
}
bool QmlQXlsx::deleteSheet(const QString &name)
{
    return xlsx->deleteSheet(name);
}

bool QmlQXlsx::save() const
{
    return xlsx->save();
}
bool QmlQXlsx::saveAs(const QString &xlsXname) const
{
    return xlsx->saveAs(xlsXname);
}

bool QmlQXlsx::clear()
{
    xlsx->deleteLater();
    xlsx = nullptr;
    xlsx = new QXlsx::Document(this);
    return xlsx != nullptr ? true : false;
}








