#ifndef QMLQXLSX_H
#define QMLQXLSX_H

#include <QQuickItem>
#include "xlsxdocument.h"
class QmlQXlsx : public QQuickItem
{
    Q_OBJECT
    Q_DISABLE_COPY(QmlQXlsx)

public:
    QmlQXlsx(QQuickItem *parent = nullptr);
    ~QmlQXlsx();

public slots:
    bool write(int row, int col, const QVariant &value);
    QVariant read(int row, int col) const;

    bool setColumnWidth(int column, double width);
    bool setColumnHidden(int column, bool hidden);
    bool setColumnWidth(int colFirst, int colLast, double width);
    bool setColumnHidden(int colFirst, int colLast, bool hidden);
    double columnWidth(int column);
    bool isColumnHidden(int column);

    bool setRowHeight(int row, double height);
    bool setRowHidden(int row, bool hidden);
    bool setRowHeight(int rowFirst, int rowLast, double height);
    bool setRowHidden(int rowFirst, int rowLast, bool hidden);

    double rowHeight(int row);
    bool isRowHidden(int row);

    bool groupRows(int rowFirst, int rowLast, bool collapsed = true);
    bool groupColumns(int colFirst, int colLast, bool collapsed = true);


    bool defineName(const QString &name, const QString &formula, const QString &comment=QString(), const QString &scope=QString());

    QString documentProperty(const QString &name) const;
    void setDocumentProperty(const QString &name, const QString &property);
    QStringList documentPropertyNames() const;

    QStringList sheetNames() const;
    bool addSheet(const QString &name = QString());
    bool insertSheet(int index, const QString &name = QString());
    bool selectSheet(const QString &name);
    bool renameSheet(const QString &oldName, const QString &newName);
    bool copySheet(const QString &srcName, const QString &distName = QString());
    bool moveSheet(const QString &srcName, int distIndex);
    bool deleteSheet(const QString &name);

    bool save() const;
    bool saveAs(const QString &xlsXname) const;

    bool clear();
private:
    QXlsx::Document  *xlsx = nullptr;
};

#endif // QMLQXLSX_H
