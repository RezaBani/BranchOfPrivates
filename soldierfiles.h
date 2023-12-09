#ifndef SOLDIERFILES_H
#define SOLDIERFILES_H

#include <QObject>
#include <QFile>
#include <QDir>
#include <QFileInfo>
#include <QFileInfoList>
#include <QVariant>
#include <cmath>
#include "soldier.h"
#include "letterdialog.h"
#include "datecalc.h"

#include <xlsxdocument.h>
#include <xlsxchartsheet.h>
#include <xlsxcellrange.h>
#include <xlsxchart.h>
#include <xlsxrichstring.h>
#include <xlsxworkbook.h>
#include <xlsxformat.h>
#include <xlsxformat_p.h>

class SoldierFiles : public QObject
{
    Q_OBJECT
public:
    explicit SoldierFiles(const QMap<QString,QStringList> &combosMap, QObject *parent = nullptr);
    static void createExcel(const QString &tab, const QString &code, const QFont &font);
    static QString createOrder(const QString &dirPath, const int &orderNo, const QFont &font);
    static QString createStatistics(const QString &dirPath, const int &mainYear, const int &year, const int &month, const QMap<QString, QStringList> &combosMap, const QFont &font);
    bool validateRemove(const QString &tab, const QString &code, const QString &letterType, const int &row);
    bool validateLetter(const QString &tab, const QString &code, const QString &letterType, const QStringList &values, bool summaryRequest = false);

signals:
    void changeColumnData(const int &columnIndex, const QString &columnNewData);
    void showLetterWarning(const QString &message);
    void changeRowColor(const QString &tab, const int &rowIndex, const int &colorCode);
    void moveToTable(const QString &tab, const QString &code, const QString &newTab, const QStringList &lastValues);
    void addToTable(const QString &tab, const QString &code, const QString &newTab, const QStringList &lastValues);
    void addToOrder(const QString &tab, const QString &code, const QString &sheetName, const QStringList &values);
    void changeInOrder(const QString &sheetName, const QStringList &values, const int &newValues);
    void removeOrder(const QString &sheetName, const QStringList &values);
    void summaryList(const QList<int> &list);

private:
    QMap<QString,QStringList> m_combosMap;
    QString m_dbFilename;
    QString m_ostanFilename;
    QString m_orderDir;
};

#endif // SOLDIERFILES_H
