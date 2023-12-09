#ifndef SOLDIER_H
#define SOLDIER_H

#include <QStringList>
#include <QChar>
#include <QString>
#include <QDebug>
#include <QLocale>
#include <QCalendar>
#include <QCalendarWidget>
#include <QTextCharFormat>
#include <QMap>

class Soldier
{
public:
    Soldier() = delete;
    static QStringList titles(const bool &isCode = false);
    static QStringList tabs(const bool &isCode = false);
    static QStringList letters(const bool &isCode = false);
    static QStringList orders(const bool &isCode = false);
    static QStringList addLastTitles(const QStringList &titles, const QString &tab, const QString &codes, const bool &isCode = false);
    static QStringList removeCodes(const QStringList &list);
    static QStringList codes(const QStringList &list);
    static QString PerToEngNo(const QString &text);
    static QLocale persian();
    static QCalendar jalali();
    static QString displayFormat(const bool &setDateFormat = false);
    static void setCalenderWidgetFormat(QCalendarWidget *calenderWidget);
    static int maxSpinBox();
    static bool isDuplicate(const QString &text, const QMap<QString,QString> &codes, const int &tab, const int &row);
};

#endif // SOLDIER_H
