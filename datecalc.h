#ifndef DATECALC_H
#define DATECALC_H

#include <QString>
#include <QStringList>
#include <QDate>
#include "soldier.h"

class DateCalc
{
public:
    DateCalc() = delete;
    static QString minusTwoDate(const QString &start, const QString &end, const bool &isEndDate = true, const bool &dateOutput = false);
    static QString addTwoDate(const QString &start, const QString &end, const bool &dateOutput = true);
    static QString dateOutput(const QString &date);
    static QString makeStringDate(const int year, const int month, const int day);
    static QString daysToDate(const int &days);
    static int dateToDays(const QString &date);
    static QString addDaysToDate(const QString &date, const int days);
    static void nahastCalc(const QString &start, const QString &end, int &total, int &khala, int &oldNahast, int &nahast, int &eidNahast, const int &oldNahastMulti, const int &nahastMulti, const int &eidNahastMulti, const bool &isKan = false);
    static bool isMinusValid(const QString &start, const QString &end);
    static bool isCrossover(const QString &start, const QString &end, const QStringList &starts, const QStringList &ends);
    static QString dateDistanceInWords(const QString &date);
};

#endif // DATECALC_H
