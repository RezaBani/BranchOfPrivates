#include "datecalc.h"

QString DateCalc::minusTwoDate(const QString &start, const QString &end, const bool &isEndDate, const bool &dateOutput)
{
    int startDay = Soldier::PerToEngNo(start).split("/").at(2).toInt();
    int startMonth = Soldier::PerToEngNo(start).split("/").at(1).toInt();
    int startYear = Soldier::PerToEngNo(start).split("/").at(0).toInt();
    int endDay = Soldier::PerToEngNo(end).split("/").at(2).toInt();
    int endMonth = Soldier::PerToEngNo(end).split("/").at(1).toInt();
    int endYear = Soldier::PerToEngNo(end).split("/").at(0).toInt();
    int days = 0;
    int months = 0;
    int years = 0;
    if (endDay < startDay){
        if ((endMonth == 1) && (isEndDate)){
            endDay += 29;
            endMonth = 12;
            endYear -= 1;
        }
        else if ((endMonth < 8) && (isEndDate)){
            endDay += 31;
            endMonth -= 1;
        }
        else {
            endDay += 30;
            endMonth -= 1;
        }
    }
    days = endDay - startDay;
    if (endMonth < startMonth){
        endYear -= 1;
        endMonth += 12;
    }
    months = endMonth - startMonth;
    years = endYear - startYear;
    QString date = DateCalc::makeStringDate(years,months,days);
    if(dateOutput){
        return DateCalc::dateOutput(date);
    } else {
        return date;
    }
}

QString DateCalc::addTwoDate(const QString &start, const QString &end, const bool &dateOutput)
{
    int startDay = Soldier::PerToEngNo(start).split("/").at(2).toInt();
    int startMonth = Soldier::PerToEngNo(start).split("/").at(1).toInt();
    int startYear = Soldier::PerToEngNo(start).split("/").at(0).toInt();
    int endDay = Soldier::PerToEngNo(end).split("/").at(2).toInt();
    int endMonth = Soldier::PerToEngNo(end).split("/").at(1).toInt();
    int endYear = Soldier::PerToEngNo(end).split("/").at(0).toInt();
    int days = 0;
    int months = 0;
    int years = 0;
    days = startDay + endDay;
    months = startMonth + endMonth;
    years = startYear + endYear;
    if(dateOutput){
        if(months > 12){
            years += (months/12);
            months %= 12;
        }
        if(months < 7){
            months += (days / 32);
            days %= 32;
        } else if (months < 12){
            months += (days / 31);
            days %= 31;
        } else if (months == 12){
            months += (days / 30);
            days %= 30;
        }
    } else {
        months += (days / 30);
        days %= 30;
    }
    years += (months/12);
    months %= 12;
    QString date = DateCalc::makeStringDate(years,months,days);
    if(dateOutput){
        return DateCalc::dateOutput(date);
    } else {
        return date;
    }
}

QString DateCalc::dateOutput(const QString &date)
{
    int days = Soldier::PerToEngNo(date).split("/").at(2).toInt();
    int months = Soldier::PerToEngNo(date).split("/").at(1).toInt();
    int years = Soldier::PerToEngNo(date).split("/").at(0).toInt();
    if (months == 0){
        years -= 1;
        months = 12;
    } if (days == 0){
        months -= 1;
        days = 30;
    } if (months == 0){
        years -= 1;
        months = 12;
    }
    return DateCalc::makeStringDate(years,months,days);
}

QString DateCalc::makeStringDate(const int year, const int month, const int day)
{
    QString date;
    if(year < 1000){
        date += QString::number(0);
    }
    if(year < 100){
        date += QString::number(0);
    }
    if(year < 10){
        date += QString::number(0);
    }
    date += QString::number(year)+ "/";
    if(month < 10){
        date += QString::number(0);
    }
    date += QString::number(month) + "/";
    if(day < 10){
        date += QString::number(0);
    }
    date += QString::number(day);
    return date;
}

QString DateCalc::daysToDate(const int &days)
{
    int day = days % 30;
    int month = days / 30;
    int year = month / 12;
    month = month % 12;
    QString date = DateCalc::makeStringDate(year,month,day);
    return date;
}

int DateCalc::dateToDays(const QString &date)
{
    int days = Soldier::PerToEngNo(date).split("/").at(2).toInt();
    int months = Soldier::PerToEngNo(date).split("/").at(1).toInt();
    int years = Soldier::PerToEngNo(date).split("/").at(0).toInt();
    return (years*360 + months*30 + days);
}

QString DateCalc::addDaysToDate(const QString &date, const int days)
{
    return QDate::fromString(Soldier::PerToEngNo(date),Soldier::displayFormat(true),Soldier::jalali()).addDays(days).toString(Soldier::displayFormat(true),Soldier::jalali());
}

void DateCalc::nahastCalc(const QString &start, const QString &end, int &total, int &khala, int &oldNahast, int &nahast, int &eidNahast, const int &oldNahastMulti, const int &nahastMulti, const int &eidNahastMulti, const bool &isKan)
{
    int sign = 1;
    if(isKan){
        sign = -1;
    }
    int diff = DateCalc::minusTwoDate(start,end).split("/").at(2).toInt();
    khala += diff*sign;
    if((start.split("/").at(1).toInt() == 1) && (start.split("/").at(2).toInt() < 14)){
        if((end.split("/").at(1).toInt() == 1) && (end.split("/").at(2).toInt() < 14)){
            eidNahast += diff*sign;
            total += eidNahastMulti*diff*sign;
        } else {
            int eidDiff = DateCalc::minusTwoDate(start,start.split("/").at(0)+"/01/14").split("/").at(2).toInt();
            eidNahast += eidDiff*sign;
            total += eidNahastMulti*eidDiff*sign;
            int normalDiff = DateCalc::minusTwoDate(start.split("/").at(0)+"/01/14",end).split("/").at(2).toInt();
            if (end.split("/").at(0).toInt() > 1395){
                nahast += normalDiff*sign;
                total += nahastMulti*normalDiff*sign;
            } else {
                oldNahast += normalDiff*sign;
                total += oldNahastMulti*normalDiff*sign;
            }
        }
    } else if ((end.split("/").at(1).toInt() == 1) && (end.split("/").at(2).toInt() < 14)){
        int eidDiff = DateCalc::minusTwoDate(end.split("/").at(0)+"01/01",end).split("/").at(2).toInt();
        eidNahast += eidDiff*sign;
        total += eidNahastMulti*eidDiff*sign;
        int normalDiff = DateCalc::minusTwoDate(start,end.split("/").at(0)+"01/01").split("/").at(2).toInt();
        if (end.split("/").at(0).toInt() > 1395){
            nahast += normalDiff*sign;
            total += nahastMulti*normalDiff*sign;
        } else {
            oldNahast += normalDiff*sign;
            total += oldNahastMulti*normalDiff*sign;
        }
    } else if((start.mid(5) == "12/29") && (end.mid(5) == "01/15")){
        int eidDiff = DateCalc::minusTwoDate(end.split("/").at(0)+"01/01",end).split("/").at(2).toInt();
        eidNahast += eidDiff*sign;
        total += eidNahastMulti*eidDiff*sign;
        int normalDiff = DateCalc::minusTwoDate(start,end.split("/").at(0)+"01/01").split("/").at(2).toInt();
        if (end.split("/").at(0).toInt() > 1395){
            nahast += normalDiff*sign;
            total += nahastMulti*normalDiff*sign;
        } else {
            oldNahast += normalDiff*sign;
            total += oldNahastMulti*normalDiff*sign;
        }
    } else if (DateCalc::isMinusValid(start,"1395/06/18")){
        oldNahast += diff*sign;
        total += oldNahastMulti*diff*sign;
    } else {
        nahast += diff*sign;
        total += nahastMulti*diff*sign;
    }
}

bool DateCalc::isMinusValid(const QString &start, const QString &end)
{
    if(QDate::fromString(end,Soldier::displayFormat(true),Soldier::jalali()) < QDate::fromString(start,Soldier::displayFormat(true),Soldier::jalali())){
        return false;
    }
    return true;
}

bool DateCalc::isCrossover(const QString &start, const QString &end, const QStringList &starts, const QStringList &ends)
{
    for (int i = 0; i < starts.size(); i++){
        if((DateCalc::isMinusValid(starts.at(i),start)) && (DateCalc::isMinusValid(start,ends.at(i))) && (start != ends.at(i))){
            return true;
        }
        if((DateCalc::isMinusValid(starts.at(i),end)) && (DateCalc::isMinusValid(end,ends.at(i))) && (end != starts.at(i))){
            return true;
        }
    }
    return false;
}

QString DateCalc::dateDistanceInWords(const QString &date)
{
    QString text;
    int days = Soldier::PerToEngNo(date).split("/").at(2).toInt();
    int months = Soldier::PerToEngNo(date).split("/").at(1).toInt();
    int years = Soldier::PerToEngNo(date).split("/").at(0).toInt();
    if(years > 0){
        text += QString::number(years) + "سال";
        if(months > 0 || days > 0){
            text += QString(" ") + "و" + QString(" ");
        }
    }
    if(months > 0){
        text += QString::number(months) + "ماه";
        if(days > 0){
            text += QString(" ") + "و" + QString(" ");
        }
    }
    if(days > 0){
        text += QString::number(days) + "روز";
    }
    if(text.isEmpty()){
        text += "0";
        text += "روز";
    }
    return text;
}
