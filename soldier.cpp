#include "soldier.h"

QStringList Soldier::titles(const bool &isCode)
{
    QStringList titles;
    titles.append("نام و نشان?t");
    titles.append("نام پدر?t");
    titles.append("کد ملی?t");
    titles.append("تاریخ اعزام?d");
    titles.append("تاریخ معرفی?d");
    titles.append("وضعیت جسمانی?c");
    titles.append("وضعیت روانی?c");
    titles.append("سنوات?n");
    titles.append("وضعیت تاهل?c");
    titles.append("تعداد عائله?n");
    titles.append("حوزه اعزام?c");
    titles.append("آموزشی?c");
    titles.append("گروه خونی?c");
    titles.append("کد پرسنلی?t");
    titles.append("شماره حساب?t");
    titles.append("شماره بیمه?t");
    titles.append("وضعیت بومی?c");
    titles.append("استان?c");
    titles.append("میزان تحصیلات?c");
    titles.append("درجه?c");
    titles.append("رسته?c");
    titles.append("قسمت?c");
    titles.append("تخصص?c");
    titles.append("دین?c");
    titles.append("مذهب?c");
    titles.append("تاریخ تولد?d");
    titles.append("محل تولد?t");
    titles.append("آدرس?t");
    titles.append("تلفن?t");
    if (!isCode){
        return Soldier::removeCodes(titles);
    }
    return Soldier::codes(titles);
}

QStringList Soldier::tabs(const bool &isCode)
{
    QStringList tabs;
    tabs.append("لیست کارکنان وظیفه?CT");
    tabs.append("فرار ۶ ماه?D");
    tabs.append("ترخیص?DC");
    tabs.append("معافیت?dt");
    tabs.append("معافیت موقت?dt");
    tabs.append("استخدام?dc");
    tabs.append("انتقالی?dc");
    tabs.append("فوتی?dt");
    tabs.append("خرید خدمت?dt");
    tabs.append("عفو رهبری?dt");
    tabs.append("مامور?CDDC");
    tabs.append("بایگانی فرار ۱۵ روز?DD");
    tabs.append("بایگانی فرار ۶ ماه?DD");
    if (!isCode){
        return Soldier::removeCodes(tabs);
    }
    return Soldier::codes(tabs);
}

QStringList Soldier::removeCodes(const QStringList &list)
{
    QStringList result;
    for(const QString &item:list){
        result.append(item.split("?").at(0));
    }
    return result;
}

QStringList Soldier::codes(const QStringList &list)
{
    QStringList codes;
    for(const QString &item:list){
        codes.append(item.split("?").at(1));
    }
    return codes;
}

QString Soldier::PerToEngNo(const QString &text)
{
    QString result = text;
    QString perNumbers = "۰۱۲۳۴۵۶۷۸۹";
    for(int i = 0; i < text.size(); i++){
        int index = perNumbers.indexOf(text.at(i));
        if( index != -1){
            result[i] = QString::number(index).at(0);
        }
    }
    return result;
}

QLocale Soldier::persian()
{
    return QLocale(QLocale::Persian,QLocale::Iran);
}

QCalendar Soldier::jalali()
{
    return QCalendar(QCalendar::System::Jalali);
}

QString Soldier::displayFormat(const bool &setDateFormat)
{
    if(setDateFormat){
        return "yyyy/MM/dd";
    }
    return "dd/MM/yyyy";
}

void Soldier::setCalenderWidgetFormat(QCalendarWidget *calenderWidget)
{
    calenderWidget->setFirstDayOfWeek(Qt::Saturday);
    QTextCharFormat format1 = calenderWidget->weekdayTextFormat(Qt::Saturday);
    format1.setForeground(QBrush(Qt::black, Qt::SolidPattern));
    calenderWidget->setWeekdayTextFormat(Qt::Saturday, format1);

    QTextCharFormat format2 = calenderWidget->weekdayTextFormat(Qt::Sunday);
    format2.setForeground(QBrush(Qt::black, Qt::SolidPattern));
    calenderWidget->setWeekdayTextFormat(Qt::Sunday, format2);

    QTextCharFormat format3 = calenderWidget->weekdayTextFormat(Qt::Friday);
    format3.setForeground(QBrush(Qt::red, Qt::SolidPattern));
    calenderWidget->setWeekdayTextFormat(Qt::Friday, format3);
}

int Soldier::maxSpinBox()
{
    return 10000;
}

bool Soldier::isDuplicate(const QString &text, const QMap<QString, QString> &codes, const int &tab, const int &row)
{
    if(!codes.contains(text)){
        return false;
    }
    if((codes.value(text).split(",").at(0).toInt() == tab) && (codes.value(text).split(",").at(1).toInt() == row)){
        return false;
    }
    return true;
}

QStringList Soldier::orders(const bool &isCode)
{
    QStringList orders;
    orders.append("مراجعت?Addn");
    orders.append("استراحت پزشکی?Addn");
    orders.append("تشویقی?Ant");
    orders.append("تنبیهی?Antc");
    orders.append("ماموریت?Addtt");
    orders.append("ترخیص?AdT");
    if (!isCode){
        return Soldier::removeCodes(orders);
    }
    return Soldier::codes(orders);
}

QStringList Soldier::letters(const bool &isCode)
{
    QStringList letters;
    letters.append("نهست?Ado");
    letters.append("مراجعت?Addo");
    letters.append("مرخصی?AddNo");
    letters.append("شروع ماموریت?Adtto");
    letters.append("پایان ماموریت?Addtto");
    letters.append("کان لم یکن?AddNo");
    letters.append("کسر و اضافه به آمار?Adcco");
    letters.append("تشویقی?Anto");
    letters.append("اضافه خدمت?Anto");
    letters.append("بازداشت با خدمت?Anto");
    letters.append("بازداشت بدون خدمت?Anto");
    letters.append("شورا پزشکی?Acco");
    letters.append("نوبت شورا?Ado");
    letters.append("فرار ۱۵ روز?Addo");
    letters.append("فرار ۶ ماه?Addo");
    letters.append("مراجعت از فرار ۱۵ روز?Addo");
    letters.append("کان لم یکن فرار ۱۵ روز?AddNo");
    letters.append("مراجعت از فرار ۶ ماه?Addo");
    letters.append("کسری?Anco");
    letters.append("کسری منطقه عملیاتی?Addco");
    letters.append("بخشش اضافه خدمت?Ano");
    letters.append("بخشش سنوات?Ano");
    letters.append("ترخیص?Ado");
    letters.append("وضعیت درخواست کارت?Aco");
    letters.append("معافیت?Adto");
    letters.append("معافیت موقت?Adto");
    letters.append("استخدام?Adco");
    letters.append("انتقالی?Adco");
    letters.append("فوتی?Adto");
    letters.append("خرید خدمت?Adto");
    letters.append("عفو رهبری?Adto");
    if (!isCode){
        return Soldier::removeCodes(letters);
    }
    return Soldier::codes(letters);
}

QStringList Soldier::addLastTitles(const QStringList &titles, const QString &tab, const QString &codes, const bool &isCode)
{
    QStringList result = titles;
    if(isCode){
        for(const QString &code:codes.toLower().split("",Qt::SkipEmptyParts)){
            result.append(code);
        }
        return result;
    }
    if(codes == "CT"){
        result.append("وضعیت حضور");
        result.append("توضیحات");
    } else if (codes == "D"){
        result.append("تاریخ نهست");
    } else if (codes == "DC") {
        result.append("تاریخ ترخیص");
        result.append("وضعیت درخواست کارت");
    } else if(codes == "DD") {
        result.append("تاریخ نهست");
        result.append("تاریخ مراجعت");
    } else if(codes == "CDDC"){
        result.append("یگان اصلی مامور");
        result.append("تاریخ ورود به یگان");
        result.append("تاریخ پایان ماموریت");
        result.append("وضعیت حضور");
    } else {
        for(const QString &code:codes.split("",Qt::SkipEmptyParts)){
            if(code == "d"){
                result.append("تاریخ " + tab);
            } else if (code == "c"){
                result.append("محل " + tab);
            } else if (code == "t"){
                result.append("علت " + tab);
            }
        }
    }
    return result;
}
