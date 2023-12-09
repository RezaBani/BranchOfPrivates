#include "soldierfiles.h"

SoldierFiles::SoldierFiles(const QMap<QString, QStringList> &combosMap, QObject *parent)
    : QObject{parent}
{
    m_combosMap = combosMap;
    m_dbFilename = m_combosMap.value("Files").at(1);
    m_ostanFilename = m_combosMap.value("Files").at(2);
    m_orderDir = m_combosMap.value("Files").at(3);
}

void SoldierFiles::createExcel(const QString &tab, const QString &code, const QFont &font)
{
    QStringList letters = Soldier::letters();
    QStringList lettersCodes = Soldier::letters(true);
    QDir dir(QDir::currentPath()+QDir::separator()+tab);
    if (!dir.exists()){
        dir.cdUp();
        dir.mkdir(tab);
    }
    QFileInfoList infoList = dir.entryInfoList(QDir::Filter::Files | QDir::Filter::NoDotAndDotDot);
    for(const QFileInfo &fi:infoList){
        if(fi.completeBaseName()+".xlsx" == code+".xlsx"){
            return;
        }
    }
    QXlsx::Document doc(QDir::currentPath()+QDir::separator()+tab+QDir::separator()+code+".xlsx");
    for(int i = 0; i < letters.size(); i++){
        doc.addSheet(letters.at(i));
        int col = 1;
        QString letterCodes = lettersCodes.at(i);
        bool cc_check = false;
        bool dd_check = false;
        bool tt_check = false;
        for(const QString &letterCode:letterCodes){
            if(letterCode == "A"){
                doc.write(1,1,QVariant(QString("صاحب نامه")));
                doc.write(1,2,QVariant(QString("شماره نامه")));
                doc.write(1,3,QVariant(QString("تاریخ نامه")));
                col += 3;
            } else if(letterCode == "N"){
                doc.write(1,col,QVariant(QString("سالیانه")));
                doc.write(1,col+1,QVariant(QString("تشویقی")));
                doc.write(1,col+2,QVariant(QString("توراهی")));
                doc.write(1,col+3,QVariant(QString("تعطیلی")));
                doc.write(1,col+4,QVariant(QString("استراحت پزشکی")));
                col += 5;
            } else if(letterCode == "n"){
                doc.write(1,col,QVariant(QString("تعداد روز")));
                col++;
            } else if(letterCode == "c" && !cc_check){
                if(letterCodes.contains("cc")){
                    if(letters.at(i).contains("کسر و اضافه به آمار")){
                        doc.write(1,col,QVariant(QString("قسمت کسر کننده")));
                        doc.write(1,col+1,QVariant(QString("قسمت اضافه شونده")));
                    } else if(letters.at(i).contains("شورا پزشکی")){
                        doc.write(1,col,QVariant(QString("وضعیت جسمانی")));
                        doc.write(1,col+1,QVariant(QString("وضعیت روانی")));
                    }
                    cc_check = true;
                    col += 2;
                } else {
                    doc.write(1,col,QVariant(QString(letters.at(i))));
                    col++;
                }
            } else if(letterCode == "d" && !dd_check){
                if(letterCodes.contains("dd")){
                    doc.write(1,col,QVariant(QString("تاریخ شروع")));
                    doc.write(1,col+1,QVariant(QString("تاریخ پایان")));
                    dd_check = true;
                    col += 2;
                } else {
                    doc.write(1,col,QVariant(QString("تاریخ "+letters.at(i))));
                    col++;
                }
            } else if(letterCode == "t" && !tt_check){
                if(letterCodes.contains("tt")){
                    doc.write(1,col,QVariant(QString("محل")));
                    doc.write(1,col+1,QVariant(QString("شماره امریه")));
                    tt_check = true;
                    col += 2;
                } else {
                    doc.write(1,col,QVariant(QString("علت")));
                    col++;
                }
            } else if(letterCode == "o"){
                doc.write(1,col,QVariant(QString("شماره دستور")));
                col++;
            }
        }
        QXlsx::Format f = doc.currentWorksheet()->rowFormat(1);
        f.setPatternBackgroundColor(QColor(Qt::gray));
        f.setFont(font);
        f.setHorizontalAlignment(QXlsx::Format::HorizontalAlignment::AlignHCenter);
        f.setVerticalAlignment(QXlsx::Format::VerticalAlignment::AlignVCenter);
        doc.currentWorksheet()->setRowFormat(1,1,f);
    }
    doc.save();
    for(QString &sheetName: doc.sheetNames()){
        doc.selectSheet(sheetName);
        doc.autosizeColumnWidth();
    }
    doc.save();
}

QString SoldierFiles::createOrder(const QString &dirPath, const int &orderNo, const QFont &font)
{
    QStringList orders = Soldier::orders();
    QStringList ordersCodes = Soldier::orders(true);
    QDir dir(QDir::currentPath()+QDir::separator()+dirPath);
    if (!dir.exists()){
        dir.cdUp();
        dir.mkdir(dirPath);
    }
    QString path = QDir::currentPath()+QDir::separator()+dirPath+QDir::separator()+QString::number(orderNo)+".xlsx";
    QFileInfoList infoList = dir.entryInfoList(QDir::Filter::Files | QDir::Filter::NoDotAndDotDot);
    for(const QFileInfo &fi:infoList){
        if(fi.completeBaseName()+".xlsx" == QString::number(orderNo)+".xlsx"){
            return path;
        }
    }
    QXlsx::Document doc(path);
    for(int i = 0; i < orders.size(); i++){
        doc.addSheet(orders.at(i));
        int col = 1;
        QString orderCodes = ordersCodes.at(i);
        bool dd_check = false;
        bool tt_check = false;
        for(const QString &orderCode:orderCodes){
            if(orderCode == "A"){
                doc.write(1,1,QVariant(QString("نام و نشان")));
                doc.write(1,2,QVariant(QString("نام پدر")));
                doc.write(1,3,QVariant(QString("تاریخ اعزام")));
                doc.write(1,4,QVariant(QString("کد ملی")));
                col += 4;
            } else if(orderCode == "n"){
                doc.write(1,col,QVariant(QString("تعداد روز")));
                col++;
            } else if(orderCode == "c"){
                doc.write(1,col,QVariant(QString("نوع")));
                col++;
            } else if(orderCode == "d" && !dd_check){
                if(orderCodes.contains("dd")){
                    doc.write(1,col,QVariant(QString("تاریخ شروع")));
                    doc.write(1,col+1,QVariant(QString("تاریخ پایان")));
                    dd_check = true;
                    col += 2;
                } else {
                    doc.write(1,col,QVariant(QString("تاریخ "+orders.at(i))));
                    col++;
                }
            } else if(orderCode == "t" && !tt_check){
                if(orderCodes.contains("tt")){
                    doc.write(1,col,QVariant(QString("محل")));
                    doc.write(1,col+1,QVariant(QString("شماره امریه")));
                    tt_check = true;
                    col += 2;
                } else {
                    doc.write(1,col,QVariant(QString("علت")));
                    col++;
                }
            } else if(orderCode == "T"){
                doc.write(1,col,QVariant(QString("کسری")));
                doc.write(1,col+1,QVariant(QString("کد پرسنلی")));
                doc.write(1,col+2,QVariant(QString("محل تولد")));
                doc.write(1,col+3,QVariant(QString("تاریخ تولد")));
                col += 4;
            }
        }
        QXlsx::Format f = doc.currentWorksheet()->rowFormat(1);
        f.setPatternBackgroundColor(QColor(Qt::gray));
        f.setFont(font);
        f.setHorizontalAlignment(QXlsx::Format::HorizontalAlignment::AlignHCenter);
        f.setVerticalAlignment(QXlsx::Format::VerticalAlignment::AlignVCenter);
        doc.currentWorksheet()->setRowFormat(1,1,f);
    }
    doc.save();
    for(QString &sheetName: doc.sheetNames()){
        doc.selectSheet(sheetName);
        doc.autosizeColumnWidth();
    }
    doc.save();
    return path;
}

QString SoldierFiles::createStatistics(const QString &dirPath, const int &mainYear, const int &year, const int &month, const QMap<QString, QStringList> &combosMap, const QFont &font)
{
    QDir dir(QDir::currentPath()+QDir::separator()+dirPath);
    if (!dir.exists()){
        dir.cdUp();
        dir.mkdir(dirPath);
    }
    QFileInfoList infoList = dir.entryInfoList(QDir::Filter::Files | QDir::Filter::NoDotAndDotDot);
    for(const QFileInfo &fi:infoList){
        if(fi.completeBaseName()+".xlsx" == QString::number(year)+"_"+QString::number(month)+".xlsx"){
            QFile::remove(fi.absoluteFilePath());
        }
    }
    QString path = QDir::currentPath()+QDir::separator()+dirPath+QDir::separator()+QString::number(year)+"_"+QString::number(month)+".xlsx";
    QXlsx::Document doc(path);
    QMap<QString, QStringList> combos = combosMap;
    QStringList titles = Soldier::addLastTitles(Soldier::titles(),Soldier::tabs().at(0),Soldier::tabs(true).at(0));
    for(auto it = combosMap.begin(); it != combosMap.end(); it++){
        if(!titles.contains(it.key())){
            combos.remove(it.key());
        } else {
            if(it.value().isEmpty()){
                combos.insert(it.key(),QStringList{""});
            }
        }
    }
    for(auto it = combos.begin(); it != combos.end(); it++){
        doc.addSheet(it.key());
        doc.selectSheet(it.key());
        for(int i = 0; i < it.value().size(); i++){
            doc.write(i+3,1,QVariant(it.value().at(i)));
            if(i == it.value().size()-1){
                doc.write(i+4,1,QVariant(QString("جمع")));
            }
        }
        for(int i = 0; i < 3; i++){
            QString title = "پذیرش در سال";
            title += " ";
            title += QString::number(mainYear-1+i);
            if(i == 0){
                title += " ";
                title += "و ماقبل";
            }
            doc.write(1,i*12+2,QVariant(title));
            for(int j = 0; j < 12; j++){
                doc.write(2,i*12+j+2,QVariant(Soldier::jalali().monthName(Soldier::persian(),j+1)));
            }
        }
        for(int r = 3; r < it.value().size()+4; r++){
            for(int c = 2; c < 39; c++){
                doc.write(r,c,QVariant(0));
            }
        }
        doc.write(1,38,QVariant(QString("جمع")));
        doc.currentWorksheet()->mergeCells("B1:M1");
        doc.currentWorksheet()->mergeCells("N1:Y1");
        doc.currentWorksheet()->mergeCells("Z1:AK1");
        doc.currentWorksheet()->mergeCells("AL1:AL2");
        for(int r = 1; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            QXlsx::Format f = doc.currentWorksheet()->rowFormat(r);
            if(r == 2){
                f.setRotation(90);
            }
            f.setFont(font);
            f.setHorizontalAlignment(QXlsx::Format::HorizontalAlignment::AlignHCenter);
            f.setVerticalAlignment(QXlsx::Format::VerticalAlignment::AlignVCenter);
            doc.currentWorksheet()->setRowFormat(r,r,f);
        }
    }
    doc.save();
    for(QString &sheetName: doc.sheetNames()){
        doc.selectSheet(sheetName);
        doc.autosizeColumnWidth();
    }
    doc.save();
    return path;
}

bool SoldierFiles::validateRemove(const QString &tab, const QString &code, const QString &letterType, const int &row)
{
    QDir dir(QDir::currentPath()+QDir::separator()+tab);
    if (!dir.exists()){
        emit showLetterWarning("پوشه پرونده سربازان یافت نشد");
        return false;
    }
    bool fileExist = false;
    QFileInfoList infoList = dir.entryInfoList(QDir::Filter::Files | QDir::Filter::NoDotAndDotDot);
    for(const QFileInfo &fi:infoList){
        if(fi.completeBaseName()+".xlsx" == code+".xlsx"){
            fileExist = true;
            break;
        }
    }
    if(!fileExist){
        emit showLetterWarning("فایل پرونده سرباز یافت نشد");
        return false;
    }
    QXlsx::Document doc(QDir::currentPath()+QDir::separator()+tab+QDir::separator()+code+".xlsx");
    QXlsx::Document db(m_dbFilename);
    doc.selectSheet(letterType);
    if(doc.currentWorksheet()->rowFormat(row).patternBackgroundColor().isValid()){
        emit showLetterWarning("این نامه به دلیل مرتبط بودن با نامه‌ای دیگر قابل حذف شدن نیست");
        return false;
    }
    if(doc.read(row,doc.currentWorksheet()->dimension().columnCount()).toString() == m_combosMap.value("Order").at(0)){
        QStringList ordersSheets;
        ordersSheets.append("مراجعت");
        ordersSheets.append("پایان ماموریت");
        ordersSheets.append("کان لم یکن");
        ordersSheets.append("مرخصی");
        ordersSheets.append("اضافه خدمت");
        ordersSheets.append("بازداشت با خدمت");
        ordersSheets.append("بازداشت بدون خدمت");
        ordersSheets.append("تشویقی");
        ordersSheets.append("ترخیص");
        if(ordersSheets.contains(letterType)){
            QStringList orderValues;
            QString sheet;
            if(ordersSheets.indexOf(letterType) < 8){
                orderValues.append(doc.read(row,4).toString());
                orderValues.append(doc.read(row,5).toString());
            }
            if(ordersSheets.indexOf(letterType) == 8){
                orderValues.append(doc.read(row,4).toString());
            }
            if((ordersSheets.indexOf(letterType) > 3) && (ordersSheets.indexOf(letterType) < 7)){
                orderValues.append(letterType);
                sheet = "تنبیهی";
            }
            for(QString &order:Soldier::orders()){
                if(letterType.contains(order)){
                    sheet = order;
                    break;
                }
            }
            if(sheet.isEmpty()){
                sheet = "استراحت پزشکی";
            }
            emit removeOrder(sheet,orderValues);
        }
    }
    QString linkerValue;
    QStringList outerLetter;
    outerLetter.append("مراجعت از فرار ۶ ماه");
    outerLetter.append("فرار ۶ ماه");
    outerLetter.append("مراجعت از فرار ۱۵ روز");
    outerLetter.append("فرار ۱۵ روز");
    outerLetter.append("مراجعت");
    outerLetter.append("پایان ماموریت");
    QStringList innerLetter;
    innerLetter.append("فرار ۶ ماه");
    innerLetter.append("فرار ۱۵ روز");
    innerLetter.append("فرار ۱۵ روز");
    innerLetter.append("نهست");
    innerLetter.append("نهست");
    innerLetter.append("شروع ماموریت");
    if(outerLetter.contains(letterType)){
        linkerValue = Soldier::PerToEngNo(doc.read(row,4).toString());
        doc.selectSheet(innerLetter.at(outerLetter.indexOf(letterType)));
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            if(Soldier::PerToEngNo(doc.read(r,4).toString()) == linkerValue){
                emit changeRowColor(doc.currentWorksheet()->sheetName(),r,-1);
                if(letterType == "مراجعت از فرار ۶ ماه"){
                    QStringList lastValues = {doc.currentWorksheet()->read(r,4).toString()};
                    emit moveToTable(tab,code,"فرار ۶ ماه",lastValues);
                } else if(letterType == "فرار ۶ ماه"){
                    QStringList lastValues = {"فرار", ""};
                    emit moveToTable(tab,code,"لیست کارکنان وظیفه",lastValues);
                } else if(letterType == "مراجعت از فرار ۱۵ روز"){
                    db.selectSheet(tab);
                    for (int r = 2; r <= db.currentWorksheet()->dimension().rowCount(); r++){
                        if(code == db.read(r,3).toString()){
                            for(int c = 1; c <= db.currentWorksheet()->dimension().columnCount(); c++){
                                if(db.read(1,c).toString() == "وضعیت حضور"){
                                    emit changeColumnData(c-1,"فرار");
                                    break;
                                }
                            }
                            break;
                        }
                    }
                } else if(letterType == "پایان ماموریت"){
                    db.selectSheet(tab);
                    for (int r = 2; r <= db.currentWorksheet()->dimension().rowCount(); r++){
                        if(code == db.read(r,3).toString()){
                            for(int c = 1; c <= db.currentWorksheet()->dimension().columnCount(); c++){
                                if(db.read(1,c).toString() == "وضعیت حضور"){
                                    emit changeColumnData(c-1,"ماموریت");
                                    break;
                                }
                            }
                            break;
                        }
                    }
                } else {
                    db.selectSheet(tab);
                    for (int r = 2; r <= db.currentWorksheet()->dimension().rowCount(); r++){
                        if(code == db.read(r,3).toString()){
                            for(int c = 1; c <= db.currentWorksheet()->dimension().columnCount(); c++){
                                if(db.read(1,c).toString() == "وضعیت حضور"){
                                    emit changeColumnData(c-1,"نهست");
                                    break;
                                }
                            }
                            break;
                        }
                    }
                }
                return true;
            }
        }
    }
    bool moved = false;
    outerLetter.clear();
    outerLetter.append("ترخیص");
    outerLetter.append("معافیت");
    outerLetter.append("معافیت موقت");
    outerLetter.append("استخدام");
    outerLetter.append("انتقالی");
    outerLetter.append("فوتی");
    outerLetter.append("خرید خدمت");
    outerLetter.append("عفو رهبری");
    innerLetter.removeDuplicates();
    if(outerLetter.contains(letterType)){
        for(QString &sheet: innerLetter){
            doc.selectSheet(sheet);
            for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
                if(doc.currentWorksheet()->rowFormat(r).patternBackgroundColor().name() == QColor(Qt::red).name()){
                    emit changeRowColor(doc.currentWorksheet()->sheetName(),r,-1);
                    if(sheet == "فرار ۶ ماه"){
                        QStringList lastValues = {doc.currentWorksheet()->read(r,4).toString()};
                        emit moveToTable(tab,code,"فرار ۶ ماه",lastValues);
                        moved = true;
                    } else {
                        QStringList lastValues = {"حضور", ""};
                        emit moveToTable(tab,code,"لیست کارکنان وظیفه",lastValues);
                    }
                }
            }
        }
        if (!moved){
            QStringList lastValues = {"حضور", ""};
            emit moveToTable(tab,code,"لیست کارکنان وظیفه",lastValues);
        }
    }
    outerLetter.clear();
    innerLetter.clear();
    outerLetter.append("کان لم یکن");
    outerLetter.append("کان لم یکن فرار ۱۵ روز");
    innerLetter.append("مراجعت");
    innerLetter.append("مراجعت از فرار ۱۵ روز");
    if(outerLetter.contains(letterType)){
        doc.selectSheet(letterType);
        QString start = Soldier::PerToEngNo(doc.read(row,4).toString());
        QString end = Soldier::PerToEngNo(doc.read(row,5).toString());
        doc.selectSheet(innerLetter.at(outerLetter.indexOf(letterType)));
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            QString nahast = Soldier::PerToEngNo(doc.read(r,4).toString());
            QString morajeat = Soldier::PerToEngNo(doc.read(r,5).toString());
            if((DateCalc::isMinusValid(nahast,start)) && (DateCalc::isMinusValid(end,morajeat))){
                if(letterType == "کان لم یکن"){
                    QStringList changeValues;
                    changeValues.append(doc.read(r,4).toString());
                    changeValues.append(doc.read(r,5).toString());
                    QXlsx::Document orderDoc(QDir::currentPath()+QDir::separator()+m_orderDir+QDir::separator()+ m_combosMap.value("Order").at(0)+".xlsx");
                    orderDoc.selectSheet(innerLetter.at(outerLetter.indexOf(letterType)));
                    bool match = false;
                    int row = 0;
                    for(int r = 2; r <= orderDoc.currentWorksheet()->dimension().rowCount(); r++){
                        if(match){
                            break;
                        }
                        if(orderDoc.read(r,4).toString() != code){
                            continue;
                        }
                        for(int c = 0; c < changeValues.size(); c++){
                            if(orderDoc.read(r,c+5).toString() != changeValues.at(c)){
                                break;
                            }
                            if(c == changeValues.size() -1){
                                match = true;
                                row = r;
                            }
                        }
                    }
                    if(row > 1){
                        int newValue = (Soldier::PerToEngNo(orderDoc.read(row,7).toString()).toInt() + DateCalc::dateToDays(DateCalc::minusTwoDate(start,end)));
                        emit changeInOrder(doc.currentWorksheet()->sheetName(),changeValues,newValue);
                    }
                }
                bool changeColor = true;
                doc.selectSheet(letterType);
                for(int ri = 2; ri <= doc.currentWorksheet()->dimension().rowCount(); ri++){
                    if(ri == row){
                        continue;
                    }
                    start = Soldier::PerToEngNo(doc.read(r,4).toString());
                    end = Soldier::PerToEngNo(doc.read(r,5).toString());
                    if((DateCalc::isMinusValid(nahast,start)) && (DateCalc::isMinusValid(end,morajeat))){
                        changeColor = false;
                        break;
                    }
                }
                if(changeColor){
                    doc.selectSheet(innerLetter.at(outerLetter.indexOf(letterType)));
                    emit changeRowColor(doc.currentWorksheet()->sheetName(),r,-1);
                    return true;
                }
            }
        }
    }
    return true;
}

bool SoldierFiles::validateLetter(const QString &tab, const QString &code, const QString &letterType, const QStringList &values, bool summaryRequest)
{
    QDir dir(QDir::currentPath()+QDir::separator()+tab);
    if (!dir.exists()){
        emit showLetterWarning("پوشه پرونده سربازان یافت نشد");
        return false;
    }
    bool fileExist = false;
    QFileInfoList infoList = dir.entryInfoList(QDir::Filter::Files | QDir::Filter::NoDotAndDotDot);
    for(const QFileInfo &fi:infoList){
        if(fi.completeBaseName()+".xlsx" == code+".xlsx"){
            fileExist = true;
            break;
        }
    }
    if(!fileExist){
        emit showLetterWarning("فایل پرونده سرباز یافت نشد");
        return false;
    }
    QXlsx::Document doc(QDir::currentPath()+QDir::separator()+tab+QDir::separator()+code+".xlsx");
    QXlsx::Document db(m_dbFilename);
    doc.selectSheet(letterType);
    bool repetitive = false;
    for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
        repetitive = true;
        for(int c = 1; c <= doc.currentWorksheet()->dimension().columnCount(); c++){
            if(doc.read(r,c).toString() != values.at(c-1)){
                repetitive = false;
                break;
            }
        }
    }
    if((repetitive) && (!summaryRequest)){
        emit showLetterWarning("نامه تکراری است");
        return false;
    }
    QStringList letters = Soldier::letters();
    QStringList lettersCodes = Soldier::letters(true);
    for(int i = 0; i < letters.size(); i++){
        if((letters.at(i) == letterType) && (!summaryRequest)){
            if(lettersCodes.at(i).contains("dd")){
                QString startDate = Soldier::PerToEngNo(values.at(3));
                QString endDate = Soldier::PerToEngNo(values.at(4));
                if((!DateCalc::isMinusValid(startDate,endDate)) || (startDate == endDate)){
                    emit showLetterWarning("تاریخ پایان باید پس از تاریخ شروع باشد");
                    return false;
                }
                if(lettersCodes.at(i).contains("N")){
                    int sum = 0;
                    for (int j = 0; j < 5; j++){
                        sum += Soldier::PerToEngNo(values.at(values.size()-6+j)).toInt();
                    }
                    if(DateCalc::addDaysToDate(startDate,sum) != endDate){
                        emit showLetterWarning("جمع مقادیر سالیانه، تشویقی و... با فاصله بازه برابر نیست");
                        return false;
                    }
                }
            }
        }
    }
    if(letterType == "نهست"){
        doc.selectSheet(letterType);
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            if(!doc.currentWorksheet()->rowFormat(r).patternBackgroundColor().isValid() && (values.at(values.size()-1).isEmpty())){
                emit showLetterWarning("نهست باز دارد");
                return false;
            }
        }
        doc.selectSheet("فرار ۱۵ روز");
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            if(!doc.currentWorksheet()->rowFormat(r).patternBackgroundColor().isValid() && (values.at(values.size()-1).isEmpty())){
                emit showLetterWarning("فرار باز دارد");
                return false;
            }
        }
        doc.selectSheet("فرار ۶ ماه");
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            if(!doc.currentWorksheet()->rowFormat(r).patternBackgroundColor().isValid() && (values.at(values.size()-1).isEmpty())){
                emit showLetterWarning("فرار باز دارد");
                return false;
            }
        }
        doc.selectSheet(letterType);
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            if(doc.read(r,4).toString() == values.at(3)){
                emit showLetterWarning("نهست تکراری است");
                return false;
            }
        }
        QStringList starts;
        QStringList ends;
        QStringList sheets;
        sheets.append("مراجعت");
        sheets.append("مرخصی");
        sheets.append("مراجعت از فرار ۱۵ روز");
        sheets.append("مراجعت از فرار ۶ ماه");
        for(QString &sheet: sheets){
            doc.selectSheet(sheet);
            for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
                starts.append(Soldier::PerToEngNo(doc.read(r,4).toString()));
                ends.append(Soldier::PerToEngNo(doc.read(r,5).toString()));
            }
        }
        QString tommorow = DateCalc::addDaysToDate(Soldier::PerToEngNo(values.at(3)),1);
        if(DateCalc::isCrossover(Soldier::PerToEngNo(values.at(3)),tommorow,starts,ends)){
            emit showLetterWarning(letterType + " به دلیل تداخل با بازه‌های ثبت شده، قابل ثبت نیست ");
            return false;
        }
        if(!values.at(values.size()-1).isEmpty()){
            return true;
        }
        db.selectSheet(tab);
        for (int r = 2; r <= db.currentWorksheet()->dimension().rowCount(); r++){
            if(code == db.read(r,3).toString()){
                for(int c = 1; c <= db.currentWorksheet()->dimension().columnCount(); c++){
                    if(db.read(1,c).toString() == "وضعیت حضور"){
                        emit changeColumnData(c-1,letterType);
                        break;
                    }
                }
                break;
            }
        }
    } else if(letterType == "مراجعت"){
        doc.selectSheet("نهست");
        int row = 0;
        bool isOpenNahast = false;
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            if(!doc.currentWorksheet()->rowFormat(r).patternBackgroundColor().isValid()){
                isOpenNahast = true;
                if(doc.currentWorksheet()->read(r,4).toString() == values.at(3)){
                    row = r;
                    break;
                }
            }
        }
        if(!isOpenNahast){
            emit showLetterWarning("نهست این بازه قابل دسترسی نیست یا ثبت نشده است");
            return false;
        } else if(row < 1){
            emit showLetterWarning("عدم تطابق با تاریخ نهست");
            return false;
        }
        QStringList starts;
        QStringList ends;
        QStringList sheets;
        sheets.append("مراجعت");
        sheets.append("مرخصی");
        sheets.append("مراجعت از فرار ۱۵ روز");
        sheets.append("مراجعت از فرار ۶ ماه");
        for(QString &sheet: sheets){
            doc.selectSheet(sheet);
            for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
                starts.append(Soldier::PerToEngNo(doc.read(r,4).toString()));
                ends.append(Soldier::PerToEngNo(doc.read(r,5).toString()));
            }
        }
        if(DateCalc::isCrossover(Soldier::PerToEngNo(values.at(3)),Soldier::PerToEngNo(values.at(4)),starts,ends)){
            emit showLetterWarning(letterType + " به دلیل تداخل با بازه‌های ثبت شده، قابل ثبت نیست ");
            return false;
        }
        if(DateCalc::dateToDays(DateCalc::minusTwoDate(Soldier::PerToEngNo(values.at(3)),Soldier::PerToEngNo(values.at(4)))) > 15){
            emit showLetterWarning("بازه نهست و مراجعت نباید بیشتر از ۱۵ روز باشد");
            return false;
        }
        doc.selectSheet("نهست");
        emit changeRowColor(doc.currentWorksheet()->sheetName(),row,Qt::gray);
        if(!values.at(values.size()-1).isEmpty()){
            return true;
        }
        QStringList orderValues;
        orderValues.append(values.at(3));
        orderValues.append(values.at(4));
        orderValues.append(QString::number(DateCalc::dateToDays(DateCalc::minusTwoDate(Soldier::PerToEngNo(values.at(3)),Soldier::PerToEngNo(values.at(4))))));
        emit addToOrder(tab,code,letterType,orderValues);
        db.selectSheet(tab);
        for (int r = 2; r <= db.currentWorksheet()->dimension().rowCount(); r++){
            if(code == db.read(r,3).toString()){
                for(int c = 1; c <= db.currentWorksheet()->dimension().columnCount(); c++){
                    if(db.read(1,c).toString() == "وضعیت حضور"){
                        emit changeColumnData(c-1,"حضور");
                    }
                    break;
                }
                break;
            }
        }
    } else if(letterType == "مرخصی"){
        QStringList starts;
        QStringList ends;
        QStringList sheets;
        sheets.append("مراجعت");
        sheets.append("مرخصی");
        sheets.append("مراجعت از فرار ۱۵ روز");
        sheets.append("مراجعت از فرار ۶ ماه");
        for(QString &sheet: sheets){
            doc.selectSheet(sheet);
            for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
                starts.append(Soldier::PerToEngNo(doc.read(r,4).toString()));
                ends.append(Soldier::PerToEngNo(doc.read(r,5).toString()));
            }
        }
        if((DateCalc::isCrossover(Soldier::PerToEngNo(values.at(3)),Soldier::PerToEngNo(values.at(4)),starts,ends)) && (!summaryRequest)){
            emit showLetterWarning(letterType + " به دلیل تداخل با بازه‌های ثبت شده، قابل ثبت نیست ");
            return false;
        }
        doc.selectSheet("نهست");
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            if((!doc.currentWorksheet()->rowFormat(r).patternBackgroundColor().isValid()) && (!summaryRequest) && (values.at(values.size()-1).isEmpty())){
                emit showLetterWarning("نهست باز دارد");
                return false;
            }
        }
        doc.selectSheet("فرار ۱۵ روز");
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            if((!doc.currentWorksheet()->rowFormat(r).patternBackgroundColor().isValid()) && (!summaryRequest) && (values.at(values.size()-1).isEmpty())){
                emit showLetterWarning("فرار باز دارد");
                return false;
            }
        }
        doc.selectSheet("فرار ۶ ماه");
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            if((!doc.currentWorksheet()->rowFormat(r).patternBackgroundColor().isValid()) && (!summaryRequest) && (values.at(values.size()-1).isEmpty())){
                emit showLetterWarning("فرار باز دارد");
                return false;
            }
        }
        int usedSaliane = 0;
        int usedTashvighi = 0;
        int usedTorahi = 0;
        int usedEstelaji = 0;
        sheets.clear();
        sheets.append("مرخصی");
        sheets.append("کان لم یکن");
        sheets.append("کان لم یکن فرار ۱۵ روز");
        for(QString &sheet: sheets){
            doc.selectSheet(sheet);
            for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
                usedSaliane += doc.read(r,6).toInt();
                usedTashvighi += doc.read(r,7).toInt();
                usedTorahi += doc.read(r,8).toInt();
                usedEstelaji += doc.read(r,10).toInt();
            }
        }
        int availTashvighi = 0;
        doc.selectSheet("تشویقی");
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            availTashvighi += doc.read(r,4).toInt();
        }
        if(availTashvighi > m_combosMap.value("variablesOptions").at(2).toInt()){
            availTashvighi = m_combosMap.value("variablesOptions").at(2).toInt();
        }
        if(availTashvighi < (usedTashvighi + values.at(6).toInt())){
            emit showLetterWarning("تشویقی باقی مانده کافی نیست و وارد خلا تشویقی می شود");
        }
        db.selectSheet(tab);
        bool boomi = false;
        bool mojarad = false;
        QString ostan;
        for (int r = 2; r <= db.currentWorksheet()->dimension().rowCount(); r++){
            if(code == db.read(r,3).toString()){
                for(int c = 1; c <= db.currentWorksheet()->dimension().columnCount(); c++){
                    if(db.read(1,c).toString() == "وضعیت بومی"){
                        if(db.read(r,c).toString() == "بومی"){
                            boomi = true;
                        }
                    }
                    if(db.read(1,c).toString() == "وضعیت تاهل"){
                        if(db.read(r,c).toString() == "مجرد"){
                            mojarad = true;
                        }
                    }
                    if(db.read(1,c).toString() == "استان"){
                        ostan = db.read(r,c).toString();
                    }
                }
                break;
            }
        }
        int maxTorahi = -1;
        int maxMultiplier = 0;
        int maxTorahiMultiplier = 0;
        QFile ostanFile (m_ostanFilename);
        if(!ostanFile.exists()){
            emit showLetterWarning("فایل مربوط به استان ها یافت نشد");
            return false;
        }
        QXlsx::Document ostanList(m_ostanFilename);
        ostanList.selectSheet(0);
        for(int r = 2; r <= ostanList.currentWorksheet()->dimension().rowCount(); r++){
            if(ostanList.read(r,1).toString() == ostan){
                maxTorahi = Soldier::PerToEngNo(ostanList.read(r,2).toString()).toInt();
                break;
            }
        }
        if(maxTorahi == -1){
            emit showLetterWarning("استان سرباز در فایل استان ها یافت نشد");
            return false;
        }
        if(boomi){
            maxMultiplier = m_combosMap.value("variablesOptions").at(6).toInt();
        } else {
            maxMultiplier = m_combosMap.value("variablesOptions").at(5).toInt();
        }
        if(mojarad){
            maxTorahiMultiplier = m_combosMap.value("variablesOptions").at(3).toInt();
        } else {
            maxTorahiMultiplier = m_combosMap.value("variablesOptions").at(4).toInt();
        }
        maxTorahi = maxTorahi * maxTorahiMultiplier;
        int maxSaliane = std::round(m_combosMap.value("variablesOptions").at(0).toDouble()*maxMultiplier);
        int maxEstelaji = std::round(m_combosMap.value("variablesOptions").at(1).toDouble()*maxMultiplier);
        doc.selectSheet("کسری");
        int kasri = 0;
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            kasri += Soldier::PerToEngNo(doc.read(r,4).toString()).toInt();
        }
        doc.selectSheet("کسری منطقه عملیاتی");
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            QString start = Soldier::PerToEngNo(doc.read(r,4).toString());
            QString end = Soldier::PerToEngNo(doc.read(r,5).toString());
            kasri += std::round(DateCalc::dateToDays(DateCalc::minusTwoDate(start,end))/m_combosMap.value("multipliesOptions").at(5).toDouble());
        }
        maxSaliane = std::round(maxSaliane - ((m_combosMap.value("variablesOptions").at(0).toDouble()*kasri)/30));
        maxEstelaji = std::round(maxEstelaji - ((m_combosMap.value("variablesOptions").at(1).toDouble()*kasri)/30));
        if(summaryRequest){
            QList<int> list;
            list.append(maxSaliane);
            list.append(availTashvighi);
            list.append(maxTorahi);
            list.append(maxEstelaji);
            list.append(usedSaliane);
            list.append(usedTashvighi);
            list.append(usedTorahi);
            list.append(usedEstelaji);
            list.append(maxSaliane - usedSaliane);
            list.append(availTashvighi - usedTashvighi);
            list.append(maxTorahi - usedTorahi);
            list.append(maxEstelaji - usedEstelaji);
            list.append(kasri);
            list.append(maxMultiplier);
            emit SoldierFiles::summaryList(list);
            return true;
        }
        if(maxTorahi < (usedTorahi + values.at(7).toInt())){
            emit showLetterWarning("توراهی باقی مانده کافی نیست و وارد خلا توراهی می شود");
        }
        if(maxSaliane < (usedSaliane + values.at(5).toInt())){
            emit showLetterWarning("سالیانه باقی مانده کافی نیست و وارد خلا سالیانه می شود");
        }
        if(maxEstelaji < (usedEstelaji + values.at(9).toInt())){
            emit showLetterWarning("استعلاجی باقی مانده کافی نیست و وارد خلا استعلاجی می شود");
        }
        if(!values.at(values.size()-1).isEmpty()){
            return true;
        }
        if(values.at(9).toInt() > 0){
            QStringList orderValues;
            orderValues.append(values.at(3));
            orderValues.append(values.at(4));
            orderValues.append(values.at(9));
            emit addToOrder(tab,code,"استراحت پزشکی",orderValues);
        }
        for (int r = 2; r <= db.currentWorksheet()->dimension().rowCount(); r++){
            if(code == db.read(r,3).toString()){
                for(int c = 1; c <= db.currentWorksheet()->dimension().columnCount(); c++){
                    if(db.read(1,c).toString() == "وضعیت حضور"){
                        emit changeColumnData(c-1,"مرخصی");
                        break;
                    }
                }
                break;
            }
        }
    } else if(letterType == "شروع ماموریت"){
        doc.selectSheet("نهست");
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            if(!doc.currentWorksheet()->rowFormat(r).patternBackgroundColor().isValid() && (values.at(values.size()-1).isEmpty())){
                emit showLetterWarning("نهست باز دارد");
                return false;
            }
        }
        doc.selectSheet("فرار ۱۵ روز");
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            if(!doc.currentWorksheet()->rowFormat(r).patternBackgroundColor().isValid() && (values.at(values.size()-1).isEmpty())){
                emit showLetterWarning("فرار باز دارد");
                return false;
            }
        }
        doc.selectSheet("فرار ۶ ماه");
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            if(!doc.currentWorksheet()->rowFormat(r).patternBackgroundColor().isValid() && (values.at(values.size()-1).isEmpty())){
                emit showLetterWarning("فرار باز دارد");
                return false;
            }
        }
        if(values.at(4).trimmed().isEmpty() || values.at(5).trimmed().isEmpty()){
            emit showLetterWarning("محل عزیمت و شماره امریه باید ثبت شود");
            return false;
        }
        QStringList starts;
        QStringList ends;
        QStringList sheets;
        sheets.append("مراجعت");
        sheets.append("مرخصی");
        sheets.append("مراجعت از فرار ۱۵ روز");
        sheets.append("مراجعت از فرار ۶ ماه");
        for(QString &sheet: sheets){
            doc.selectSheet(sheet);
            for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
                starts.append(Soldier::PerToEngNo(doc.read(r,4).toString()));
                ends.append(Soldier::PerToEngNo(doc.read(r,5).toString()));
            }
        }
        QString end = DateCalc::addDaysToDate(Soldier::PerToEngNo(values.at(3)),1);
        if(DateCalc::isCrossover(Soldier::PerToEngNo(values.at(3)),end,starts,ends)){
            emit showLetterWarning(letterType + " به دلیل تداخل با بازه‌های ثبت شده، قابل ثبت نیست ");
            return false;
        }
        if(!values.at(values.size()-1).isEmpty()){
            return true;
        }
        for (int r = 2; r <= db.currentWorksheet()->dimension().rowCount(); r++){
            if(code == db.read(r,3).toString()){
                for(int c = 1; c <= db.currentWorksheet()->dimension().columnCount(); c++){
                    if(db.read(1,c).toString() == "وضعیت حضور"){
                        emit changeColumnData(c-1,"ماموریت");
                        break;
                    }
                }
                break;
            }
        }
    } else if(letterType == "پایان ماموریت"){
        if(values.at(5).trimmed().isEmpty() || values.at(6).trimmed().isEmpty()){
            emit showLetterWarning("محل عزیمت و شماره امریه باید ثبت شود");
            return false;
        }
        doc.selectSheet("شروع ماموریت");
        int row = 0;
        bool isOpenMission = false;
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            if(!doc.currentWorksheet()->rowFormat(r).patternBackgroundColor().isValid()){
                isOpenMission = true;
                if(doc.currentWorksheet()->read(r,4).toString() == values.at(3)){
                    row = r;
                    break;
                }
            }
        }
        if(!isOpenMission){
            emit showLetterWarning("شروع ماموریت این بازه قابل دسترسی نیست یا ثبت نشده است");
            return false;
        } else if(row < 1){
            emit showLetterWarning("عدم تطابق با تاریخ شروع ماموریت");
            return false;
        }
        emit changeRowColor(doc.currentWorksheet()->sheetName(),row,Qt::gray);
        if(!values.at(values.size()-1).isEmpty()){
            return true;
        }
        QStringList orderValues;
        orderValues.append(values.at(3));
        orderValues.append(values.at(4));
        orderValues.append(values.at(5));
        orderValues.append(values.at(6));
        emit addToOrder(tab,code,"ماموریت",orderValues);
        for (int r = 2; r <= db.currentWorksheet()->dimension().rowCount(); r++){
            if(code == db.read(r,3).toString()){
                for(int c = 1; c <= db.currentWorksheet()->dimension().columnCount(); c++){
                    if(db.read(1,c).toString() == "وضعیت حضور"){
                        emit changeColumnData(c-1,"حضور");
                        break;
                    }
                }
                break;
            }
        }
    } else if(letterType == "کان لم یکن"){
        QString start = Soldier::PerToEngNo(values.at(3));
        QString end = Soldier::PerToEngNo(values.at(4));
        QStringList starts;
        QStringList ends;
        QStringList sheets;
        sheets.append(letterType);
        for(QString &sheet: sheets){
            doc.selectSheet(sheet);
            for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
                starts.append(Soldier::PerToEngNo(doc.read(r,4).toString()));
                ends.append(Soldier::PerToEngNo(doc.read(r,5).toString()));
            }
        }
        if(DateCalc::isCrossover(start,end,starts,ends)){
            emit showLetterWarning(letterType + " به دلیل تداخل با بازه‌های ثبت شده، قابل ثبت نیست ");
            return false;
        }
        doc.selectSheet("مراجعت");
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            QString nahast = Soldier::PerToEngNo(doc.read(r,4).toString());
            QString morajeat = Soldier::PerToEngNo(doc.read(r,5).toString());
            if ((DateCalc::isMinusValid(nahast,start)) && (DateCalc::isMinusValid(end,morajeat))){
                emit changeRowColor(doc.currentWorksheet()->sheetName(),r,Qt::gray);
                if(!values.at(values.size()-1).isEmpty()){
                    return true;
                }
                if(values.at(9).toInt() > 0){
                    QStringList orderValues;
                    orderValues.append(values.at(3));
                    orderValues.append(values.at(4));
                    orderValues.append(values.at(9));
                    emit addToOrder(tab,code,"استراحت پزشکی",orderValues);
                }
                if(doc.read(r,doc.currentWorksheet()->dimension().columnCount()).toString() == m_combosMap.value("Order").at(0)){
                    QStringList changeValues;
                    changeValues.append(doc.read(r,4).toString());
                    changeValues.append(doc.read(r,5).toString());
                    QXlsx::Document orderDoc(QDir::currentPath()+QDir::separator()+m_orderDir+QDir::separator()+ m_combosMap.value("Order").at(0)+".xlsx");
                    orderDoc.selectSheet(doc.currentWorksheet()->sheetName());
                    bool match = false;
                    int row = 0;
                    for(int r = 2; r <= orderDoc.currentWorksheet()->dimension().rowCount(); r++){
                        if(match){
                            break;
                        }
                        if(orderDoc.read(r,4).toString() != code){
                            continue;
                        }
                        for(int c = 0; c < changeValues.size(); c++){
                            if(orderDoc.read(r,c+5).toString() != changeValues.at(c)){
                                break;
                            }
                            if(c == changeValues.size() -1){
                                match = true;
                                row = r;
                            }
                        }
                    }
                    if(row > 1){
                        int newValue = (Soldier::PerToEngNo(orderDoc.read(row,7).toString()).toInt() - DateCalc::dateToDays(DateCalc::minusTwoDate(start,end)));
                        emit changeInOrder(doc.currentWorksheet()->sheetName(),changeValues,newValue);
                    }
                }
                return true;
            }
        }
        emit showLetterWarning("مراجعتی برای بازه ورودی ثبت نشده است");
        return false;
    } else if(letterType == "کسر و اضافه به آمار"){
        db.selectSheet(tab);
        for (int r = 2; r <= db.currentWorksheet()->dimension().rowCount(); r++){
            if(code == db.read(r,3).toString()){
                for(int c = 1; c <= db.currentWorksheet()->dimension().columnCount(); c++){
                    if(db.read(1,c).toString() == "قسمت"){
                        if(values.at(4) != db.read(r,c).toString()){
                            emit showLetterWarning("قسمت کسرکننده، قسمت فعلی سرباز نیست");
                            return false;
                        } else {
                            emit changeColumnData(c-1,values.at(5));
                            return true;
                        }
                    }
                }
            }
        }
    } else if(letterType == "شورا پزشکی"){
        db.selectSheet(tab);
        for (int r = 2; r <= db.currentWorksheet()->dimension().rowCount(); r++){
            if(code == db.read(r,3).toString()){
                for(int c = 1; c <= db.currentWorksheet()->dimension().columnCount(); c++){
                    if(db.read(1,c).toString() == "وضعیت جسمانی"){
                        emit changeColumnData(c-1,values.at(3));
                    } if(db.read(1,c).toString() == "وضعیت روانی"){
                        emit changeColumnData(c-1,values.at(4));
                    }
                }
                break;
            }
        }
    } else if(letterType == "نوبت شورا"){
        db.selectSheet(tab);
        for (int r = 2; r <= db.currentWorksheet()->dimension().rowCount(); r++){
            if(code == db.read(r,3).toString()){
                for(int c = 1; c <= db.currentWorksheet()->dimension().columnCount(); c++){
                    if(db.read(1,c).toString() == "وضعیت جسمانی"){
                        emit changeColumnData(c-1,"نوبت شورا");
                    } if(db.read(1,c).toString() == "وضعیت روانی"){
                        emit changeColumnData(c-1,"نوبت شورا");
                    }
                }
                break;
            }
        }
    } else if(letterType == "فرار ۱۵ روز"){
        doc.selectSheet("نهست");
        int row = 0;
        bool isOpenNahast = false;
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            if(!doc.currentWorksheet()->rowFormat(r).patternBackgroundColor().isValid()){
                isOpenNahast = true;
                if(doc.currentWorksheet()->read(r,4).toString() == values.at(3)){
                    row = r;
                    break;
                }
            }
        }
        if(!isOpenNahast){
            emit showLetterWarning("نهست این بازه قابل دسترسی نیست یا ثبت نشده است");
            return false;
        } else if(row < 1){
            emit showLetterWarning("عدم تطابق با تاریخ نهست");
            return false;
        }
        QString startDate = Soldier::PerToEngNo(values.at(3));
        QString endDate = Soldier::PerToEngNo(values.at(4));
        if(DateCalc::minusTwoDate(startDate,endDate) != "0000/00/15"){
            emit showLetterWarning("تاریخ فرار باید دقیقا ۱۵ روز پس از نهست باشد");
            return false;
        }
        emit changeRowColor(doc.currentWorksheet()->sheetName(),row,Qt::gray);
        if(!values.at(values.size()-1).isEmpty()){
            return true;
        }
        db.selectSheet(tab);
        for (int r = 2; r <= db.currentWorksheet()->dimension().rowCount(); r++){
            if(code == db.read(r,3).toString()){
                for(int c = 1; c <= db.currentWorksheet()->dimension().columnCount(); c++){
                    if(db.read(1,c).toString() == "وضعیت حضور"){
                        emit changeColumnData(c-1,"فرار");
                        break;
                    }
                }
                break;
            }
        }
    } else if(letterType == "فرار ۶ ماه"){
        doc.selectSheet("فرار ۱۵ روز");
        int row = 0;
        bool isOpenFarar = false;
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            if(!doc.currentWorksheet()->rowFormat(r).patternBackgroundColor().isValid()){
                isOpenFarar = true;
                if(doc.currentWorksheet()->read(r,4).toString() == values.at(3)){
                    row = r;
                    break;
                }
            }
        }
        if(!isOpenFarar){
            emit showLetterWarning("فرار ۱۵ روز این بازه قابل دسترسی نیست یا ثبت نشده است");
            return false;
        } else if(row < 1){
            emit showLetterWarning("عدم تطابق با تاریخ نهست");
            return false;
        }
        QString startDate = Soldier::PerToEngNo(values.at(3));
        QString endDate = Soldier::PerToEngNo(values.at(4));
        if(DateCalc::minusTwoDate(startDate,endDate) != "0000/06/00"){
            emit showLetterWarning("تاریخ فرار باید دقیقا ۶ ماه پس از نهست باشد");
            return false;
        }
        emit changeRowColor(doc.currentWorksheet()->sheetName(),row,Qt::gray);
        QStringList lastValues = {values.at(3)};
        emit moveToTable(tab,code,letterType,lastValues);
    } else if(letterType == "مراجعت از فرار ۱۵ روز"){
        doc.selectSheet("فرار ۱۵ روز");
        int row = 0;
        bool isOpenFarar = false;
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            if(!doc.currentWorksheet()->rowFormat(r).patternBackgroundColor().isValid()){
                isOpenFarar = true;
                if(doc.currentWorksheet()->read(r,4).toString() == values.at(3)){
                    row = r;
                    break;
                }
            }
        }
        if(!isOpenFarar){
            emit showLetterWarning("فرار ۱۵ روز این بازه قابل دسترسی نیست یا ثبت نشده است");
            return false;
        } else if(row < 1){
            emit showLetterWarning("عدم تطابق با تاریخ نهست");
            return false;
        }
        QString fararDate = Soldier::PerToEngNo(doc.currentWorksheet()->read(row,5).toString());
        QString morajeatDate = Soldier::PerToEngNo(values.at(4));
        if(DateCalc::isMinusValid(fararDate,morajeatDate)){
            emit changeRowColor(doc.currentWorksheet()->sheetName(),row,Qt::gray);
            if(!values.at(values.size()-1).isEmpty()){
                return true;
            }
            db.selectSheet(tab);
            for (int r = 2; r <= db.currentWorksheet()->dimension().rowCount(); r++){
                if(code == db.read(r,3).toString()){
                    for(int c = 1; c <= db.currentWorksheet()->dimension().columnCount(); c++){
                        if(db.read(1,c).toString() == "وضعیت حضور"){
                            emit changeColumnData(c-1,"حضور");
                            break;
                        }
                    }
                    break;
                }
            }
            QStringList lastValues = {values.at(3),values.at(4)};
            emit addToTable(tab,code,"بایگانی فرار ۱۵ روز",lastValues);
            return true;
        } else {
            emit showLetterWarning("تاریخ مراجعت باید پس از تاریخ فرار باشد");
            return false;
        }
    } else if (letterType == "کان لم یکن فرار ۱۵ روز"){
        QString start = Soldier::PerToEngNo(values.at(3));
        QString end = Soldier::PerToEngNo(values.at(4));
        QStringList starts;
        QStringList ends;
        QStringList sheets;
        sheets.append(letterType);
        for(QString &sheet: sheets){
            doc.selectSheet(sheet);
            for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
                starts.append(Soldier::PerToEngNo(doc.read(r,4).toString()));
                ends.append(Soldier::PerToEngNo(doc.read(r,5).toString()));
            }
        }
        if(DateCalc::isCrossover(start,end,starts,ends)){
            emit showLetterWarning(letterType + " به دلیل تداخل با بازه‌های ثبت شده، قابل ثبت نیست ");
            return false;
        }
        doc.selectSheet("مراجعت از فرار ۱۵ روز");
        for (int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            QString nahastDate = Soldier::PerToEngNo(doc.currentWorksheet()->read(r,4).toString());
            QString morajeatDate = Soldier::PerToEngNo(doc.currentWorksheet()->read(r,5).toString());
            bool startCheck = DateCalc::isMinusValid(nahastDate,start);
            bool endCheck = DateCalc::isMinusValid(end,morajeatDate);
            if(startCheck && endCheck){
                emit changeRowColor(doc.currentWorksheet()->sheetName(),r,Qt::gray);
                return true;
            }
        }
        emit showLetterWarning("مراجعت از فراری برای بازه ورودی ثبت نشده است");
        return false;
    } else if(letterType == "مراجعت از فرار ۶ ماه"){
        if(tab == "فرار ۶ ماه"){
            doc.selectSheet(tab);
            int row = 0;
            bool isOpenFarar = false;
            for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
                if(!doc.currentWorksheet()->rowFormat(r).patternBackgroundColor().isValid()){
                    isOpenFarar = true;
                    if(doc.currentWorksheet()->read(r,4).toString() == values.at(3)){
                        row = r;
                        break;
                    }
                }
            }
            if(!isOpenFarar){
                emit showLetterWarning("فرار ۶ ماه این بازه قابل دسترسی نیست یا ثبت نشده است");
                return false;
            } else if(row < 1){
                emit showLetterWarning("عدم تطابق با تاریخ نهست");
                return false;
            }
            QString fararDate = Soldier::PerToEngNo(doc.currentWorksheet()->read(row,5).toString());
            QString morajeatDate = Soldier::PerToEngNo(values.at(4));
            if(DateCalc::isMinusValid(fararDate,morajeatDate)){
                db.selectSheet(tab);
                for (int r = 2; r <= db.currentWorksheet()->dimension().rowCount(); r++){
                    if(code == db.read(r,3).toString()){
                        for(int c = 1; c <= db.currentWorksheet()->dimension().columnCount(); c++){
                            if(db.read(1,c).toString() == "تاریخ معرفی"){
                                emit changeColumnData(c-1,values.at(4));
                                break;
                            }
                        }
                        break;
                    }
                }
                QStringList lastValues = { "حضور" , "" };
                emit changeRowColor(tab,row,Qt::gray);
                QStringList lastValues2 = {values.at(3),values.at(4)};
                if(!values.at(values.size()-1).isEmpty()){
                    emit addToTable(tab,code,"بایگانی فرار ۶ ماه",lastValues2);
                }
                emit SoldierFiles::moveToTable(tab,code,"لیست کارکنان وظیفه",lastValues);
                return true;
            } else {
                emit showLetterWarning("تاریخ مراجعت باید پس از تاریخ فرار باشد");
                return false;
            }
        }
        emit showLetterWarning("سرباز در فرار ۶ ماه نیست");
        return false;
    } else if (letterType == "ترخیص"){
        QStringList lastValues = {values.at(3),"اقدام نگردیده"};
        if(values.at(values.size()-1).isEmpty()){
            QStringList orderValues;
            orderValues.append(values.at(3));
            emit addToOrder(tab,code,letterType,orderValues);
        }
        emit SoldierFiles::moveToTable(tab,code,letterType,lastValues);
    } else if (letterType == "معافیت" || letterType == "معافیت موقت" || letterType == "استخدام"){
        QStringList lastValues = {values.at(3),values.at(4)};
        emit SoldierFiles::moveToTable(tab,code,letterType,lastValues);
    } else if(letterType == "فوتی" || letterType == "خرید خدمت" || letterType == "عفو رهبری"){
        QStringList sheets;
        sheets.append("نهست");
        sheets.append("شروع ماموریت");
        sheets.append("فرار ۱۵ روز");
        sheets.append("فرار ۶ ماه");
        for(QString &sheet: sheets){
            doc.selectSheet(sheet);
            for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
                if(!doc.currentWorksheet()->rowFormat(r).patternBackgroundColor().isValid()){
                    emit changeRowColor(doc.currentWorksheet()->sheetName(),r,Qt::red);
                }
            }
        }
        QStringList lastValues = {values.at(3),values.at(4)};
        emit SoldierFiles::moveToTable(tab,code,letterType,lastValues);
    } else if (letterType == "انتقالی"){
        QStringList lastValues = {values.at(3),values.at(4)};
        emit SoldierFiles::moveToTable(tab,code,letterType,lastValues);
    } else if (letterType == "کسری"){
        doc.selectSheet(letterType);
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            if(values.at(4) == doc.currentWorksheet()->read(r,5).toString()){
                emit showLetterWarning("این نوع کسری برای این فرد قبلا ثبت شده است");
                return false;
            }
        }
    } else if (letterType == "کسری منطقه عملیاتی"){
        QString start = Soldier::PerToEngNo(values.at(3));
        QString end = Soldier::PerToEngNo(values.at(4));
        QStringList starts;
        QStringList ends;
        QStringList sheets;
        sheets.append(letterType);
        for(QString &sheet: sheets){
            doc.selectSheet(sheet);
            for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
                starts.append(Soldier::PerToEngNo(doc.read(r,4).toString()));
                ends.append(Soldier::PerToEngNo(doc.read(r,5).toString()));
            }
        }
        if(DateCalc::isCrossover(start,end,starts,ends)){
            emit showLetterWarning(letterType + " به دلیل تداخل با بازه‌های ثبت شده، قابل ثبت نیست ");
            return false;
        }
    } else if (letterType == "بخشش اضافه خدمت"){
        int oldNahast = 0;
        int nahast = 0;
        int eidNahast = 0;
        int farar = 0;
        int ezaf = 0;
        int baKhedmat = 0;
        int bedoneKhedmat = 0;
        int khala = 0;
        int total = 0;
        int forgiven = 0;
        int oldNahastMulti = m_combosMap.value("multipliesOptions").at(0).toInt();
        int nahastMulti = m_combosMap.value("multipliesOptions").at(1).toInt();
        int eidNahastMulti = m_combosMap.value("multipliesOptions").at(2).toInt();
        int baKhedmatMulti = m_combosMap.value("multipliesOptions").at(3).toInt();
        int bedoneKhedmatMulti = m_combosMap.value("multipliesOptions").at(4).toInt();
        doc.selectSheet("مراجعت");
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            QString start = Soldier::PerToEngNo(doc.read(r,4).toString());
            QString end = Soldier::PerToEngNo(doc.read(r,5).toString());
            if(DateCalc::isMinusValid(start,end)){
                DateCalc::nahastCalc(start,end,total,khala,oldNahast,nahast,eidNahast,oldNahastMulti,nahastMulti,eidNahastMulti,false);
            }
        }
        doc.selectSheet("کان لم یکن");
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            QString start = Soldier::PerToEngNo(doc.read(r,4).toString());
            QString end = Soldier::PerToEngNo(doc.read(r,5).toString());
            if(DateCalc::isMinusValid(start,end)){
                DateCalc::nahastCalc(start,end,total,khala,oldNahast,nahast,eidNahast,oldNahastMulti,nahastMulti,eidNahastMulti,true);
            }
        }
        doc.selectSheet("مراجعت از فرار ۶ ماه");
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            QString start = Soldier::PerToEngNo(doc.read(r,4).toString());
            QString end = Soldier::PerToEngNo(doc.read(r,5).toString());
            if(DateCalc::isMinusValid(start,end)){
                QString diff = DateCalc::minusTwoDate(start,end);
                khala += DateCalc::dateToDays(diff);
                total += DateCalc::dateToDays(diff);
                farar += DateCalc::dateToDays(diff);
            }
        }
        doc.selectSheet("مراجعت از فرار ۱۵ روز");
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            QString start = Soldier::PerToEngNo(doc.read(r,4).toString());
            QString end = Soldier::PerToEngNo(doc.read(r,5).toString());
            if(DateCalc::isMinusValid(start,end)){
                QString diff = DateCalc::minusTwoDate(start,end);
                khala += DateCalc::dateToDays(diff);
                total += DateCalc::dateToDays(diff);
                farar += DateCalc::dateToDays(diff);
                if(doc.currentWorksheet()->rowFormat(r).patternBackgroundColor().name() == QColor(Qt::gray).name()){
                    doc.selectSheet("کان لم یکن فرار ۱۵ روز");
                    QStringList kansStarts;
                    QStringList kansEnds;
                    QStringList notKansStarts;
                    QStringList notKansEnds;
                    for(int kr = 2; kr <= doc.currentWorksheet()->dimension().rowCount(); kr++){
                        QString startKan = Soldier::PerToEngNo(doc.read(kr,4).toString());
                        QString endKan = Soldier::PerToEngNo(doc.read(kr,5).toString());
                        bool startCheck = DateCalc::isMinusValid(start,startKan);
                        bool endCheck = DateCalc::isMinusValid(endKan,end);
                        if(startCheck && endCheck){
                            kansStarts.append(startKan);
                            kansEnds.append(endKan);
                        }
                    }
                    if(!kansStarts.isEmpty()){
                        std::sort(kansStarts.begin(),kansStarts.end(),DateCalc::isMinusValid);
                        std::sort(kansEnds.begin(),kansEnds.end(),DateCalc::isMinusValid);
                    }
                    for(int i = 0; i < kansStarts.size(); i++){
                        if((i == (kansStarts.size()-1)) && (kansEnds.at(i) != end)){
                            notKansStarts.append(kansEnds.at(i));
                            notKansEnds.append(end);
                        }
                        if((i == 0) && (kansStarts.at(i) != start)){
                            notKansStarts.append(start);
                            notKansEnds.append(kansStarts.at(i));
                            continue;
                        } if( i != 0){
                        notKansStarts.append(kansEnds.at(i-1));
                        notKansEnds.append(kansStarts.at(i));
                        }
                    }
                    for(int i = 0; i < notKansStarts.size(); i++){
                        QString notKanDiff = DateCalc::minusTwoDate(notKansStarts.at(i),notKansEnds.at(i));
                        if(DateCalc::dateToDays(notKanDiff) < 15){
                            DateCalc::nahastCalc(notKansStarts.at(i),notKansEnds.at(i),total,khala,oldNahast,nahast,eidNahast,oldNahastMulti,nahastMulti,eidNahastMulti,false);
                            khala -= DateCalc::dateToDays(notKanDiff);
                            total -= DateCalc::dateToDays(notKanDiff);
                            farar -= DateCalc::dateToDays(notKanDiff);
                        }
                    }
                    doc.selectSheet("مراجعت از فرار ۱۵ روز");
                }
            }
        }
        doc.selectSheet("کان لم یکن فرار ۱۵ روز");
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            QString start = Soldier::PerToEngNo(doc.read(r,4).toString());
            QString end = Soldier::PerToEngNo(doc.read(r,5).toString());
            if(DateCalc::isMinusValid(start,end)){
                QString diff = DateCalc::minusTwoDate(start,end);
                khala -= DateCalc::dateToDays(diff);
                farar -= DateCalc::dateToDays(diff);
                total -= DateCalc::dateToDays(diff);
            }
        }
        doc.selectSheet("اضافه خدمت");
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            ezaf += Soldier::PerToEngNo(doc.read(r,4).toString()).toInt();
            total += Soldier::PerToEngNo(doc.read(r,4).toString()).toInt();
        }
        doc.selectSheet("بازداشت با خدمت");
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            baKhedmat += doc.read(r,4).toInt();
            khala += Soldier::PerToEngNo(doc.read(r,4).toString()).toInt();
            total += (Soldier::PerToEngNo(doc.read(r,4).toString()).toInt()*baKhedmatMulti);
        }
        doc.selectSheet("بازداشت بدون خدمت");
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            bedoneKhedmat += Soldier::PerToEngNo(doc.read(r,4).toString()).toInt();
            khala += Soldier::PerToEngNo(doc.read(r,4).toString()).toInt();
            total += (Soldier::PerToEngNo(doc.read(r,4).toString()).toInt()*bedoneKhedmatMulti);
        }
        doc.selectSheet(letterType);
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            forgiven += Soldier::PerToEngNo(doc.read(r,4).toString()).toInt();
            total -= Soldier::PerToEngNo(doc.read(r,4).toString()).toInt();
        }
        if(summaryRequest){
            QList<int> list;
            list.append(oldNahast);
            list.append(nahast);
            list.append(eidNahast);
            list.append(farar);
            list.append(ezaf);
            list.append(baKhedmat);
            list.append(bedoneKhedmat);
            list.append(khala);
            list.append(total+forgiven);
            list.append(forgiven);
            emit summaryList(list);
            return true;
        }
        if(values.at(3).toInt() > total){
            emit showLetterWarning("بخشش اضافه خدمت بیش از اضافه خدمت سرباز است");
            return false;
        } else if(values.at(3).toInt() > (total-khala)){
            emit showLetterWarning("بخشش اضافه خدمت نباید شامل خلا خدمتی شود");
            return false;
        }
    } else if (letterType == "بخشش سنوات"){
        db.selectSheet(tab);
        int sanavatColumn = 1;
        for(int c = 1; c < db.currentWorksheet()->dimension().columnCount(); c++){
            if(db.currentWorksheet()->read(1,c).toString() == "سنوات"){
                sanavatColumn = c;
                break;
            }
        }
        int sanavat = 0;
        int forgiven = 0;
        doc.selectSheet(letterType);
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            forgiven += Soldier::PerToEngNo(doc.currentWorksheet()->read(r,4).toString()).toInt();
        }
        for(int r = 2; r <= db.currentWorksheet()->dimension().rowCount(); r++){
            if(code == db.currentWorksheet()->read(r,3).toString()){
                sanavat = db.currentWorksheet()->read(r,sanavatColumn).toInt();
                if(summaryRequest){
                    break;
                }
                if(db.currentWorksheet()->read(r,sanavatColumn).toInt() < (values.at(3).toInt() + forgiven)){
                    emit showLetterWarning("بخشش سنوات بیش از سنوات سرباز است");
                    return false;
                }
                break;
            }
        }
        if(summaryRequest){
            QList<int> list;
            list.append(sanavat);
            list.append(forgiven);
            emit summaryList(list);
        }
    } else if (letterType == "وضعیت درخواست کارت"){
        if(tab == "ترخیص"){
            db.selectSheet(tab);
            for (int r = 2; r <= db.currentWorksheet()->dimension().rowCount(); r++){
                if(code == db.read(r,3).toString()){
                    for(int c = 1; c <= db.currentWorksheet()->dimension().columnCount(); c++){
                        if(db.read(1,c).toString() == letterType){
                            emit changeColumnData(c-1,values.at(3));
                            return true;
                        }
                    }
                }
            }
        }
        emit showLetterWarning("سرباز ترخیص نشده است");
        return false;
    } else if (letterType == "تشویقی"){
        if(!values.at(values.size()-1).isEmpty()){
            return true;
        }
        QStringList orderValues;
        orderValues.append(values.at(3));
        orderValues.append(values.at(4));
        emit addToOrder(tab,code,letterType,orderValues);
    } else if (letterType == "اضافه خدمت" || letterType == "بازداشت با خدمت" || letterType == "بازداشت بدون خدمت"){
        if(!values.at(values.size()-1).isEmpty()){
            return true;
        }
        QStringList orderValues;
        orderValues.append(values.at(3));
        orderValues.append(values.at(4));
        orderValues.append(letterType);
        emit addToOrder(tab,code,"تنبیهی",orderValues);
    }
    return true;
}
