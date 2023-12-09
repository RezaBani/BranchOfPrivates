#include "mainwindow.h"
#include "./ui_mainwindow.h"
#include "addeditdialog.h"
#include "filterdialog.h"
#include "optiondialog.h"
#include "soldier.h"
#include "soldierfiles.h"
#include "letterdialog.h"
#include "orderdialog.h"
#include "statisticsdialog.h"

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    this->setWindowTitle("شعبه افراد");

    //Setup Tabs ui
    tabWidget = new QTabWidget(this);
    tabWidget->setTabPosition(QTabWidget::South);
    this->setCentralWidget(tabWidget);
    tabWidget->setLayoutDirection(Qt::LayoutDirection::RightToLeft);
    ui->toolBar->setContextMenuPolicy(Qt::PreventContextMenu);
    loadOptions();
    this->m_saved = false;
    this->m_skipSyncMessage = false;
    this->m_skipSummaryMessage = false;

    //Setup Tables
    QStringList tableHeader = Soldier::titles();
    QStringList tabNames = Soldier::tabs();
    QStringList tabCodes = Soldier::tabs(true);
    tables.reserve(tabNames.count());
    std::fill_n(std::back_inserter(tables), tabNames.count(), nullptr);
    for(int i = 0; i < tabNames.count(); i++){
        QStringList currentHeader = Soldier::addLastTitles(tableHeader,tabNames.at(i),tabCodes.at(i));
        tables[i] = new QTableWidget(0,currentHeader.count(),this);
        tables.at(i)->setObjectName("tableWidget"+QString::number(i));
        tables.at(i)->setHorizontalHeaderLabels(currentHeader);
        tables.at(i)->setAlternatingRowColors(true);
        tables.at(i)->setEditTriggers(QAbstractItemView::NoEditTriggers);
        tables.at(i)->setSelectionMode(QAbstractItemView::SingleSelection);
        tables.at(i)->resizeColumnsToContents();
        tabWidget->addTab(tables.at(i),tabNames.at(i));
        QScrollBar *customScrollBar = new QScrollBar(Qt::Horizontal,tables.at(i));
        tables.at(i)->setHorizontalScrollBar(customScrollBar);
        customScrollBar->setLayoutDirection(Qt::LeftToRight);
        customScrollBar->setInvertedAppearance(true);
    }

    //Setup ToolBar
    ui->toolBar->setLayoutDirection(Qt::LayoutDirection::RightToLeft);
    txtSearch = new QLineEdit(ui->toolBar);
    txtSearch->setFixedWidth(180);
    txtSearch->setClearButtonEnabled(true);
    QAction *actionSearch = new QAction("جستجو", txtSearch);
    actionSearch->setIcon(QIcon(":/icons/search.png"));
    txtSearch->addAction(actionSearch, QLineEdit::LeadingPosition);
    txtSearch->setPlaceholderText("جستجو...");
    ui->toolBar->addWidget(txtSearch);
    ui->toolBar->addSeparator();
    QLabel *labelDate = new QLabel("تاریخ روز: ",ui->toolBar);
    currentDate = new QDateEdit(ui->toolBar);
    currentDate->setCalendar(Soldier::jalali());
    currentDate->setCalendarPopup(true);
    currentDate->setMinimumWidth(currentDate->width());
    Soldier::setCalenderWidgetFormat(currentDate->calendarWidget());
    currentDate->setDisplayFormat(Soldier::displayFormat());
    currentDate->setLocale(Soldier::persian());
    currentDate->setDate(QDate::currentDate());
    ui->toolBar->addWidget(labelDate);
    ui->toolBar->addWidget(currentDate);
    ui->toolBar->addSeparator();
    QAction *actionAddSoldier = new QAction("افزودن سرباز", this);
    actionAddSoldier->setIcon(QIcon(":/icons/add.png"));
    QAction *actionRemoveSoldier = new QAction("حذف سرباز", this);
    actionRemoveSoldier->setIcon(QIcon(":/icons/minus.png"));
    QAction *actionEditSoldier = new QAction("ویرایش سرباز", this);
    actionEditSoldier->setIcon(QIcon(":/icons/edit.png"));
    QAction *actionSave = new QAction("ذخیره کردن تغییرات", this);
    actionSave->setIcon(QIcon(":/icons/save.png"));
    actionSave->setObjectName("ActionSave");
    QAction *actionLoad = new QAction("بازکردن فایل", this);
    actionLoad->setIcon(QIcon(":/icons/load.png"));
    QAction *actionHide = new QAction("مخفی کردن ستون", this);
    actionHide->setIcon(QIcon(":/icons/hide.png"));
    QAction *actionShow = new QAction("نمایش تمامی ستون‌ها", this);
    actionShow->setIcon(QIcon(":/icons/view.png"));
    QAction *actionFilter = new QAction("فیلتر", this);
    actionFilter->setIcon(QIcon(":/icons/filter.png"));
    QAction *actionClear = new QAction("پاک کردن فیلترها", this);
    actionClear->setIcon(QIcon(":/icons/clear.png"));
    QAction *actionScreen = new QAction("ذخیره کردن جدول نمایش داده شده", this);
    actionScreen->setIcon(QIcon(":/icons/screenshot.png"));
    QAction *actionLetter = new QAction("نامه های بایگانی", this);
    actionLetter->setIcon(QIcon(":/icons/envelope.png"));
    QAction *actionOrder = new QAction("دستور", this);
    actionOrder->setIcon(QIcon(":/icons/papers.png"));
    QAction *actionStatistics = new QAction("آمار پایه خدمتی", this);
    actionStatistics->setIcon(QIcon(":/icons/table.png"));
    QAction *actionSync = new QAction("به روز رسانی", this);
    actionSync->setIcon(QIcon(":/icons/refresh.png"));
    QAction *actionSyncAll = new QAction("به روز رسانی همه", this);
    actionSyncAll->setIcon(QIcon(":/icons/refresh.png"));
    QToolButton *toolBtnSync = new QToolButton(ui->toolBar);
    toolBtnSync->setIcon(QIcon(":/icons/refresh.png"));
    toolBtnSync->setPopupMode(QToolButton::InstantPopup);
    toolBtnSync->addAction(actionSync);
    toolBtnSync->addAction(actionSyncAll);
    QAction *actionSummary = new QAction("خلاصه وضعیت", this);
    actionSummary->setIcon(QIcon(":/icons/list.png"));
    QAction *actionOption = new QAction("تنظیمات", this);
    actionOption->setIcon(QIcon(":/icons/settings.png"));
    QAction *actionHelp = new QAction("راهنما", this);
    actionHelp->setIcon(QIcon(":/icons/help.png"));
    QAction *actionExit = new QAction("خروج", this);
    actionExit->setIcon(QIcon(":/icons/cancel.png"));
    ui->toolBar->addAction(actionAddSoldier);
    ui->toolBar->addAction(actionRemoveSoldier);
    ui->toolBar->addAction(actionEditSoldier);
    ui->toolBar->addSeparator();
    ui->toolBar->addAction(actionSave);
    ui->toolBar->addAction(actionLoad);
    ui->toolBar->addSeparator();
    ui->toolBar->addAction(actionHide);
    ui->toolBar->addAction(actionShow);
    ui->toolBar->addAction(actionFilter);
    ui->toolBar->addAction(actionClear);
    ui->toolBar->addAction(actionScreen);
    ui->toolBar->addSeparator();
    ui->toolBar->addAction(actionLetter);
    ui->toolBar->addWidget(toolBtnSync);
    ui->toolBar->addAction(actionSummary);
    ui->toolBar->addSeparator();
    ui->toolBar->addAction(actionOrder);
    ui->toolBar->addAction(actionStatistics);
    ui->toolBar->addSeparator();
    ui->toolBar->addAction(actionOption);
    ui->toolBar->addAction(actionHelp);
    ui->toolBar->addAction(actionExit);

    //Setup Statusbar
    statusbarLabel = new QLabel (ui->statusbar);
    ui->statusbar->setLayoutDirection(Qt::RightToLeft);
    ui->statusbar->addWidget(statusbarLabel,100);

    //Connecting Signals and Slots
    connect(actionSearch,&QAction::triggered,this,&MainWindow::actionSearchTriggered);
    connect(actionAddSoldier,&QAction::triggered,this,&MainWindow::actionAddTriggered);
    connect(actionRemoveSoldier,&QAction::triggered,this,&MainWindow::actionRemoveTriggered);
    connect(actionEditSoldier,&QAction::triggered,this,&MainWindow::actionEditTriggered);
    connect(actionSave,&QAction::triggered,this,&MainWindow::actionSaveTriggered);
    connect(actionLoad,&QAction::triggered,this,&MainWindow::actionLoadTriggered);
    connect(actionHide,&QAction::triggered,this,&MainWindow::actionHideTriggered);
    connect(actionShow,&QAction::triggered,this,&MainWindow::actionShowTriggered);
    connect(actionLetter,&QAction::triggered,this,&MainWindow::actionLetterTriggered);
    connect(actionExit,&QAction::triggered,this,&QMainWindow::close);
    connect(txtSearch,&QLineEdit::returnPressed,this,&MainWindow::actionSearchTriggered);
    connect(actionOption,&QAction::triggered,this,&MainWindow::actionOptionTriggered);
    connect(actionFilter,&QAction::triggered,this,&MainWindow::actionFilterTriggered);
    connect(actionClear,&QAction::triggered,this,&MainWindow::actionClearTriggered);
    connect(actionHelp,&QAction::triggered,this,&MainWindow::actionHelpTriggered);
    connect(actionScreen,&QAction::triggered,this,&MainWindow::actionScreenTriggered);
    connect(actionSync,&QAction::triggered,this,&MainWindow::actionSyncTriggered);
    connect(actionSyncAll,&QAction::triggered,this,&MainWindow::actionSyncAllTriggered);
    connect(actionSummary,&QAction::triggered,this,&MainWindow::actionSummaryTriggered);
    connect(actionOrder,&QAction::triggered,this,&MainWindow::actionOrderTriggered);
    connect(actionStatistics,&QAction::triggered,this,&MainWindow::actionStatisticsTriggered);
    connect(tabWidget,&QTabWidget::currentChanged,this,&MainWindow::updateStatusBar);
    for(int i = 0; i < tables.size(); i++){
        connect(tables.at(i),&QTableWidget::cellDoubleClicked,this,&MainWindow::cellDoubleClicked);
    }

    //Loading Tables
    actionLoadTriggered();
    emit tabWidget->currentChanged(0);
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::actionAddTriggered()
{
    AddEditDialog *addDlg = new AddEditDialog(m_codes,m_combosMap,tabWidget->currentIndex(),-1,this);
    addDlg->exec();
    QStringList peroperties = addDlg->person();
    delete addDlg;
    if(peroperties.isEmpty()) return;
    QTableWidget *currentTable = qobject_cast<QTableWidget*>(tabWidget->currentWidget());
    currentTable->insertRow(currentTable->rowCount());
    for(int i = 0; i < peroperties.count(); i++){
        QTableWidgetItem *item = new QTableWidgetItem(peroperties.at(i));
        currentTable->setItem(currentTable->rowCount()-1,i,item);
        currentTable->item(currentTable->rowCount()-1,i)->setTextAlignment(Qt::AlignCenter);
    }
    for(int i = 0; i < currentTable->columnCount()-peroperties.count(); i++){
        QTableWidgetItem *item = new QTableWidgetItem();
        currentTable->setItem(currentTable->rowCount()-1,i+peroperties.count(),item);
        currentTable->item(currentTable->rowCount()-1,i+peroperties.count())->setTextAlignment(Qt::AlignCenter);
    }
    currentTable->resizeColumnsToContents();
    if(!tabWidget->tabText(tabWidget->currentIndex()).contains("بایگانی")){
        m_codes.insert(peroperties.at(2),QString(QString::number(tabWidget->currentIndex())+","+QString::number(currentTable->rowCount()-1)));
        m_pendingCreatCodes.insert(peroperties.at(2),tabWidget->tabText(tabWidget->currentIndex()));
    }
    updateStatusBar(tabWidget->currentIndex());
    m_saved = false;
    QMessageBox::StandardButton saveRequestResult = saveRequest();
    if (saveRequestResult == QMessageBox::StandardButton::Save){
        actionSaveTriggered();
        return;
    } else {
        return;
    }
}

void MainWindow::actionRemoveTriggered()
{
    QTableWidget *currentTable = qobject_cast<QTableWidget *>(tabWidget->currentWidget());
    QString title = "حذف سرباز";
    if(!currentTable->item(currentTable->currentRow(),0)){
        QMessageBox::warning(this, title , "ابتدا سربازی را برای حذف کردن انتخاب کنید");
        return;
    }
    QString message = "";
    message += "آیا مطمئن هستید که میخواهید سرباز با مشخصات زیر را حذف کنید؟";
    for(int i = 0; i < currentTable->columnCount(); i++){
        message += "\n";
        message += currentTable->horizontalHeaderItem(i)->text() + ": ";
        message += currentTable->item(currentTable->currentRow(),i)->text();
    }
    QMessageBox::StandardButtons btns = (QMessageBox::Yes | QMessageBox::No);
    QMessageBox *msg = new QMessageBox(QMessageBox::Question, title,message,btns,this);
    msg->setWindowIcon(QIcon(":/icons/minus.png"));
    msg->setFont(this->font());
    msg->setTextInteractionFlags(Qt::TextSelectableByMouse|Qt::TextSelectableByKeyboard);
    int answer = msg->exec();
    delete msg;
    switch (answer) {
    case QMessageBox::Yes:
        if(!tabWidget->tabText(tabWidget->currentIndex()).contains("بایگانی")){
            m_codes.remove(currentTable->item(currentTable->currentRow(),2)->text());
        }
        currentTable->removeRow(currentTable->currentRow());
        currentTable->resizeColumnsToContents();
        m_saved = false;
        updateStatusBar(tabWidget->currentIndex());
        break;
    case QMessageBox::No:
        break;
    default:
        break;
    }
    QMessageBox::StandardButton saveRequestResult = saveRequest();
    if (saveRequestResult == QMessageBox::StandardButton::Save){
        actionSaveTriggered();
        return;
    } else {
        return;
    }
}

void MainWindow::actionEditTriggered()
{
    QTableWidget *currentTable = qobject_cast<QTableWidget *>(tabWidget->currentWidget());
    QString title = "ویرایش سرباز";
    if(!currentTable->item(currentTable->currentRow(),0)){
        QMessageBox::warning(this, title , "ابتدا سربازی را برای ویرایش کردن انتخاب کنید");
        return;
    }
    if(!m_pendingChangeCode.isEmpty()){
        QMessageBox::warning(this, title , "کد ملی یک سرباز تغییر پیدا کرده است، برای ادامه ابتدا تغییرات قبلی را ذخیره کنید");
        return;
    }
    QString message = "";
    message += "آیا مطمئن هستید که میخواهید سرباز با مشخصات زیر را ویرایش کنید؟";
    for(int i = 0; i < currentTable->columnCount(); i++){
        message += "\n";
        message += currentTable->horizontalHeaderItem(i)->text() + ": ";
        message += currentTable->item(currentTable->currentRow(),i)->text();
    }
    QMessageBox::StandardButtons btns = (QMessageBox::Yes | QMessageBox::No);
    QMessageBox *msg = new QMessageBox(QMessageBox::Question, title, message, btns, this);
    msg->setWindowIcon(QIcon(":/icons/edit.png"));
    msg->setFont(this->font());
    msg->setTextInteractionFlags(Qt::TextSelectableByMouse|Qt::TextSelectableByKeyboard);
    int answer = msg->exec();
    delete msg;
    if(answer == QMessageBox::Yes){
        QStringList peroperties;
        peroperties.reserve(currentTable->columnCount());
        std::fill_n(std::back_inserter(peroperties), currentTable->columnCount(), "");
        for(int i = 0; i < peroperties.size(); i++){
            peroperties[i] = currentTable->item(currentTable->currentRow(),i)->text();
        }
        AddEditDialog *editDlg = new AddEditDialog(m_codes,m_combosMap,tabWidget->currentIndex(),currentTable->currentRow(),this,peroperties);
        int answer = editDlg->exec();
        peroperties = editDlg->person();
        delete editDlg;
        if(answer == QDialog::Rejected){
            return;
        }
        if(currentTable->item(currentTable->currentRow(),2)->text() != peroperties.at(2)){
            QString oldCode = currentTable->item(currentTable->currentRow(),2)->text();
            QString oldFilename = QDir::currentPath()+QDir::separator()+tabWidget->tabText(tabWidget->currentIndex())+QDir::separator()+oldCode+".xlsx";
            QString newFilename = QDir::currentPath()+QDir::separator()+tabWidget->tabText(tabWidget->currentIndex())+QDir::separator()+peroperties.at(2)+".xlsx";
            QFile oldFile (oldFilename);
            if(oldFile.exists()){
                m_pendingChangeCode.append(oldFilename);
                m_pendingChangeCode.append(newFilename);
            }
            if(!tabWidget->tabText(tabWidget->currentIndex()).contains("بایگانی")){
                m_codes.remove(currentTable->item(currentTable->currentRow(),2)->text());
                m_codes.insert(peroperties.at(2),QString(QString::number(tabWidget->currentIndex())+","+QString::number(currentTable->currentRow())));
            }
        }
        for(int i = 0; i < peroperties.size(); i++){
            QTableWidgetItem *item = new QTableWidgetItem(peroperties.at(i));
            currentTable->setItem(currentTable->currentRow(),i,item);
            currentTable->item(currentTable->currentRow(),i)->setTextAlignment(Qt::AlignCenter);
        }
        currentTable->resizeColumnsToContents();
        actionSyncTriggered();
        m_saved = false;
    }
    if(answer == QMessageBox::No){
        return;
    }
    QMessageBox::StandardButton saveRequestResult = saveRequest();
    if (saveRequestResult == QMessageBox::StandardButton::Save){
        actionSaveTriggered();
        return;
    } else {
        return;
    }
}

void MainWindow::actionSaveTriggered()
{
    if(sender()->objectName() == "ActionSave"){
        QMessageBox::StandardButton saveRequestResult = saveRequest();
        if (saveRequestResult != QMessageBox::StandardButton::Save){
            return;
        }
    }
    if(!m_pendingCreatCodes.isEmpty()){
        for(auto it = m_pendingCreatCodes.begin(); it != m_pendingCreatCodes.end(); it++){
            SoldierFiles::createExcel(it.value(),it.key(),this->font());
        }
        m_pendingCreatCodes.clear();
    }
    if(!m_pendingChangeCode.isEmpty()){
        QFile::copy(m_pendingChangeCode.at(0),m_pendingChangeCode.at(1));
        QFile::remove(m_pendingChangeCode.at(0));
        m_pendingChangeCode.clear();
    }
    if(!m_saved){
        QFile file(m_dbFilename);
        if(!file.open(QIODevice::WriteOnly)){
            QMessageBox::critical(this,"Error Saving File",file.errorString());
            return;
        }
        file.close();

        QXlsx::Document doc(m_dbFilename);
        for(int t = 0; t < tables.count(); t++){
            doc.addSheet(tabWidget->tabText(t));
            doc.selectSheet(t);
            for(int h = 0; h < tables.at(t)->columnCount(); h++){
                doc.write(1,h+1,QVariant(tables.at(t)->horizontalHeaderItem(h)->text()));
                QXlsx::Format f = doc.columnFormat(h+1);
                f.setFont(this->font());
                f.setHorizontalAlignment(QXlsx::Format::HorizontalAlignment::AlignHCenter);
                f.setVerticalAlignment(QXlsx::Format::VerticalAlignment::AlignVCenter);
                doc.setColumnFormat(h+1,f);
                for(int r = 0; r < tables.at(t)->rowCount(); r++){
                    doc.write(r+2,h+1,tables.at(t)->item(r,h)->text());
                }
            }
        }
        doc.save();
        for(QString &sheetName: doc.sheetNames()){
            doc.selectSheet(sheetName);
            doc.autosizeColumnWidth();
        }
        doc.save();
        m_saved = true;
    }
    if(!m_pendingLetter.isEmpty()){
        for(auto it = m_pendingLetter.begin(); it != m_pendingLetter.end(); it++){
            QDir dir(QDir::currentPath()+QDir::separator()+it.key());
            if (!dir.exists()){
                return;
            }
            bool fileExist = false;
            QFileInfoList infoList = dir.entryInfoList(QDir::Filter::Files | QDir::Filter::NoDotAndDotDot);
            for(const QFileInfo &fi:infoList){
                if(fi.completeBaseName()+".xlsx" == it.value().at(0)+".xlsx"){
                    fileExist = true;
                    break;
                }
            }
            if(!fileExist){
                return;
            }
            QXlsx::Document doc(QDir::currentPath()+QDir::separator()+it.key()+QDir::separator()+it.value().at(0)+".xlsx");
            doc.selectSheet(it.value().at(1));
            int row = doc.currentWorksheet()->dimension().lastRow()+1;
            for(int c = 2; c < it.value().size(); c++){
                doc.write(row,c-1,it.value().at(c));
                QXlsx::Format f = doc.columnFormat(c-1);
                f.setFont(this->font());
                f.setHorizontalAlignment(QXlsx::Format::HorizontalAlignment::AlignHCenter);
                f.setVerticalAlignment(QXlsx::Format::VerticalAlignment::AlignVCenter);
                doc.setColumnFormat(c-1,f);
            }
            doc.save();
            for(QString &sheetName: doc.sheetNames()){
                doc.selectSheet(sheetName);
                doc.autosizeColumnWidth();
            }
            doc.save();
        }
        m_pendingLetter.clear();
    }
    if(!m_pendingRemoveLetter.isEmpty()){
        for(auto it = m_pendingRemoveLetter.begin(); it != m_pendingRemoveLetter.end(); it++){
            QDir dir(QDir::currentPath()+QDir::separator()+it.key());
            if (!dir.exists()){
                return;
            }
            bool fileExist = false;
            QFileInfoList infoList = dir.entryInfoList(QDir::Filter::Files | QDir::Filter::NoDotAndDotDot);
            for(const QFileInfo &fi:infoList){
                if(fi.completeBaseName()+".xlsx" == it.value().at(0)+".xlsx"){
                    fileExist = true;
                    break;
                }
            }
            if(!fileExist){
                return;
            }
            QString oldFilename = QDir::currentPath()+QDir::separator()+it.key()+QDir::separator()+it.value().at(0)+".xlsx";
            QString newFilename = QDir::currentPath()+QDir::separator()+it.value().at(0)+".xlsx";
            QFile::copy(oldFilename,newFilename);
            QFile::remove(oldFilename);
            xlnt::workbook wb;
            xlnt::path path (QString(it.value().at(0)+".xlsx").toStdString());
            wb.load(path);
            xlnt::worksheet ws = wb.sheet_by_index(it.value().at(1).toInt());
            ws.delete_rows(it.value().at(2).toInt(),1);
            wb.save(path);
            QFile::copy(newFilename,oldFilename);
            QFile::remove(newFilename);
        }
        m_pendingRemoveLetter.clear();
    }
    if(!m_pendingOrder.isEmpty()){
        for(auto it = m_pendingOrder.begin(); it != m_pendingOrder.end(); it++){
            QDir dir(QDir::currentPath()+QDir::separator()+m_orderDir);
            if (!dir.exists()){
                return;
            }
            bool fileExist = false;
            QFileInfoList infoList = dir.entryInfoList(QDir::Filter::Files | QDir::Filter::NoDotAndDotDot);
            for(const QFileInfo &fi:infoList){
                if(fi.completeBaseName()+".xlsx" == m_combosMap.value("Order").at(0)+".xlsx"){
                    fileExist = true;
                    break;
                }
            }
            if(!fileExist){
                return;
            }
            QXlsx::Document doc(QDir::currentPath()+QDir::separator()+m_orderDir+QDir::separator()+ m_combosMap.value("Order").at(0)+".xlsx");
            doc.selectSheet(it.key());
            int row = doc.currentWorksheet()->dimension().lastRow()+1;
            for(int c = 0; c < it.value().size(); c++){
                doc.write(row,c+1,it.value().at(c));
                QXlsx::Format f = doc.columnFormat(c+1);
                f.setFont(this->font());
                f.setHorizontalAlignment(QXlsx::Format::HorizontalAlignment::AlignHCenter);
                f.setVerticalAlignment(QXlsx::Format::VerticalAlignment::AlignVCenter);
                doc.setColumnFormat(c+1,f);
            }
            doc.save();
            for(QString &sheetName: doc.sheetNames()){
                doc.selectSheet(sheetName);
                doc.autosizeColumnWidth();
            }
            doc.save();
        }
        m_pendingOrder.clear();
    }
    if(!m_pendingRemoveOrder.isEmpty()){
        for(auto it = m_pendingRemoveOrder.begin(); it != m_pendingRemoveOrder.end(); it++){
            QDir dir(QDir::currentPath()+QDir::separator()+m_orderDir);
            if (!dir.exists()){
                return;
            }
            bool fileExist = false;
            QFileInfoList infoList = dir.entryInfoList(QDir::Filter::Files | QDir::Filter::NoDotAndDotDot);
            for(const QFileInfo &fi:infoList){
                if(fi.completeBaseName()+".xlsx" == m_combosMap.value("Order").at(0)+".xlsx"){
                    fileExist = true;
                    break;
                }
            }
            if(!fileExist){
                return;
            }
            QXlsx::Document doc(QDir::currentPath()+QDir::separator()+m_orderDir+QDir::separator()+ m_combosMap.value("Order").at(0)+".xlsx");
            doc.selectSheet(it.key());
            bool match = false;
            int row = 0;
            for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
                if(match){
                    break;
                }
                for(int c = 0; c < it.value().size(); c++){
                    if(doc.read(r,c+1).toString() != it.value().at(c)){
                        break;
                    }
                    if(c == it.value().size() -1){
                        match = true;
                        row = r;
                    }
                }
            }
            QString oldFilename = QDir::currentPath()+QDir::separator()+m_orderDir+QDir::separator()+m_combosMap.value("Order").at(0)+".xlsx";
            QString newFilename = QDir::currentPath()+QDir::separator()+m_combosMap.value("Order").at(0)+".xlsx";
            QFile::copy(oldFilename,newFilename);
            QFile::remove(oldFilename);
            xlnt::workbook wb;
            xlnt::path path (QString(m_combosMap.value("Order").at(0)+".xlsx").toStdString());
            wb.load(path);
            xlnt::worksheet ws = wb.sheet_by_index(Soldier::orders().indexOf(it.key()));
            ws.delete_rows(row,1);
            wb.save(path);
            QFile::copy(newFilename,oldFilename);
            QFile::remove(newFilename);
        }
        m_pendingRemoveOrder.clear();
    }
    if(!m_pendingChangeOrder.isEmpty()){
        for(auto it = m_pendingChangeOrder.begin(); it != m_pendingChangeOrder.end(); it++){
            QDir dir(QDir::currentPath()+QDir::separator()+m_orderDir);
            if (!dir.exists()){
                return;
            }
            bool fileExist = false;
            QFileInfoList infoList = dir.entryInfoList(QDir::Filter::Files | QDir::Filter::NoDotAndDotDot);
            for(const QFileInfo &fi:infoList){
                if(fi.completeBaseName()+".xlsx" == m_combosMap.value("Order").at(0)+".xlsx"){
                    fileExist = true;
                    break;
                }
            }
            if(!fileExist){
                return;
            }
            QXlsx::Document doc(QDir::currentPath()+QDir::separator()+m_orderDir+QDir::separator()+ m_combosMap.value("Order").at(0)+".xlsx");
            doc.selectSheet(it.key());
            bool match = false;
            int row = 0;
            for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
                if(match){
                    break;
                }
                for(int c = 1; c < doc.currentWorksheet()->dimension().columnCount(); c++){
                    if(doc.read(r,c).toString() != it.value().at(c-1)){
                        break;
                    }
                    if( c == it.value().size() - 1){
                        match = true;
                        row = r;
                    }
                }
            }
            for(int c = 0; c < it.value().size(); c++){
                doc.write(row,c+1,it.value().at(c));
                QXlsx::Format f = doc.columnFormat(c+1);
                f.setFont(this->font());
                f.setHorizontalAlignment(QXlsx::Format::HorizontalAlignment::AlignHCenter);
                f.setVerticalAlignment(QXlsx::Format::VerticalAlignment::AlignVCenter);
                doc.setColumnFormat(c+1,f);
            }
            doc.save();
            for(QString &sheetName: doc.sheetNames()){
                doc.selectSheet(sheetName);
                doc.autosizeColumnWidth();
            }
            doc.save();
        }
        m_pendingChangeOrder.clear();
    }
    if(!m_pendingRowColor.isEmpty()){
        for(auto it = m_pendingRowColor.begin(); it != m_pendingRowColor.end(); it++){
            QDir dir(QDir::currentPath()+QDir::separator()+it.key());
            if (!dir.exists()){
                return;
            }
            bool fileExist = false;
            QFileInfoList infoList = dir.entryInfoList(QDir::Filter::Files | QDir::Filter::NoDotAndDotDot);
            for(const QFileInfo &fi:infoList){
                if(fi.completeBaseName()+".xlsx" == it.value().at(0)+".xlsx"){
                    fileExist = true;
                    break;
                }
            }
            if(!fileExist){
                return;
            }
            QXlsx::Document doc(QDir::currentPath()+QDir::separator()+it.key()+QDir::separator()+it.value().at(0)+".xlsx");
            doc.selectSheet(it.value().at(1));
            QXlsx::Format f = doc.currentWorksheet()->rowFormat(it.value().at(2).toInt());
            if(it.value().at(3).toInt() == -1){
                f.setPatternBackgroundColor(QColor(QColor::Invalid));
            } else {
                QColor color = QColor(Qt::GlobalColor(it.value().at(3).toInt()));
                f.setPatternBackgroundColor(color);
            }
            doc.currentWorksheet()->setRowFormat(it.value().at(2).toInt(),it.value().at(2).toInt(),f);
            doc.save();
        }
        m_pendingRowColor.clear();
    }
    if(!m_pendingMove.isEmpty()){
        for(auto it = m_pendingMove.begin(); it != m_pendingMove.end(); it++){
            QDir oldDir(QDir::currentPath()+QDir::separator()+it.value().at(0));
            QDir newDir(QDir::currentPath()+QDir::separator()+it.value().at(1));
            if (!oldDir.exists()){
                return;
            }
            bool fileExist = false;
            QFileInfoList infoList = oldDir.entryInfoList(QDir::Filter::Files | QDir::Filter::NoDotAndDotDot);
            for(const QFileInfo &fi:infoList){
                if(fi.completeBaseName()+".xlsx" == it.key()+".xlsx"){
                    fileExist = true;
                    break;
                }
            }
            if(!fileExist){
                return;
            }
            if(!newDir.exists()){
                QDir dir (QDir::currentPath());
                dir.mkdir(it.value().at(1));
            }
            QString oldFilename = QDir::currentPath()+QDir::separator()+it.value().at(0)+QDir::separator()+it.key()+".xlsx";
            QString newFilename = QDir::currentPath()+QDir::separator()+it.value().at(1)+QDir::separator()+it.key()+".xlsx";
            QFile::copy(oldFilename,newFilename);
            QFile::remove(oldFilename);
        }
        m_pendingMove.clear();
    }
}

void MainWindow::actionLoadTriggered()
{
    if(sender()){
        QMessageBox::StandardButton answer = saveRequest();
        if(answer == QMessageBox::StandardButton::Save){
            actionSaveTriggered();
        } else if (answer == QMessageBox::StandardButton::Cancel){
            return;
        }
        m_codes.clear();
    }
    QFile file(m_dbFilename);
    if(!file.exists()){
        return;
    }
    if(!file.open(QIODevice::ReadOnly)){
        QMessageBox::critical(this,"Error Loading File",file.errorString());
        return;
    }
    file.close();
    QStringList tabs = Soldier::tabs();
    QXlsx::Document doc(m_dbFilename);
    for(int t = 0; t < tables.count(); t++){
        doc.selectSheet(t);
        tables.at(t)->setRowCount(doc.currentWorksheet()->dimension().rowCount()-1);
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            for(int c = 1; c <= doc.currentWorksheet()->dimension().columnCount(); c++){
                QTableWidgetItem *item = new QTableWidgetItem(doc.read(r,c).toString());
                tables.at(t)->setItem(r-2,c-1,item);
                tables.at(t)->item(r-2,c-1)->setTextAlignment(Qt::AlignCenter);
            }
            if(!tabs.at(t).contains("بایگانی")){
                m_codes.insert(tables.at(t)->item(r-2,2)->text(),QString(QString::number(t)+","+QString::number(r-2)));
                SoldierFiles::createExcel(tabWidget->tabText(t),tables.at(t)->item(r-2,2)->text(),this->font());
            }
        }
        tables.at(t)->resizeColumnsToContents();
    }
    m_saved = true;
    m_pendingLetter.clear();
    m_pendingRemoveLetter.clear();
    m_pendingMove.clear();
    m_pendingRowColor.clear();
    m_pendingOrder.clear();
    m_pendingChangeOrder.clear();
    m_pendingRemoveOrder.clear();
    m_pendingCreatCodes.clear();
    m_pendingChangeCode.clear();
    updateStatusBar(tabWidget->currentIndex());
}

void MainWindow::actionHideTriggered()
{
    QTableWidget *currentTable = qobject_cast<QTableWidget *>(tabWidget->currentWidget());
    if(!currentTable->item(0,currentTable->currentColumn())){
        QString title = "مخفی کردن ستون";
        QMessageBox::warning(this, title , "ابتدا ستونی را برای مخفی کردن انتخاب کنید");
        return;
    }
    QSet<int> columns;
    for(int i = 0; i < currentTable->selectedItems().size(); i++){
        columns.insert(currentTable->selectedItems().at(i)->column());
    }
    for(int j = 0; j < currentTable->columnCount(); j++){
        if (columns.contains(j)){
            columns.remove(j);
            currentTable->setColumnHidden(j,true);
        }
    }
}

void MainWindow::actionShowTriggered()
{
    QTableWidget *currentTable = qobject_cast<QTableWidget *>(tabWidget->currentWidget());
    for(int j = 0; j < currentTable->columnCount(); j++){
        currentTable->setColumnHidden(j,false);
    }
}

void MainWindow::actionOptionTriggered()
{
    OptionDialog *optDlg = new OptionDialog(this,false);
    int answer = optDlg->exec();
    if(answer == QDialog::Accepted){
        m_combosMap = optDlg->combosMap();
    }
    delete optDlg;
    loadOptions();
}

void MainWindow::actionLetterTriggered()
{
    QString title = "افزودن نامه";
    QString message;
    QTableWidget *currentTable = qobject_cast<QTableWidget *>(tabWidget->currentWidget());
    if(!currentTable->item(currentTable->currentRow(),0)){
        message = "ابتدا سربازی را برای اضافه کردن نامه انتخاب کنید";
        QMessageBox::warning(this, title, message);
        return;
    }
    if(tabWidget->tabText(tabWidget->currentIndex()).contains("بایگانی")){
        message = "افزودن نامه برای بخش بایگانی فرار امکانپذیر نیست";
        QMessageBox::critical(this, title, message);
        return;
    }
    if(!m_pendingChangeCode.isEmpty()){
        message = "کد ملی یک سرباز تغییر پیدا کرده است، برای ادامه ابتدا تغییرات قبلی را ذخیره کنید";
        QMessageBox::warning(this, title, message);
        QMessageBox::StandardButton saveRequestResult = saveRequest();
        if (saveRequestResult == QMessageBox::StandardButton::Save){
            actionSaveTriggered();
        } else {
            return;
        }
    }
    if(!m_pendingLetter.isEmpty() || !m_pendingRemoveLetter.isEmpty() || !m_pendingMove.isEmpty() || !m_pendingRowColor.isEmpty()){
        message = "ابتدا نامه وارد شده قبلی را ذخیره کنید سپس نامه جدیدی اضافه کنید";
        QMessageBox::warning(this, title, message);
        QMessageBox::StandardButton saveRequestResult = saveRequest();
        if (saveRequestResult == QMessageBox::StandardButton::Save){
            actionSaveTriggered();
        } else {
            return;
        }
    }
    actionSyncTriggered();
    LetterDialog *ltrDlg = new LetterDialog(tabWidget->tabText(tabWidget->currentIndex()),currentTable->item(currentTable->currentRow(),2)->text(),m_combosMap,this);
    connect(ltrDlg,&LetterDialog::changeColumnData,this,&MainWindow::changeColumnData);
    connect(ltrDlg,&LetterDialog::addLetter,this,&MainWindow::addLetter);
    connect(ltrDlg,&LetterDialog::removeLetter,this,&MainWindow::removeLetter);
    connect(ltrDlg,&LetterDialog::changeRowColor,this,&MainWindow::changeRowColor);
    connect(ltrDlg,&LetterDialog::moveToTable,this,&MainWindow::moveToTable);
    connect(ltrDlg,&LetterDialog::addToTable,this,&MainWindow::addToTable);
    connect(ltrDlg,&LetterDialog::addToOrder,this,&MainWindow::addToOrder);
    connect(ltrDlg,&LetterDialog::changeInOrder,this,&MainWindow::changeInOrder);
    connect(ltrDlg,&LetterDialog::removeOrder,this,&MainWindow::removeOrder);
    int answer = ltrDlg->exec();
    delete ltrDlg;
    if(answer == QDialog::Rejected){
        return;
    }
    QMessageBox::StandardButton saveRequestResult = saveRequest();
    if (saveRequestResult == QMessageBox::StandardButton::Save){
        actionSaveTriggered();
    }
}

void MainWindow::actionFilterTriggered()
{
    FilterDialog *filterDlg = new FilterDialog(tabWidget->currentIndex(), m_combosMap, this);
    int answer = filterDlg->exec();
    if(answer == QDialog::Rejected){
        delete filterDlg;
        return;
    }
    QStringList filteredValues = filterDlg->filteredValues();
    int filteredColumn = filterDlg->filteredColumn();
    delete filterDlg;
    QTableWidget *table = tables.at(tabWidget->currentIndex());
    for(int r = 0; r < table->rowCount(); r++){
        if(!table->isRowHidden(r)){
            bool hide = true;
            QTableWidgetItem *item = table->item(r,filteredColumn);
            for(QString &value: filteredValues){
                if(item->text().contains(value)){
                    hide = false;
                    break;
                }
            }
            table->setRowHidden(r,hide);
        }
    }
    updateStatusBar(tabWidget->currentIndex());
}

void MainWindow::actionClearTriggered()
{
    txtSearch->clear();
    actionSearchTriggered();
}

void MainWindow::actionHelpTriggered()
{
    QString title = "درباره برنامه";
    QString message;
    message += "Branch of Privates v1.0 Build Date: 1402/04/29";
    message += "\n";
    message += "Qt v5.15.2 (MSVC 2019 64-bit) https://www.qt.io";
    message += "\n";
    message += "QXlsx v1.4.5 https://github.com/QtExcel/QXlsx";
    message += "\n";
    message += "xlnt v1.5.0 https://github.com/tfussell/xlnt";
    message += "\n";
    message += "Vazirmatn v33.003 https://github.com/rastikerdar/vazirmatn";
    message += "\n";
    message += "\n";
    message += "مجوزهای برنامه:";
    message += "\n";
    message += "مجوز";
    message += " GPLv3 ";
    message += "شرایط مجوزهای استفاده شده در برنامه شعبه افراد و کتابخانه‌های استفاده شده بررسی کنید";
    message += "\n";
    message += "\n";
    message += "شرایط استفاده از کد برنامه:";
    message += "\n";
    message += "این برنامه به صورت شخصی و با کمک افراد و کتابخانه‌های نرم‌افزاری ذکرشده، به امید اینکه مفید باشد، طراحی شده است و با مجوز";
    message += " GPLv3 ";
    message += "انتشار یافته است";
    message += "\n";
    message += "مجوز جی پی ال به معنی آن است که استفاده از کدهای این برنامه فقط در صورتی امکان‌پذیر است که کدهای تغییرداده شده و به پروژه اضافه شده (به اصطلاح لینک شده) تحت همین مجوز به صورت آزاد منتشر شوند";
    message += "\n";
    message += "انتخاب این مجوز به دلیل استفاده از کدهایی با همین مجوز بوده است و استفاده از کدهای این برنامه در برنامه‌های غیرمتن‌باز از نظر مجوز استفاده‌شده در این پروژه و خصوصا کتابخانه‌های جی‌پی‌ال استفاده شده نادرست است و در صورت استفاده از کدها و کتابخانه‌های جی‌پی‌ال نیاز است منبع و مجوز آن را به دقت در چنین صفحه‌ای در برنامه بیاورید";
    message += "\n";
    message += "\n";
    message += "صحت عملکرد:";
    message += "\n";
    message += "تضمینی در صحت کارکرد برنامه وجود ندارد، سعی شده برنامه‌ای مفید ارائه شود. مشهور است که هیچ برنامه، نرم‌افزار و الگوریتمی بی‌نقص نیست و ممکن است نیاز باشد خروجی برنامه را با منابع دیگر بررسی کنید";
    message += "\n";
    QMessageBox *msg = new QMessageBox(QMessageBox::Information,title,message);
    msg->setTextInteractionFlags(Qt::TextSelectableByMouse|Qt::TextSelectableByKeyboard);
    msg->setWindowIcon(QIcon(":/icons/help.png"));
    msg->setFont(this->font());
    msg->exec();
    delete msg;
}

void MainWindow::actionScreenTriggered()
{
    QString saveFilename = QFileDialog::getSaveFileName(this,"Save File",QString(),"Excel File (*.xlsx);;All File (*.*)");
    if(saveFilename.isEmpty()){
        return;
    }
    QFile file(saveFilename);
    if(file.exists()){
        file.remove();
    }
    QXlsx::Document doc(saveFilename);
    int row = 1;
    int col = 0;
    QTableWidget *table = tables.at(tabWidget->currentIndex());
    for(int h = 0; h < table->columnCount(); h++){
        if(!table->isColumnHidden(h)){
            col++;
            doc.write(row,col,table->horizontalHeaderItem(h)->text());
        }
    }
    col = 0;
    for(int r = 0; r < table->rowCount(); r++){
        if(!table->isRowHidden(r)){
            row++;
            for(int c = 0; c < table->columnCount(); c++){
                if(!table->isColumnHidden(c)){
                    col++;
                    doc.write(row,col,table->item(r,c)->text());
                }
            }
            col = 0;
        }
    }
    for(int r = 1; r <= doc.dimension().columnCount(); r++){
        QXlsx::Format f = doc.rowFormat(r);
        f.setFont(this->font());
        f.setHorizontalAlignment(QXlsx::Format::HorizontalAlignment::AlignHCenter);
        f.setVerticalAlignment(QXlsx::Format::VerticalAlignment::AlignVCenter);
        doc.setRowFormat(r,f);
    }
    doc.save();
    doc.autosizeColumnWidth();
    doc.save();
}

void MainWindow::actionSyncTriggered()
{
    QTableWidget *currentTable = qobject_cast<QTableWidget *>(tabWidget->currentWidget());
    QString title = "بروزرسانی وضعیت";
    QString message;
    if(!currentTable->item(currentTable->currentRow(),0)){
        message = "ابتدا سربازی را برای بروزرسانی وضعیت انتخاب کنید";
        QMessageBox::warning(this, title, message);
        return;
    }
    QString tab = tabWidget->tabText(tabWidget->currentIndex());
    QString code = currentTable->item(currentTable->currentRow(),2)->text();
    QDir dir(QDir::currentPath()+QDir::separator()+tab);
    if (!dir.exists()){
        message =  "پوشه پرونده سربازان یافت نشد";
        if(!m_skipSyncMessage){
            QMessageBox::warning(this, title, message);
        }
        return;
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
        message = "فایل پرونده سرباز یافت نشد";
        if(!m_skipSyncMessage){
            QMessageBox::warning(this, title, message);
        }
        return;
    }
    QXlsx::Document doc(QDir::currentPath()+QDir::separator()+tab+QDir::separator()+code+".xlsx");
    doc.selectSheet("فرار ۶ ماه");
    for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
        if(!doc.currentWorksheet()->rowFormat(r).patternBackgroundColor().isValid()){
            if(tab != doc.currentWorksheet()->sheetName()){
                message = "سرباز باید در فرار ۶ ماه باشد";
                updateRowColor(tabWidget->currentIndex(),currentTable->currentRow(),Qt::red,message);
                return;
            } else {
                updateRowColor(tabWidget->currentIndex(),currentTable->currentRow());
                return;
            }
        }
    }
    int vaziatColumn = -1;
    int jesmani = -1;
    int ravani = -1;
    for (int c = 0; c < currentTable->columnCount(); c++){
        if (currentTable->horizontalHeaderItem(c)->text() == "وضعیت حضور"){
            vaziatColumn = c;
        }
        if (currentTable->horizontalHeaderItem(c)->text() == "وضعیت جسمانی"){
            jesmani = c;
        }
        if (currentTable->horizontalHeaderItem(c)->text() == "وضعیت روانی"){
            ravani = c;
        }
    }
    if(vaziatColumn == -1){
        if(!m_skipSyncMessage){
            QMessageBox::warning(this, title , "سرباز در لیست کارکنان وظیفه نیست");
        }
        return;
    }
    doc.selectSheet("فرار ۱۵ روز");
    for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
        if(!doc.currentWorksheet()->rowFormat(r).patternBackgroundColor().isValid()){
            if(currentTable->item(currentTable->currentRow(),vaziatColumn)->text() != "فرار"){
                message = "وضعیت سرباز از " +currentTable->item(currentTable->currentRow(),vaziatColumn)->text() + " به فرار تغییر کرد";
                changeColumnData(vaziatColumn,"فرار");
                updateRowColor(tabWidget->currentIndex(),currentTable->currentRow(),Qt::cyan,message);
                return;
            }
            QString start = DateCalc::addTwoDate(Soldier::PerToEngNo(doc.read(r,4).toString()),"0000/06/00");
            if(DateCalc::isMinusValid(start,Soldier::PerToEngNo(currentDate->text()))){
                message = "موعد فرار ۶ ماه سرباز رسیده است";
                updateRowColor(tabWidget->currentIndex(),currentTable->currentRow(),Qt::yellow,message);
                return;
            } else {
                updateRowColor(tabWidget->currentIndex(),currentTable->currentRow());
                return;
            }
        }
    }
    bool present = true;
    doc.selectSheet("نهست");
    for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
        if(!doc.currentWorksheet()->rowFormat(r).patternBackgroundColor().isValid()){
            present = false;
            if(currentTable->item(currentTable->currentRow(),vaziatColumn)->text() != doc.currentWorksheet()->sheetName()){
                message = "وضعیت سرباز از " +currentTable->item(currentTable->currentRow(),vaziatColumn)->text() + " به نهست تغییر کرد";
                changeColumnData(vaziatColumn, doc.currentWorksheet()->sheetName());
                updateRowColor(tabWidget->currentIndex(),currentTable->currentRow(),Qt::cyan,message);
                return;
            }
            QString start = DateCalc::addTwoDate(Soldier::PerToEngNo(doc.read(r,4).toString()),"0000/00/15");
            if(DateCalc::isMinusValid(start,Soldier::PerToEngNo(currentDate->text()))){
                message = "موعد فرار ۱۵ روز سرباز رسیده است";
                updateRowColor(tabWidget->currentIndex(),currentTable->currentRow(),Qt::yellow,message);
                return;
            } else {
                updateRowColor(tabWidget->currentIndex(),currentTable->currentRow());
            }
        }
    }
    if(present){
        doc.selectSheet("مرخصی");
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            if(DateCalc::isMinusValid(Soldier::PerToEngNo(currentDate->text()),Soldier::PerToEngNo(doc.read(r,5).toString()))){
                present = false;
                if(currentTable->item(currentTable->currentRow(),vaziatColumn)->text() != doc.currentWorksheet()->sheetName()){
                    message = "وضعیت سرباز از " +currentTable->item(currentTable->currentRow(),vaziatColumn)->text() + " به مرخصی تغییر کرد";
                    changeColumnData(vaziatColumn, doc.currentWorksheet()->sheetName());
                    updateRowColor(tabWidget->currentIndex(),currentTable->currentRow(),Qt::cyan,message);
                    return;
                } else {
                    updateRowColor(tabWidget->currentIndex(),currentTable->currentRow());
                }
            }
        }
    }
    if(present){
        doc.selectSheet("شروع ماموریت");
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            if(!doc.currentWorksheet()->rowFormat(r).patternBackgroundColor().isValid()){
                present = false;
                if(currentTable->item(currentTable->currentRow(),vaziatColumn)->text() != "ماموریت"){
                    message = "وضعیت سرباز از " +currentTable->item(currentTable->currentRow(),vaziatColumn)->text() + " به ماموریت تغییر کرد";
                    changeColumnData(vaziatColumn,"ماموریت");
                    updateRowColor(tabWidget->currentIndex(),currentTable->currentRow(),Qt::cyan,message);
                    return;
                }
                QString start = DateCalc::addTwoDate(Soldier::PerToEngNo(doc.read(r,4).toString()),"0000/02/00");
                if(DateCalc::isMinusValid(start,Soldier::PerToEngNo(currentDate->text()))){
                    message = "بیش از دو ماه از شروع ماموریت سرباز گذشته است";
                    updateRowColor(tabWidget->currentIndex(),currentTable->currentRow(),Qt::yellow,message);
                    return;
                } else {
                    updateRowColor(tabWidget->currentIndex(),currentTable->currentRow());
                }
            }
        }
    }
    if((currentTable->item(currentTable->currentRow(),jesmani)->text() == "نوبت شورا") || (currentTable->item(currentTable->currentRow(),ravani)->text() == "نوبت شورا")){
        doc.selectSheet("نوبت شورا");
        if(!doc.currentWorksheet()->rowFormat(doc.currentWorksheet()->dimension().rowCount()).patternBackgroundColor().isValid()){
            QString date = Soldier::PerToEngNo(doc.read(doc.currentWorksheet()->dimension().rowCount(),4).toString());
            if(DateCalc::isMinusValid(date,Soldier::PerToEngNo(currentDate->text()))){
                QString diff = DateCalc::minusTwoDate(date,Soldier::PerToEngNo(currentDate->text()));
                if(DateCalc::dateToDays(diff) > 90){
                    message = "بیش از سه ماه از نوبت شورا سرباز گذشته است";
                    updateRowColor(tabWidget->currentIndex(),currentTable->currentRow(),Qt::yellow,message);
                    return;
                } else {
                    updateRowColor(tabWidget->currentIndex(),currentTable->currentRow());
                }
            }
        }
    }
    if((currentTable->item(currentTable->currentRow(),vaziatColumn)->text() != "حضور") && (present)){
        message = "وضعیت سرباز از " +currentTable->item(currentTable->currentRow(),vaziatColumn)->text() + " به حضور تغییر کرد";
        changeColumnData(vaziatColumn,"حضور");
        updateRowColor(tabWidget->currentIndex(),currentTable->currentRow(),Qt::cyan,message);
        return;
    } else if (present) {
        updateRowColor(tabWidget->currentIndex(),currentTable->currentRow());
    }
    m_tarkhis.clear();
    m_ghanoni.clear();
    m_mofid.clear();
    m_skipSummaryMessage = true;
    actionSummaryTriggered();
    if(m_tarkhis.isEmpty() || m_mofid.isEmpty() || m_ghanoni.isEmpty()){
        m_skipSummaryMessage = false;
        return;
    }
    int month = Soldier::PerToEngNo(currentDate->text()).split("/").at(1).toInt();
    int year = Soldier::PerToEngNo(currentDate->text()).split("/").at(0).toInt();
    int lowLimit = 23;
    int upLimit = 53;
    if(month == 12){
        month = 2;
        year++;
    } else {
        month++;
    }
    if(month == 12){
        upLimit = 83;
    }
    QString commisionDate = DateCalc::makeStringDate(year,month,1);
    if(DateCalc::isMinusValid(commisionDate,m_tarkhis)){
        if((DateCalc::dateToDays(DateCalc::minusTwoDate(commisionDate,m_tarkhis)) > lowLimit) && (DateCalc::dateToDays(DateCalc::minusTwoDate(commisionDate,m_tarkhis)) < upLimit)){
            message = "در کمیسیون ترخیص بعدی مطرح گردد";
            updateRowColor(tabWidget->currentIndex(),currentTable->currentRow(),Qt::green,message);
            m_skipSummaryMessage = false;
            return;
        }
    }
    m_mofid = DateCalc::addTwoDate(m_mofid,DateCalc::minusTwoDate(Soldier::PerToEngNo(currentDate->text()),commisionDate),false);
    if(DateCalc::dateToDays(m_mofid) < DateCalc::dateToDays(m_ghanoni)){
        if((DateCalc::dateToDays(DateCalc::minusTwoDate(m_mofid,m_ghanoni,false)) > lowLimit) && (DateCalc::dateToDays(DateCalc::minusTwoDate(m_mofid,m_ghanoni,false)) < upLimit)){
            message = "در کمیسیون ترخیص بعدی (خدمت مفید) مطرح گردد";
            updateRowColor(tabWidget->currentIndex(),currentTable->currentRow(),Qt::green,message);
            m_skipSummaryMessage = false;
            return;
        }
    }
    m_skipSummaryMessage = false;
}

void MainWindow::actionSyncAllTriggered()
{
    QString title = "بروزرسانی همه";
    QString message = "آیا می‌خواهید وضعیت همه سربازان را بروز رسانی کنید؟";
    message += "\n";
    message += "این روند ممکن است مدتی طول بکشد";
    QMessageBox::StandardButton answer = QMessageBox::question(this,title,message);
    if(answer == QMessageBox::StandardButton::No){
        return;
    }
    m_skipSyncMessage = true;
    QProgressBar *progressBar = new QProgressBar(this);
    QLabel *refreshLabel = new QLabel("در حال بروزرسانی",this);
    progressBar->setTextVisible(false);
    progressBar->setLayoutDirection(Qt::LeftToRight);
    progressBar->setFixedHeight(ui->statusbar->height()/2);
    refreshLabel->setFont(ui->statusbar->font());
    ui->statusbar->addWidget(progressBar);
    ui->statusbar->addWidget(refreshLabel);
    int maxRow = tables.at(tabWidget->currentIndex())->rowCount();
    for(int r = 0; r < maxRow; r++){
        tables.at(tabWidget->currentIndex())->setCurrentCell(r,0);
        actionSyncTriggered();
        progressBar->setValue(100*((r+1)/double(maxRow)));
    }
    ui->statusbar->removeWidget(progressBar);
    ui->statusbar->removeWidget(refreshLabel);
    delete progressBar;
    delete refreshLabel;
    tables.at(tabWidget->currentIndex())->clearSelection();
    m_skipSyncMessage = false;
}

void MainWindow::actionSummaryTriggered()
{
    QTableWidget *currentTable = qobject_cast<QTableWidget *>(tabWidget->currentWidget());
    QString title = "خلاصه وضعیت";
    QString message;
    if(!currentTable->item(currentTable->currentRow(),0)){
        if(!m_skipSummaryMessage){
            message = "ابتدا سربازی را برای مشاهده خلاصه وضعیت انتخاب کنید";
            QMessageBox::warning(this, title, message);
        }
        return;
    }
    if(tabWidget->tabText(tabWidget->currentIndex()).contains("بایگانی")){
        if(!m_skipSummaryMessage){
            message = "خلاصه وضعیت برای بخش بایگانی فرار امکانپذیر نیست";
            QMessageBox::warning(this, title, message);
        }
        return;
    }
    bool check = false;
    SoldierFiles sfiles (m_combosMap);
    connect(&sfiles,&SoldierFiles::summaryList,this,&MainWindow::summaryList);
    QStringList args;
    args.append("");
    args.append("");
    args.append("");
    args.append("0");
    args.append("0");
    args.append("0");
    args.append("0");
    args.append("0");
    args.append("");
    check = sfiles.validateLetter(tabWidget->tabText(tabWidget->currentIndex()),currentTable->item(currentTable->currentRow(),2)->text(),"مرخصی",args,true);
    if(!check){
        return;
    }
    QString message1 = "";
    message1 += "خلاصه وضعیت: " + currentTable->item(currentTable->currentRow(),0)->text() + " فرزند: " + currentTable->item(currentTable->currentRow(),1)->text();
    message1 += "\n";
    message1 += "تعداد کل سالیانه مجاز: ";
    message1 += QString::number(summaryResult.at(0));
    message1 += "روز\n";
    message1 += "تعداد کل تشویقی مجاز: ";
    message1 += QString::number(summaryResult.at(1));
    message1 += "روز\n";
    message1 += "تعداد کل توراهی مجاز: ";
    message1 += QString::number(summaryResult.at(2));
    message1 += "روز\n";
    message1 += "تعداد کل استعلاجی مجاز: ";
    message1 += QString::number(summaryResult.at(3));
    message1 += "روز\n";
    message1 += "\n";
    message1 += "تعداد کل سالیانه استفاده شده: ";
    message1 += QString::number(summaryResult.at(4));
    message1 += "روز\n";
    message1 += "تعداد کل تشویقی استفاده شده: ";
    message1 += QString::number(summaryResult.at(5));
    message1 += "روز\n";
    message1 += "تعداد کل توراهی استفاده شده: ";
    message1 += QString::number(summaryResult.at(6));
    message1 += "روز\n";
    message1 += "تعداد کل استعلاجی استفاده شده: ";
    message1 += QString::number(summaryResult.at(7));
    message1 += "روز\n";
    message1 += "\n";
    message1 += "تعداد کل سالیانه باقیمانده: ";
    message1 += QString::number(summaryResult.at(8));
    message1 += "روز\n";
    message1 += "تعداد کل تشویقی باقیمانده: ";
    message1 += QString::number(summaryResult.at(9));
    message1 += "روز\n";
    message1 += "تعداد کل توراهی باقیمانده: ";
    message1 += QString::number(summaryResult.at(10));
    message1 += "روز\n";
    message1 += "تعداد کل استعلاجی باقیمانده: ";
    message1 += QString::number(summaryResult.at(11));
    message1 += "روز\n";
    if(!m_skipSummaryMessage){
        QMessageBox *msg1 = new QMessageBox(QMessageBox::Information,title,message1);
        msg1->setTextInteractionFlags(Qt::TextSelectableByMouse|Qt::TextSelectableByKeyboard);
        msg1->setWindowIcon(QIcon(":/icons/list.png"));
        msg1->setFont(this->font());
        msg1->exec();
        delete msg1;
    }

    QString finalDate = currentDate->text();
    QStringList tabs = Soldier::tabs(true);
    int titlesCount = Soldier::titles().size();
    if((tabs.at(tabWidget->currentIndex()).toLower().contains("d")) && (!tabs.at(tabWidget->currentIndex()).toLower().contains("dd"))){
        finalDate = Soldier::PerToEngNo(currentTable->item(currentTable->currentRow(),titlesCount)->text());
    }
    QString message2 = "";
    int salianeKhala = 0;
    int tashvighiKhala = 0;
    int torahiKhala = 0;
    int estelajiKhala = 0;
    if(summaryResult.at(8) < 0){
        salianeKhala = summaryResult.at(8)*-1;
    }
    if(summaryResult.at(9) < 0){
        tashvighiKhala = summaryResult.at(9)*-1;
    }
    if(summaryResult.at(10) < 0){
        torahiKhala = summaryResult.at(10)*-1;
    }
    if(summaryResult.at(11) < 0){
        estelajiKhala = summaryResult.at(11)*-1;
    }
    int offKhala = salianeKhala + tashvighiKhala + torahiKhala + estelajiKhala;
    int kasri = summaryResult.at(12);
    int ghanoni = summaryResult.at(13);
    m_ghanoni = DateCalc::daysToDate(ghanoni*30);
    args.clear();
    args.append("");
    args.append("");
    args.append("");
    args.append("0");
    args.append("");
    check = sfiles.validateLetter(tabWidget->tabText(tabWidget->currentIndex()),currentTable->item(currentTable->currentRow(),2)->text(),"بخشش اضافه خدمت",args,true);
    if(!check){
        return;
    }
    QString dispatch = Soldier::PerToEngNo(currentTable->item(currentTable->currentRow(),3)->text());
    message2 += "تاریخ اعزام: ";
    message2 += dispatch;
    message2 += "\n";
    message2 += "خدمت قانونی: ";
    message2 += QString::number(ghanoni);
    message2 += "ماه\n";
    QString tarkhis = DateCalc::addTwoDate(dispatch,DateCalc::daysToDate(ghanoni*30));
    message2 += "تاریخ ترخیص قانونی: ";
    message2 += tarkhis;
    message2 += "\n";
    message2 += "\n";
    message2 += "خلا سالیانه: ";
    message2 += QString::number(salianeKhala);
    message2 += "روز\n";
    message2 += "خلا تشویقی: ";
    message2 += QString::number(tashvighiKhala);
    message2 += "روز\n";
    message2 += "خلا توراهی: ";
    message2 += QString::number(torahiKhala);
    message2 += "روز\n";
    message2 += "خلا استعلاجی: ";
    message2 += QString::number(estelajiKhala);
    message2 += "روز\n";
    message2 += "\n";
    message2 += "نهست ضریب دو: ";
    message2 += DateCalc::dateDistanceInWords(DateCalc::daysToDate(summaryResult.at(0)));
    message2 += "\n";
    message2 += "نهست ضریب سه: ";
    message2 += DateCalc::dateDistanceInWords(DateCalc::daysToDate(summaryResult.at(1)));
    message2 += "\n";
    message2 += "نهست ضریب چهار: ";
    message2 += DateCalc::dateDistanceInWords(DateCalc::daysToDate(summaryResult.at(2)));
    message2 += "\n";
    message2 += "فرار: ";
    message2 += DateCalc::dateDistanceInWords(DateCalc::daysToDate(summaryResult.at(3)));
    message2 += "\n";
    message2 += "\n";
    message2 += "گروهانی: ";
    message2 += DateCalc::dateDistanceInWords(DateCalc::daysToDate(summaryResult.at(4)));
    message2 += "\n";
    message2 += "زندان با خدمت: ";
    message2 += DateCalc::dateDistanceInWords(DateCalc::daysToDate(summaryResult.at(5)));
    message2 += "\n";
    message2 += "زندان بدون خدمت: ";
    message2 += DateCalc::dateDistanceInWords(DateCalc::daysToDate(summaryResult.at(6)));
    message2 += "\n";
    message2 += "\n";
    message2 += "کل خلا خدمتی: ";
    message2 += DateCalc::dateDistanceInWords(DateCalc::daysToDate(summaryResult.at(7)+offKhala));
    message2 += "\n";
    message2 += "کل اضافه خدمت: ";
    message2 += DateCalc::dateDistanceInWords(DateCalc::daysToDate(summaryResult.at(8)+offKhala));
    message2 += "\n";
    message2 += "کل اضافه خدمت قابل بخشش: ";
    message2 += DateCalc::dateDistanceInWords(DateCalc::daysToDate(summaryResult.at(8) - summaryResult.at(7)));
    message2 += "\n";
    tarkhis = DateCalc::addTwoDate(tarkhis,DateCalc::daysToDate(summaryResult.at(8)+offKhala),true);
    message2 += "تاریخ ترخیص با احتساب کل اضافات خدمتی: ";
    message2 += tarkhis;
    message2 += "\n";
    message2 += "\n";
    message2 += "کل کسری ها: ";
    message2 += DateCalc::dateDistanceInWords(DateCalc::daysToDate(kasri));
    message2 += "\n";
    message2 += "کل بخشش ها: ";
    message2 += DateCalc::dateDistanceInWords(DateCalc::daysToDate(summaryResult.at(9)));
    message2 += "\n";
    tarkhis = DateCalc::minusTwoDate(DateCalc::daysToDate(summaryResult.at(9)+kasri),tarkhis);
    message2 += "تاریخ ترخیص با احتساب کسری ها و بخشش ها: ";
    message2 += tarkhis;
    message2 += "\n";
    message2 += "\n";
    this->m_tarkhis = tarkhis;
    QString mofid = DateCalc::minusTwoDate(dispatch,Soldier::PerToEngNo(finalDate));
    mofid = DateCalc::minusTwoDate(DateCalc::daysToDate(summaryResult.at(7)+offKhala),mofid,false);
    message2 += "خدمت مفید تا تاریخ " + Soldier::PerToEngNo(finalDate) + " : ";
    message2 += DateCalc::dateDistanceInWords(mofid);
    message2 += "\n";
    mofid = DateCalc::addTwoDate(mofid,DateCalc::daysToDate(kasri),false);
    message2 += "خدمت مفید با احتساب کسری ها تا تاریخ " + Soldier::PerToEngNo(finalDate) + " : ";
    message2 += DateCalc::dateDistanceInWords(mofid);
    message2 += "\n";
    message2 += "\n";
    this->m_mofid = mofid;
    args.clear();
    args.append("");
    args.append("");
    args.append("");
    args.append("0");
    args.append("");
    check = sfiles.validateLetter(tabWidget->tabText(tabWidget->currentIndex()),currentTable->item(currentTable->currentRow(),2)->text(),"بخشش سنوات",args,true);
    if(!check){
        return;
    }
    message2 += "سنوات: ";
    message2 += DateCalc::dateDistanceInWords(DateCalc::daysToDate(summaryResult.at(0)));
    message2 += "\n";
    tarkhis = DateCalc::addTwoDate(tarkhis,DateCalc::daysToDate(summaryResult.at(0)));
    message2 += "تاریخ ترخیص با احتساب سنوات: ";
    message2 += tarkhis;
    message2 += "\n";
    message2 += "بخشش سنوات: ";
    message2 += DateCalc::dateDistanceInWords(DateCalc::daysToDate(summaryResult.at(1)));
    message2 += "\n";
    tarkhis = DateCalc::minusTwoDate(DateCalc::daysToDate(summaryResult.at(1)),tarkhis);
    message2 += "تاریخ ترخیص با احتساب بخشش سنوات: ";
    message2 += tarkhis;
    if(!m_skipSummaryMessage){
        QMessageBox *msg2 = new QMessageBox(QMessageBox::Information,title,message2);
        msg2->setTextInteractionFlags(Qt::TextSelectableByMouse|Qt::TextSelectableByKeyboard);
        msg2->setWindowIcon(QIcon(":/icons/list.png"));
        msg2->setFont(this->font());
        msg2->exec();
        delete msg2;
    }
}

void MainWindow::actionOrderTriggered()
{
    QString filePath = SoldierFiles::createOrder(m_orderDir,m_combosMap.value("Order").at(0).toInt(),this->font());
    OrderDialog *ordDlg = new OrderDialog(filePath,m_combosMap.value("Order").at(0).toInt(),this);
    ordDlg->setAttribute(Qt::WA_DeleteOnClose, true);
    ordDlg->open();
}

void MainWindow::actionStatisticsTriggered()
{
    QString title = "آمار پایه خدمتی";
    QString label = "لطفا نوع تاریخ برای تولید آمار پایه خدمتی را وارد کنید";
    QStringList dateTitles;
    for(int i = 0; i < Soldier::titles().size(); i++){
        if(Soldier::titles(true).at(i) == "d"){
            dateTitles.append(Soldier::titles().at(i));
        }
    }
    bool ok = false;
    int dateColumn = 0;
    QString dateType = QInputDialog::getItem(this,title,label,dateTitles,0,false,&ok);
    if(ok){
        dateColumn = Soldier::titles().indexOf(dateType);
    } else {
        return;
    }
    label = "لطفا سال میانی برای تولید آمار پایه خدمتی را وارد کنید";
    QSet<int> years;
    for(int r = 0; r < tables.at(0)->rowCount(); r++){
        years.insert(Soldier::PerToEngNo(tables.at(0)->item(r,dateColumn)->text().split("/").at(0)).toInt());
    }
    if(years.isEmpty()){
        years.insert(1);
    }
    int minYear = *(years.begin());
    int maxYear = *(years.begin());
    for(auto it = years.begin(); it != years.end(); it++){
        minYear = std::min(minYear,*it);
        maxYear = std::max(maxYear,*it);
    }
    ok = false;
    int mainYear = QInputDialog::getInt(this,title,label,maxYear-1,minYear-1,maxYear+1,1,&ok);
    QString path;
    int year = Soldier::PerToEngNo(currentDate->text().split("/").at(0)).toInt();
    int month = Soldier::PerToEngNo(currentDate->text().split("/").at(1)).toInt();
    if(ok){
        path = SoldierFiles::createStatistics(m_statisticsDir,mainYear,year,month,m_combosMap,this->font());
    } else {
        return;
    }
    QMap<QString,int> comboColumns;
    QStringList titles = Soldier::addLastTitles(Soldier::titles(),Soldier::tabs().at(0),Soldier::tabs(true).at(0));
    for(auto it = m_combosMap.begin(); it != m_combosMap.end(); it++){
        if(titles.contains(it.key())){
            comboColumns.insert(it.key(),titles.indexOf(it.key()));
        }
    }
    QXlsx::Document doc(path);
    for(int r = 0; r < tables.at(0)->rowCount(); r++){
        int year = Soldier::PerToEngNo(tables.at(0)->item(r,dateColumn)->text().split("/").at(0)).toInt();
        int month = Soldier::PerToEngNo(tables.at(0)->item(r,dateColumn)->text().split("/").at(1)).toInt();
        int column = month + 1;
        if(year == mainYear){
            column += 12;
        } else if(year > mainYear){
            column += 24;
        }
        for(QString &sheetName:doc.sheetNames()){
            doc.selectSheet(sheetName);
            QString value = tables.at(0)->item(r,comboColumns.value(sheetName))->text();
            int row = m_combosMap.value(sheetName).indexOf(value);
            if(row == -1){
                row = 0;
            }
            row += 3;
            int oldValue = Soldier::PerToEngNo(doc.read(row,column).toString()).toInt();
            doc.write(row,column,QVariant(oldValue + 1));
            doc.save();
        }
    }
    for(QString &sheetName:doc.sheetNames()){
        doc.selectSheet(sheetName);
        for(int r = 3; r < doc.currentWorksheet()->dimension().rowCount(); r++){
            int total = 0;
            for(int c = 2; c < doc.currentWorksheet()->dimension().columnCount(); c++){
                total += Soldier::PerToEngNo(doc.read(r,c).toString()).toInt();
            }
            doc.write(r,doc.currentWorksheet()->dimension().columnCount(),QVariant(total));
        }
        for(int c = 2; c <= doc.currentWorksheet()->dimension().columnCount(); c++){
            int total = 0;
            for(int r = 2; r < doc.currentWorksheet()->dimension().rowCount(); r++){
                total += Soldier::PerToEngNo(doc.read(r,c).toString()).toInt();
            }
            doc.write(doc.currentWorksheet()->dimension().rowCount(),c,QVariant(total));
        }
        doc.save();
    }
    StatisticsDialog *stDlg = new StatisticsDialog(path,year,month,this);
    stDlg->setAttribute(Qt::WA_DeleteOnClose, true);
    stDlg->open();
}

void MainWindow::updateStatusBar(int index)
{
    statusbarLabel->clear();
    QString statusbarText = "";
    statusbarText += "تعداد کل سربازان: ";
    statusbarText += QString::number(tables.at(0)->rowCount());
    statusbarText += " | ";
    int farar = 0;
    int c = 0;
    for (c = 0; c < tables.at(0)->columnCount(); c++){
        if(tables.at(0)->horizontalHeaderItem(c)->text() == "وضعیت حضور"){
            break;
        }
    }
    for (int r = 0; r < tables.at(0)->rowCount(); r++){
        if(tables.at(0)->item(r,c)->text() == "فرار"){
            farar++;
        }
    }
    statusbarText += "تعداد فرار ۱۵ روز: ";
    statusbarText += QString::number(farar);
    statusbarText += " | ";
    size_t shown = 0;
    for (int r = 0; r < tables.at(index)->rowCount(); r++){
        if(!tables.at(index)->isRowHidden(r)){
            shown++;
        }
    }
    statusbarText += "تعداد نفرات نمایش داده شده: ";
    statusbarText += QString::number(shown);
    statusbarLabel->setText(statusbarText);
}

void MainWindow::cellDoubleClicked(int row, int column)
{
    QString currentText = tables.at(tabWidget->currentIndex())->item(row,column)->text();
    QStringList titles = Soldier::titles();
    QStringList titlesCodes = Soldier::titles(true);
    QStringList tabCodes = Soldier::tabs(true);
    titles = Soldier::addLastTitles(titles,tabWidget->tabText(tabWidget->currentIndex()),tabCodes.at(tabWidget->currentIndex()),false);
    titlesCodes = Soldier::addLastTitles(titlesCodes,tabWidget->tabText(tabWidget->currentIndex()),tabCodes.at(tabWidget->currentIndex()),true);
    QString title = "تغییر " + titles.at(column);
    QString label = "برای سرباز";
    label += " " + tables.at(tabWidget->currentIndex())->item(row,0)->text() + " ";
    label += "فرزند";
    label += " " + tables.at(tabWidget->currentIndex())->item(row,1)->text() + " ";
    label += titles.at(column) + " ";
    bool ok = false;
    if(titlesCodes.at(column) == "c"){
        label += "جدیدی انتخاب کنید";
        int currentIndex = m_combosMap.value(titles.at(column)).indexOf(currentText);
        QString newText = QInputDialog::getItem(this,title,label,m_combosMap.value(titles.at(column)),currentIndex,false,&ok);
        if(ok){
            tables.at(tabWidget->currentIndex())->item(row,column)->setText(newText);
            tables.at(tabWidget->currentIndex())->resizeColumnsToContents();
            m_saved = false;
        }
    }
    if(titlesCodes.at(column) == "n"){
        label += "جدیدی وارد کنید";
        int newValue = QInputDialog::getInt(this,title,label,Soldier::PerToEngNo(currentText).toInt(),0,Soldier::maxSpinBox(),1,&ok);
        if(ok){
            tables.at(tabWidget->currentIndex())->item(row,column)->setText(QString::number(newValue));
            tables.at(tabWidget->currentIndex())->resizeColumnsToContents();
            m_saved = false;
        }
    }
    if(titlesCodes.at(column) == "t"){
        label += "جدیدی وارد کنید";
        QString newText = QInputDialog::getText(this,title,label,QLineEdit::Normal,currentText,&ok);
        if(ok){
            if(column == 2){
                if(!m_pendingChangeCode.isEmpty()){
                    QString message = "کد ملی یک سرباز تغییر پیدا کرده است، برای ادامه ابتدا تغییرات قبلی را ذخیره کنید";
                    QMessageBox::warning(this, title, message);
                    QMessageBox::StandardButton saveRequestResult = saveRequest();
                    if (saveRequestResult == QMessageBox::StandardButton::Save){
                        actionSaveTriggered();
                        return;
                    } else {
                        return;
                    }
                }
                newText = Soldier::PerToEngNo(newText);
                if(Soldier::isDuplicate(newText,m_codes,tabWidget->currentIndex(),row)){
                    QString title = "تکراری";
                    QString message = "سربازی با همین کد ملی موجود است";
                    QMessageBox::warning(this,title,message);
                    return;
                } else {
                    QString oldCode = Soldier::PerToEngNo(currentText);
                    QString oldFilename = QDir::currentPath()+QDir::separator()+tabWidget->tabText(tabWidget->currentIndex())+QDir::separator()+oldCode+".xlsx";
                    QString newFilename = QDir::currentPath()+QDir::separator()+tabWidget->tabText(tabWidget->currentIndex())+QDir::separator()+newText+".xlsx";
                    QFile oldFile (oldFilename);
                    if(oldFile.exists()){
                        m_pendingChangeCode.append(oldFilename);
                        m_pendingChangeCode.append(newFilename);
                    }
                    if(!tabWidget->tabText(tabWidget->currentIndex()).contains("بایگانی")){
                        m_codes.remove(oldCode);
                        m_codes.insert(newText,QString(QString::number(tabWidget->currentIndex())+","+QString::number(row)));
                    }
                    tables.at(tabWidget->currentIndex())->item(row,column)->setText(Soldier::PerToEngNo(newText));
                    tables.at(tabWidget->currentIndex())->resizeColumnsToContents();
                    m_saved = false;
                }
            }
            tables.at(tabWidget->currentIndex())->item(row,column)->setText(newText);
            tables.at(tabWidget->currentIndex())->resizeColumnsToContents();
            m_saved = false;
        }
    }
    QMessageBox::StandardButton saveRequestResult = saveRequest();
    if (saveRequestResult == QMessageBox::StandardButton::Save){
        actionSaveTriggered();
        return;
    } else {
        return;
    }
}

void MainWindow::changeColumnData(const int &columnIndex, const QString &columnNewData)
{
    QTableWidget *currentTable = qobject_cast<QTableWidget *>(tabWidget->currentWidget());
    QTableWidgetItem *item = new QTableWidgetItem();
    item->setText(columnNewData);
    item->setTextAlignment(Qt::AlignCenter);
    currentTable->setItem(currentTable->currentRow(),columnIndex,item);
    currentTable->resizeColumnsToContents();
    m_saved = false;
    updateStatusBar(tabWidget->currentIndex());
}

void MainWindow::changeRowColor(const QString &tab, const int &rowIndex, const int &colorCode)
{
    QStringList values;
    QTableWidget *currentTable = qobject_cast<QTableWidget *>(tabWidget->currentWidget());
    values.append(currentTable->item(currentTable->currentRow(),2)->text());
    values.append(tab);
    values.append(QString::number(rowIndex));
    values.append(QString::number(colorCode));
    m_pendingRowColor.insert(tabWidget->tabText(tabWidget->currentIndex()),values);
}

void MainWindow::addLetter(const QString &tab, const QString &code, const QString &letterType, const QStringList &values)
{
    QStringList newValues = values;
    newValues.prepend(letterType);
    newValues.prepend(code);
    if(newValues.last().isEmpty()){
        newValues.replace(newValues.size()-1,m_combosMap.value("Order").at(0));
    }
    m_pendingLetter.insert(tab,newValues);
}

void MainWindow::removeLetter(const QString &tab, const QString &code, const int &letterIndex, const int &row)
{
    QStringList newValues;
    newValues.append(code);
    newValues.append(QString::number(letterIndex));
    newValues.append(QString::number(row));
    m_pendingRemoveLetter.insert(tab,newValues);
}

void MainWindow::moveToTable(const QString &tab, const QString &code, const QString &newTab, const QStringList &lastValues)
{
    QStringList titles = Soldier::titles();
    QTableWidget *currentTable = qobject_cast<QTableWidget *>(tabWidget->currentWidget());
    for(int t = 0; t < tabWidget->count(); t++){
        if(tabWidget->tabText(t) == newTab){
            QTableWidget *destinatioTable = tables.at(t);
            destinatioTable->setRowCount(destinatioTable->rowCount()+1);
            for(int i = 0; i < (titles.size()+lastValues.size()); i++){
                QTableWidgetItem *item = new QTableWidgetItem();
                if(i >= titles.size()){
                    item->setText(lastValues.at(i-titles.size()));
                    item->setTextAlignment(Qt::AlignCenter);
                } else {
                    item->setText(currentTable->item(currentTable->currentRow(),i)->text());
                    item->setTextAlignment(Qt::AlignCenter);
                }
                destinatioTable->setItem(destinatioTable->rowCount()-1,i,item);
            }
            m_codes.remove(code);
            m_codes.insert(code,QString(QString::number(t)+","+QString::number(destinatioTable->rowCount()-1)));
            destinatioTable->resizeColumnsToContents();
            currentTable->removeRow(currentTable->currentRow());
            QStringList movingFiles = {tab,newTab};
            m_pendingMove.insert(code,movingFiles);
            m_saved = false;
            break;
        }
    }
    updateStatusBar(tabWidget->currentIndex());
}

void MainWindow::addToTable(const QString &tab, const QString &code, const QString &newTab, const QStringList &lastValues)
{
    QStringList titles = Soldier::titles();
    QTableWidget *currentTable = qobject_cast<QTableWidget *>(tabWidget->currentWidget());
    for(int t = 0; t < tabWidget->count(); t++){
        if(tabWidget->tabText(t) == newTab){
            QTableWidget *destinatioTable = tables.at(t);
            destinatioTable->setRowCount(destinatioTable->rowCount()+1);
            for(int i = 0; i < (titles.size()+lastValues.size()); i++){
                QTableWidgetItem *item = new QTableWidgetItem();
                if(i >= titles.size()){
                    item->setText(lastValues.at(i-titles.size()));
                    item->setTextAlignment(Qt::AlignCenter);
                } else {
                    item->setText(currentTable->item(currentTable->currentRow(),i)->text());
                    item->setTextAlignment(Qt::AlignCenter);
                }
                destinatioTable->setItem(destinatioTable->rowCount()-1,i,item);
            }
            destinatioTable->resizeColumnsToContents();
            m_saved = false;
            break;
        }
    }
    updateStatusBar(tabWidget->currentIndex());
}

void MainWindow::addToOrder(const QString &tab, const QString &code, const QString &sheetName, const QStringList &values)
{
    QStringList orderValues = values;
    QTableWidget *table = qobject_cast<QTableWidget *>(tabWidget->currentWidget());
    orderValues.prepend(code);
    orderValues.prepend(table->item(table->currentRow(),3)->text()); //dispatch
    orderValues.prepend(table->item(table->currentRow(),1)->text()); //Father
    orderValues.prepend(table->item(table->currentRow(),0)->text()); //Name
    if(sheetName == "ترخیص"){
        orderValues.reserve(orderValues.size()+4);
        std::fill_n(std::back_inserter(orderValues), orderValues.size()+4, "");
        for(int h = 0; h < table->columnCount(); h++){
            if(table->horizontalHeaderItem(h)->text() == "تاریخ تولد"){
                orderValues[8] = table->item(table->currentRow(),h)->text();
            }
            if(table->horizontalHeaderItem(h)->text() == "محل تولد"){
                orderValues[7] = table->item(table->currentRow(),h)->text();
            }
            if(table->horizontalHeaderItem(h)->text() == "کد پرسنلی"){
                orderValues[6] = table->item(table->currentRow(),h)->text();
            }
        }
        QString kasri;
        QXlsx::Document doc(QDir::currentPath()+QDir::separator()+tab+QDir::separator()+code+".xlsx");
        doc.selectSheet("کسری");
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            kasri += DateCalc::dateDistanceInWords(DateCalc::daysToDate(doc.read(r,4).toString().toInt())) + " کسری " + doc.read(r,5).toString();
            if(r != doc.currentWorksheet()->dimension().rowCount()){
                kasri += " و ";
            }
        }
        doc.selectSheet("کسری منطقه عملیاتی");
        int dayDistance = 0;
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            QString dateDistance = DateCalc::minusTwoDate(Soldier::PerToEngNo(doc.read(r,4).toString()),Soldier::PerToEngNo(doc.read(r,5).toString()));
            dayDistance += DateCalc::dateToDays(dateDistance);
        }
        if(dayDistance > 0){
            if(!kasri.isEmpty()){
                kasri += " و ";
            }
            dayDistance = std::round(dayDistance/m_combosMap.value("multipliesOptions").at(5).toDouble());
            kasri += DateCalc::dateDistanceInWords(DateCalc::daysToDate(dayDistance)) + QString(" ") + doc.currentWorksheet()->sheetName();
        }
        orderValues[5] = kasri;
    }
    m_pendingOrder.insert(sheetName,orderValues);
}

void MainWindow::changeInOrder(const QString &sheetName, const QStringList &values, const int &newValues)
{
    QStringList orderValues = values;
    QTableWidget *table = qobject_cast<QTableWidget *>(tabWidget->currentWidget());
    orderValues.prepend(table->item(table->currentRow(),2)->text()); //Code
    orderValues.prepend(table->item(table->currentRow(),3)->text()); //dispatch
    orderValues.prepend(table->item(table->currentRow(),1)->text()); //Father
    orderValues.prepend(table->item(table->currentRow(),0)->text()); //Name
    orderValues.append(QString::number(newValues));
    m_pendingChangeOrder.insert(sheetName,orderValues);
}

void MainWindow::removeOrder(const QString &sheetName, const QStringList &values)
{
    QStringList orderValues = values;
    QTableWidget *table = qobject_cast<QTableWidget *>(tabWidget->currentWidget());
    orderValues.prepend(table->item(table->currentRow(),2)->text()); //Code
    orderValues.prepend(table->item(table->currentRow(),3)->text()); //dispatch
    orderValues.prepend(table->item(table->currentRow(),1)->text()); //Father
    orderValues.prepend(table->item(table->currentRow(),0)->text()); //Name
    m_pendingRemoveOrder.insert(sheetName,orderValues);
}

void MainWindow::summaryList(const QList<int> &list)
{
    summaryResult.clear();
    summaryResult = list;
}

void MainWindow::actionSearchTriggered()
{
    QString text = txtSearch->text();
    if(text.isEmpty()){
        for(int t = 0; t < tables.count(); t++){
            for(int r = 0; r < tables.at(t)->rowCount(); r++){
                tables.at(t)->setRowHidden(r,false);
            }
        }
        updateStatusBar(tabWidget->currentIndex());
        return;
    }
    for(int t = 0; t < tables.count(); t++){
        QSet<int> shown;
        auto items = tables.at(t)->findItems(text,Qt::MatchContains);
        for(int i = 0; i < items.size(); i++){
            tables.at(t)->setRowHidden(items.at(i)->row(),false);
            shown.insert(items.at(i)->row());
        }
        for(int r = 0; r < tables.at(t)->rowCount(); r++){
            if(!shown.contains(r)){
                tables.at(t)->setRowHidden(r,true);
            }
        }
    }
    updateStatusBar(tabWidget->currentIndex());
}

QMessageBox::StandardButton MainWindow::saveRequest()
{
    if(!m_saved || !m_pendingLetter.isEmpty() || !m_pendingMove.isEmpty() || !m_pendingRowColor.isEmpty() || !m_pendingRemoveLetter.isEmpty()){
        QString message = "در جداول تغییراتی اعمال کرده‌اید که ذخیره نشده است، آیا می‌خواهید تغییرات ذخیره شوند؟";
        QMessageBox::StandardButtons options = {QMessageBox::StandardButton::Save , QMessageBox::StandardButton::Discard , QMessageBox::StandardButton::Cancel};
        QMessageBox::StandardButton answer = QMessageBox::question(this, "File Not Saved",message, options);
        return answer;
    }
    return QMessageBox::StandardButton::Save;
}

void MainWindow::loadOptions()
{
    OptionDialog *optionDlg = new OptionDialog(this, true);
    QString optionsFilename = optionDlg->optionFilename();
    delete optionDlg;
    QFile optionsFile(optionsFilename);
    if(!optionsFile.exists()){
        QString title = "فایل یافت نشد";
        QMessageBox::information(this, title, "فایل تنظیمات یافت نشد، لطفا تنظیمات را اعمال کرده و ذخیره کنید");
        actionOptionTriggered();
    }
    if(!optionsFile.open(QIODevice::ReadOnly)){
        QMessageBox::critical(this,"Error",optionsFile.errorString());
        return;
    }
    QDataStream stream (&optionsFile);
    stream >> m_combosMap;
    optionsFile.close();
    m_dbFilename = m_combosMap.value("Files").at(1);
    m_ostanFilename = m_combosMap.value("Files").at(2);
    m_orderDir = m_combosMap.value("Files").at(3);
    m_statisticsDir = m_combosMap.value("Files").at(4);
    QStringList fontPeroperties = m_combosMap.value("Font",QStringList());
    QFont f;
    f.fromString(fontPeroperties.at(0));
    f.setPointSize(fontPeroperties.at(1).toInt());
    this->setFont(f);
    SoldierFiles::createOrder(m_orderDir,m_combosMap.value("Order").at(0).toInt(),this->font());
}

void MainWindow::updateRowColor(const int &tab, const int &row, const int &colorCode, const QString &description)
{
    QTableWidget *currentTable = qobject_cast<QTableWidget *>(tabWidget->widget(tab));
    for(int c = 0; c < currentTable->columnCount(); c++){
        QTableWidgetItem *item = currentTable->item(row,c);
        if(colorCode != -1){
            if(item->background() != QBrush(QColor(Qt::GlobalColor(colorCode)))){
                item->setBackground(QBrush(QColor(Qt::GlobalColor(colorCode))));
            }
        } else {
            if(item->background() != QBrush()){
                item->setBackground(QBrush());
            }
        }
        if(currentTable->horizontalHeaderItem(c)->text() == "توضیحات"){
            if(item->text() != description){
                item->setText(description);
                currentTable->resizeColumnToContents(c);
                m_saved = false;
            }
        }
    }
}

void MainWindow::closeEvent(QCloseEvent *event)
{
    QMessageBox::StandardButton answer = saveRequest();
    if(answer == QMessageBox::StandardButton::Save){
        actionSaveTriggered();
        event->setAccepted(true);
        return;
    } else if (answer == QMessageBox::StandardButton::Discard){
        event->setAccepted(true);
        return;
    } else {
        event->ignore();
        return;
    }
    event->setAccepted(true);
    return;
}
