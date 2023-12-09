#include "letterdialog.h"
#include "soldierfiles.h"
#include "ui_letterdialog.h"

LetterDialog::LetterDialog(const QString &tab, const QString &code, QWidget *parent) :
    QDialog(parent),
    ui(new Ui::LetterDialog)
{
    ui->setupUi(this);
    this->setWindowTitle("نامه های بایگانی");
    this->setWindowIcon(QIcon(":/icons/envelope.png"));
    this->setLocale(Soldier::persian());
    this->setLayoutDirection(Qt::LayoutDirection::RightToLeft);
    m_tab = tab;
    m_code = code;
    loadOption();

    //Get letters and code
    QStringList letters = Soldier::letters();
    QStringList codesList = Soldier::letters(true);

    //Setup Existing Letters widget
    txtSearch = new QLineEdit(this);
    txtSearch->setFixedWidth(180);
    txtSearch->setClearButtonEnabled(true);
    QAction *actionSearch = new QAction("جستجو", txtSearch);
    actionSearch->setIcon(QIcon(":/icons/search.png"));
    txtSearch->addAction(actionSearch, QLineEdit::LeadingPosition);
    txtSearch->setPlaceholderText("جستجو...");
    greenLabel = new QLabel(this);
    greenLabel->setText("توجه: نامه های با پس زمینه سبزرنگ به نامه های دیگر مرتبط اند");
    greenLabel->setFont(this->font());
    QFont greenLabelFont = greenLabel->font();
    greenLabelFont.setBold(true);
    greenLabel->setFont(greenLabelFont);
    QPalette greenLabelPalette = greenLabel->palette();
    greenLabel->setAutoFillBackground(true);
    greenLabelPalette.setColor(greenLabel->foregroundRole(), QColor(Qt::darkGreen));
    greenLabelPalette.setColor(greenLabel->backgroundRole(), QColor(Qt::yellow));
    greenLabel->setPalette(greenLabelPalette);
    btnRemove = new QPushButton(this);
    btnRemove->setText("حذف نامه انتخاب شده");
    tabWidget = new QTabWidget(this);
    tabWidget->setTabPosition(QTabWidget::South);
    tabWidget->setLayoutDirection(Qt::LayoutDirection::RightToLeft);
    QDir dir(QDir::currentPath()+QDir::separator()+tab);
    if (!dir.exists()){
        emit showLetterWarning("پوشه پرونده سربازان یافت نشد");
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
        emit showLetterWarning("فایل پرونده سرباز یافت نشد");
        return;
    }
    QXlsx::Document doc(QDir::currentPath()+QDir::separator()+tab+QDir::separator()+code+".xlsx");
    for(QString &sheet:doc.sheetNames()){
        doc.selectSheet(sheet);
        QTableWidget *table = new QTableWidget(tabWidget);
        table->setColumnCount(doc.currentWorksheet()->dimension().columnCount());
        table->setRowCount(doc.currentWorksheet()->dimension().rowCount()-1);
        QStringList header;
        for(int h = 1; h <= doc.currentWorksheet()->dimension().columnCount(); h++){
            header.append(doc.read(1,h).toString());
        }
        table->setHorizontalHeaderLabels(header);
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            for(int c = 1; c <= doc.currentWorksheet()->dimension().columnCount(); c++){
                QTableWidgetItem *item = new QTableWidgetItem(doc.read(r,c).toString());
                if(doc.currentWorksheet()->rowFormat(r).patternBackgroundColor().isValid()){
                    item->setBackground(QBrush(QColor(Qt::green)));
                }
                table->setItem(r-2,c-1,item);
                table->item(r-2,c-1)->setTextAlignment(Qt::AlignCenter);
            }
        }
        table->setEditTriggers(QAbstractItemView::NoEditTriggers);
        table->setSelectionMode(QAbstractItemView::SingleSelection);
        table->setSelectionBehavior(QAbstractItemView::SelectRows);
        table->resizeColumnsToContents();
        QScrollBar *customScrollBar = new QScrollBar(Qt::Horizontal,table);
        table->setHorizontalScrollBar(customScrollBar);
        customScrollBar->setLayoutDirection(Qt::LeftToRight);
        customScrollBar->setInvertedAppearance(true);
        tabWidget->addTab(table,sheet);
    }
    txtSearch->setHidden(true);
    greenLabel->setHidden(true);
    tabWidget->setHidden(true);
    btnRemove->setHidden(true);

    //Setup list
    listOptions = new QListWidget(this);
    listOptions->setFlow(QListWidget::TopToBottom);
    listOptions->addItems(letters);
    QScrollBar *customScrollBar = new QScrollBar(Qt::Horizontal,listOptions);
    listOptions->setHorizontalScrollBar(customScrollBar);
    customScrollBar->setLayoutDirection(Qt::LeftToRight);
    customScrollBar->setInvertedAppearance(true);
    stackedWidget = new QStackedWidget(this);
    for(int i = 0; i < letters.size(); i++){
        QHBoxLayout *grpTopLayout = new QHBoxLayout;
        QHBoxLayout *grpMidLayout = new QHBoxLayout;
        QHBoxLayout *grpMidLayout2 = new QHBoxLayout;
        bool isMidLayout2 = false;
        QHBoxLayout *grpBotLayout = new QHBoxLayout;
        QHBoxLayout *grpBotLayout2 = new QHBoxLayout;
        bool isBotLayout = false;
        QVBoxLayout *grpLayout = new QVBoxLayout;
        QGroupBox *grpBox = new QGroupBox(letters.at(i),this);
        grpBox->setObjectName("GroupBox"+QString::number(i));
        QStringList codes = codesList.at(i).split("");
        codes.removeAll("");
        int dateCount = 1;
        int comboCount = 1;
        int textCount = 1;
        for(const QString &code: codes){
            if(code == "A"){
                QLabel *ownerLabel = new QLabel(grpBox);
                ownerLabel->setText("صاحب نامه");
                QComboBox *owner = new QComboBox(grpBox);
                QStringList values1 = m_combosMap.value("قسمت",QStringList());
                QStringList values2 = m_combosMap.value("صاحب نامه غیر قسمتی",QStringList());
                owner->addItems(values1+values2);
                QLabel *letterNoLabel = new QLabel(grpBox);
                letterNoLabel->setText("شماره نامه");
                QLineEdit *letterNo = new QLineEdit(grpBox);
                QLabel *dateLabel = new QLabel(grpBox);
                dateLabel->setText("تاریخ نامه");
                QDateEdit *date = new QDateEdit(grpBox);
                date->setCalendar(Soldier::jalali());
                date->setLocale(Soldier::persian());
                date->setCalendarPopup(true);
                date->setDisplayFormat(Soldier::displayFormat());
                date->setDate(QDate::currentDate());
                date->setMinimumWidth(date->width());
                Soldier::setCalenderWidgetFormat(date->calendarWidget());
                grpTopLayout->addWidget(ownerLabel);
                grpTopLayout->addWidget(owner);
                grpTopLayout->addWidget(letterNoLabel);
                grpTopLayout->addWidget(letterNo);
                grpTopLayout->addWidget(dateLabel);
                grpTopLayout->addWidget(date);
            } else if(code == "N") {
                QLabel *salianeLabel = new QLabel(grpBox);
                salianeLabel->setText("سالیانه");
                QSpinBox *saliane = new QSpinBox(grpBox);
                saliane->setMaximum(Soldier::maxSpinBox());

                QLabel *tashvighiLabel = new QLabel(grpBox);
                tashvighiLabel->setText("تشویقی");
                QSpinBox *tashvighi = new QSpinBox(grpBox);
                tashvighi->setMaximum(Soldier::maxSpinBox());

                QLabel *toRahiLabel = new QLabel(grpBox);
                toRahiLabel->setText("توراهی");
                QSpinBox *toRahi = new QSpinBox(grpBox);
                toRahi->setMaximum(Soldier::maxSpinBox());

                QLabel *tatiliLabel = new QLabel(grpBox);
                tatiliLabel->setText("تعطیلی");
                QSpinBox *tatili = new QSpinBox(grpBox);
                tatili->setMaximum(Soldier::maxSpinBox());

                QLabel *estelajiLabel = new QLabel(grpBox);
                estelajiLabel->setText("استراحت پزشکی");
                QSpinBox *estelaji = new QSpinBox(grpBox);
                estelaji->setMaximum(Soldier::maxSpinBox());

                grpBotLayout->addWidget(salianeLabel);
                grpBotLayout->addWidget(saliane);
                grpBotLayout->addWidget(tashvighiLabel);
                grpBotLayout->addWidget(tashvighi);
                grpBotLayout->addWidget(toRahiLabel);
                grpBotLayout->addWidget(toRahi);
                grpBotLayout->addWidget(tatiliLabel);
                grpBotLayout->addWidget(tatili);
                grpBotLayout->addWidget(estelajiLabel);
                grpBotLayout->addWidget(estelaji);
                isBotLayout = true;
            } else if(code == "d") {
                QLabel *dateLabel = new QLabel(grpBox);
                QDateEdit *date = new QDateEdit(grpBox);
                date->setCalendar(Soldier::jalali());
                date->setLocale(Soldier::persian());
                date->setCalendarPopup(true);
                date->setDisplayFormat(Soldier::displayFormat());
                date->setDate(QDate::currentDate());
                date->setMinimumWidth(date->width());
                Soldier::setCalenderWidgetFormat(date->calendarWidget());
                if (dateCount == 1){
                    if(!codesList.at(i).contains("dd")){
                        dateLabel->setText("تاریخ " + letters.at(i));
                    } else {
                        dateLabel->setText("تاریخ شروع");
                        dateCount++;
                    }
                } else {
                    dateLabel->setText("تاریخ پایان");
                    date->setObjectName("Date"+QString::number(i));
                }
                grpMidLayout->addWidget(dateLabel);
                grpMidLayout->addWidget(date);
            } else if(code == "c"){
                QLabel *comboLabel = new QLabel(grpBox);
                QComboBox *combo = new QComboBox(grpBox);
                if(comboCount == 1){
                    if(!codesList.at(i).contains("cc")){
                        comboLabel->setText(letters.at(i));
                        QStringList values = m_combosMap.value(letters.at(i),QStringList());
                        combo->addItems(values);
                    } else {
                        if (letters.at(i).contains("کسر و اضافه به آمار")){
                            isMidLayout2 = true;
                            comboLabel->setText("قسمت کسر کننده");
                            QStringList values = m_combosMap.value("قسمت",QStringList());
                            combo->addItems(values);
                        } else if (letters.at(i).contains("شورا پزشکی")) {
                            comboLabel->setText("وضعیت جسمانی");
                            QStringList values = m_combosMap.value("وضعیت جسمانی",QStringList());
                            combo->addItems(values);
                        }
                        comboCount++;
                    }
                } else {
                    if (letters.at(i).contains("کسر و اضافه به آمار")){
                        isMidLayout2 = true;
                        comboLabel->setText("قسمت اضافه شونده");
                        QStringList values = m_combosMap.value("قسمت",QStringList());
                        combo->addItems(values);
                    } else if (letters.at(i).contains("شورا پزشکی")) {
                        comboLabel->setText("وضعیت روانی");
                        QStringList values = m_combosMap.value("وضعیت روانی",QStringList());
                        combo->addItems(values);
                    }
                }
                if(isMidLayout2){
                    grpMidLayout2->addWidget(comboLabel);
                    grpMidLayout2->addWidget(combo);
                } else {
                    grpMidLayout->addWidget(comboLabel);
                    grpMidLayout->addWidget(combo);
                }
            } else if(code == "n"){
                QLabel *numberLabel = new QLabel(grpBox);
                numberLabel->setText("تعداد روز");
                QSpinBox *number = new QSpinBox(grpBox);
                number->setMaximum(1000);
                grpMidLayout->addWidget(numberLabel);
                grpMidLayout->addWidget(number);
            } else if(code == "t"){
                QLabel *textLabel = new QLabel(grpBox);
                if(!codesList.at(i).contains("tt")){
                    textLabel->setText("علت");
                } else {
                    isMidLayout2 = true;
                    if(textCount == 1){
                        textLabel->setText("محل");
                        textCount++;
                    } else {
                        textLabel->setText("شماره امریه");
                    }
                }
                QLineEdit *text = new QLineEdit(grpBox);
                if(isMidLayout2){
                    grpMidLayout2->addWidget(textLabel);
                    grpMidLayout2->addWidget(text);
                } else {
                    grpMidLayout->addWidget(textLabel);
                    grpMidLayout->addWidget(text);
                }
            } else if(code == "o"){
                QCheckBox *oldLetter = new QCheckBox(grpBox);
                oldLetter->setText("نامه قدیمی");
                oldLetter->setObjectName("oldLetter");
                grpBotLayout2->addWidget(oldLetter);
            }
        }
        grpTopLayout->addStretch();
        grpMidLayout->addStretch();
        grpMidLayout2->addStretch();
        grpBotLayout->addStretch();
        grpBotLayout2->addStretch();
        grpLayout->addLayout(grpTopLayout);
        grpLayout->addLayout(grpMidLayout);
        if(isMidLayout2){
            grpLayout->addLayout(grpMidLayout2);
        }
        if(isBotLayout){
            grpLayout->addLayout(grpBotLayout);
        }
        grpLayout->addLayout(grpBotLayout2);
        grpLayout->addStretch();
        grpBox->setLayout(grpLayout);
        stackedWidget->addWidget(grpBox);
    }

    //Setup ButtonBox
    QPushButton *btnOld = new QPushButton(this);
    QPushButton *btnAdd = new QPushButton(this);
    QPushButton *btnCancel = new QPushButton(this);
    btnAdd->setText("Add");
    btnOld->setText("مشاهده نامه های ثبت شده قبلی");
    btnCancel->setText("Cancel");
    QHBoxLayout *btnBoxLayout = new QHBoxLayout;
    btnBoxLayout->addWidget(btnOld);
    btnBoxLayout->addWidget(btnRemove);
    btnBoxLayout->addStretch();
    btnBoxLayout->addWidget(btnAdd);
    btnBoxLayout->addWidget(btnCancel);

    //Setup Splitters and Layout
    QSplitter *hsplit = new QSplitter(Qt::Horizontal,this);
    hsplit->addWidget(listOptions);
    hsplit->addWidget(stackedWidget);
    QHBoxLayout *hbox = new QHBoxLayout;
    hbox->addWidget(txtSearch);
    hbox->addWidget(greenLabel);
    QVBoxLayout *vbox = new QVBoxLayout;
    vbox->addWidget(hsplit);
    vbox->addLayout(btnBoxLayout);
    vbox->addLayout(hbox);
    vbox->addWidget(tabWidget);
    this->setLayout(vbox);


    //connect signal and slots
    connect(listOptions,&QListWidget::currentRowChanged,stackedWidget,&QStackedWidget::setCurrentIndex);
    connect(btnAdd,&QPushButton::clicked,this,&LetterDialog::btnAddClicked);
    connect(btnRemove,&QPushButton::clicked,this,&LetterDialog::btnRemoveClicked);
    connect(btnOld,&QPushButton::clicked,this,&LetterDialog::btnOldClicked);
    connect(btnCancel,&QPushButton::clicked,this,&QDialog::reject);
    connect(actionSearch,&QAction::triggered,this,&LetterDialog::actionSearchTriggered);
    connect(txtSearch,&QLineEdit::returnPressed,this,&LetterDialog::actionSearchTriggered);
    listOptions->setCurrentRow(0);
}

LetterDialog::~LetterDialog()
{
    delete ui;
}

void LetterDialog::btnAddClicked()
{
    QString letterType = listOptions->currentItem()->text();
    QStringList values;
    for(QWidget *widget: stackedWidget->currentWidget()->findChildren<QWidget*>(Qt::FindDirectChildrenOnly)){
        if(QLineEdit* newWidget = qobject_cast<QLineEdit*>(widget)){
            values.append(newWidget->text());
        }
        if(QComboBox* newWidget = qobject_cast<QComboBox*>(widget)){
            values.append(newWidget->currentText());
        }
        if(QSpinBox* newWidget = qobject_cast<QSpinBox*>(widget)){
            values.append(QString::number(newWidget->value()));
        }
        if(QDateEdit* newWidget = qobject_cast<QDateEdit*>(widget)){
            values.append(newWidget->text());
        }
        if(QCheckBox* newWidget = qobject_cast<QCheckBox*>(widget)){
            if(newWidget->isChecked()){
                values.append(newWidget->text());
            } else {
                values.append("");
            }
        }
    }
    SoldierFiles *soldierFiles = new SoldierFiles(this);
    connect(soldierFiles,&SoldierFiles::changeColumnData,this,&LetterDialog::sendChangeColumnData);
    connect(soldierFiles,&SoldierFiles::changeRowColor,this,&LetterDialog::sendChangeRowColor);
    connect(soldierFiles,&SoldierFiles::showLetterWarning,this,&LetterDialog::showLetterWarning);
    connect(soldierFiles,&SoldierFiles::moveToTable,this,&LetterDialog::sendMoveToTable);
    connect(soldierFiles,&SoldierFiles::addToTable,this,&LetterDialog::sendAddToTable);
    if (soldierFiles->validateLetter(m_tab,m_code,letterType,values)){
        emit addLetter(m_tab,m_code,letterType,values);
        accept();
    }
}

void LetterDialog::btnRemoveClicked()
{
    QString letterType = tabWidget->tabText(tabWidget->currentIndex());
    QString title = "حذف نامه";
    QString message;
    QTableWidget *table = qobject_cast<QTableWidget *>(tabWidget->currentWidget());
    if(table->item(table->currentRow(),0) == nullptr){
        message = "ابتدا نامه ای را برای حذف کردن انتخاب کنید";
        QMessageBox::warning(this,title,message);
        return;
    }
    message = "آیا مطمئن هستید که قصد دارید نامه زیر را حذف کنید؟";
    message += "\n";
    message += "نوع نامه: ";
    message += letterType + "\n";
    for(int c = 0; c < table->columnCount(); c++){
        message += table->horizontalHeaderItem(c)->text() + ": " + table->item(table->currentRow(),c)->text() + "\n";
    }
    QMessageBox::StandardButtons btns = (QMessageBox::Yes | QMessageBox::No);
    QMessageBox *msg = new QMessageBox(QMessageBox::Question, title,message,btns,this);
    msg->setWindowIcon(this->windowIcon());
    msg->setFont(this->font());
    msg->setTextInteractionFlags(Qt::TextSelectableByMouse|Qt::TextSelectableByKeyboard);
    int answer = msg->exec();
    delete msg;
    if(answer == QMessageBox::No){
        return;
    }
    if(answer == QMessageBox::Rejected){
        return;
    }
    SoldierFiles *soldierFiles = new SoldierFiles(this);
    connect(soldierFiles,&SoldierFiles::changeColumnData,this,&LetterDialog::sendChangeColumnData);
    connect(soldierFiles,&SoldierFiles::changeRowColor,this,&LetterDialog::sendChangeRowColor);
    connect(soldierFiles,&SoldierFiles::showLetterWarning,this,&LetterDialog::showLetterWarning);
    connect(soldierFiles,&SoldierFiles::moveToTable,this,&LetterDialog::sendMoveToTable);
    if(soldierFiles->validateRemove(m_tab,m_code,letterType,table->currentRow()+2)){
        emit removeLetter(m_tab,m_code,tabWidget->currentIndex(),table->currentRow()+2);
        accept();
    }
}

void LetterDialog::btnOldClicked()
{
    QPushButton *btnWidget = qobject_cast<QPushButton*>(sender());
    if(tabWidget->isHidden()){
        txtSearch->setHidden(false);
        greenLabel->setHidden(false);
        btnRemove->setHidden(false);
        tabWidget->setHidden(false);
        btnWidget->setText("مخفی کردن نامه های ثبت شده قبلی");
    } else {
        txtSearch->setHidden(true);
        greenLabel->setHidden(true);
        btnRemove->setHidden(true);
        tabWidget->setHidden(true);
        btnWidget->setText("مشاهده نامه های ثبت شده قبلی");
    }
}

void LetterDialog::actionSearchTriggered()
{
    QString text = txtSearch->text();
    if(text.isEmpty()){
        for(int t = 0; t < tabWidget->count(); t++){
            QTableWidget *table = qobject_cast<QTableWidget*>(tabWidget->widget(t));
            for(int r = 0; r < table->rowCount(); r++){
                table->setRowHidden(r,false);
            }
        }
        return;
    }
    for(int t = 0; t < tabWidget->count(); t++){
        QTableWidget *table = qobject_cast<QTableWidget*>(tabWidget->widget(t));
        QSet<int> shown;
        for(QTableWidgetItem *item:table->findItems(text,Qt::MatchContains)){
            table->setRowHidden(item->row(),false);
            shown.insert(item->row());
        }
        for(int r = 0; r < table->rowCount(); r++){
            if(!shown.contains(r)){
                table->setRowHidden(r,true);
            }
        }
    }
}

void LetterDialog::sendChangeColumnData(const int &columnIndex, const QString &columnNewData)
{
    emit changeColumnData(columnIndex, columnNewData);
}

void LetterDialog::sendChangeRowColor(const QString &tab, const int &rowIndex, const int &colorCode)
{
    emit changeRowColor(tab, rowIndex, colorCode);
}

void LetterDialog::showLetterWarning(const QString &message)
{
    QMessageBox::warning(this,"Failed",message);
}

void LetterDialog::sendMoveToTable(const QString &tab, const QString &code, const QString &newTab, const QStringList &lastValues)
{
    emit moveToTable(tab, code, newTab, lastValues);
}

void LetterDialog::sendAddToTable(const QString &tab, const QString &code, const QString &newTab, const QStringList &lastValues)
{
    emit addToTable(tab, code, newTab, lastValues);
}

void LetterDialog::sendAddToOrder(const QString &tab, const QString &code, const QString &sheetName, const QStringList &values)
{
    emit addToOrder(tab, code, sheetName, values);
}

void LetterDialog::loadOption()
{
    OptionDialog* optDlg = new OptionDialog(this,true);
    QFile file(optDlg->optionFilename());
    delete optDlg;
    if(!file.exists()){
        return;
    }
    if(!file.open(QIODevice::ReadOnly)){
        QMessageBox::critical(this,"Error Rading Option File",file.errorString());
        return;
    }
    QDataStream stream(&file);
    stream >> m_combosMap;
    file.close();
    QStringList fontPeroperties = m_combosMap.value("Font",QStringList());
    if(!fontPeroperties.isEmpty()){
        QFont f;
        f.fromString(fontPeroperties.at(0));
        f.setPointSize(fontPeroperties.at(1).toInt());
        this->setFont(f);
    }
    return;
}


void LetterDialog::closeEvent(QCloseEvent *event)
{
    event->accept();
}
