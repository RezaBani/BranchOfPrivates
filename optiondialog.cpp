#include "optiondialog.h"
#include "soldier.h"
#include "ui_optiondialog.h"

OptionDialog::OptionDialog(QWidget *parent, const bool &onlyGetFilename) :
    QDialog(parent),
    ui(new Ui::OptionDialog)
{
    ui->setupUi(this);
    m_optionFilename = "options.dat";
    m_dbFilename = "database.xlsx";
    m_ostanFilename = "ostanList.xlsx";
    m_orderDir = "دستور";
    m_statisticsDir = "آمار پایه خدمتی";
    if(onlyGetFilename){
        return;
    }
    this->setWindowTitle("تنظیمات");
    this->setWindowIcon(QIcon(":/icons/settings.png"));
    this->setLayoutDirection(Qt::LayoutDirection::RightToLeft);
    loadOptions();

    //Setup Buttons
    QDialogButtonBox *btnBox = new QDialogButtonBox(this);
    btnBox->addButton(QDialogButtonBox::Cancel);
    btnBox->addButton(QDialogButtonBox::SaveAll);

    //Setup Options List
    QListWidget *optionList = new QListWidget(this);
    optionList->setFlow(QListWidget::TopToBottom);
    optionList->addItem("عمومی");
    optionList->addItem("متغیرها");
    optionList->addItem("ضریب ها");
    optionList->addItem("لیست ها");
    optionList->addItem("استان ها");

    //Setup Stacked
    stackedWidget = new QStackedWidget(this);
    stackedWidget->setObjectName("stackedWidget");

    //Setup General Options
    QGroupBox *generalOptions = new QGroupBox(stackedWidget);
    generalOptions->setTitle(optionList->item(0)->text());
    QLabel *labelFontCombo = new QLabel("انتخاب فونت",generalOptions);
    fontComboBox = new QFontComboBox(generalOptions);
    QLabel *labelFontSize = new QLabel("سایز فونت",generalOptions);
    fontSize = new QSpinBox(generalOptions);
    fontSize->setMinimum(9);
    fontSize->setMaximum(16);
    QStringList fontPeroperties = m_combosMap.value("Font",QStringList());
    if(!fontPeroperties.isEmpty()){
        QFont f;
        f.fromString(fontPeroperties.at(0));
        f.setPointSize(fontPeroperties.at(1).toInt());
        this->setFont(f);
        fontComboBox->setCurrentFont(f);
        fontSize->setValue(f.pointSize());
    }
    QLabel *labelOrderNo = new QLabel("شماره دستور فعلی",generalOptions);
    orderNo = new QSpinBox(generalOptions);
    orderNo->setMinimum(1);
    orderNo->setMaximum(Soldier::maxSpinBox());
    orderNo->setValue(m_combosMap.value("Order",QStringList(QString::number(1))).at(0).toInt());
    //Setup General Options Layout
    QVBoxLayout *generalVBox = new QVBoxLayout;
    QGridLayout *generalGrid = new QGridLayout;
    generalGrid->addWidget(labelFontCombo,0,0);
    generalGrid->addWidget(fontComboBox,0,1);
    generalGrid->addWidget(labelFontSize,1,0);
    generalGrid->addWidget(fontSize,1,1);
    generalGrid->addWidget(labelOrderNo,2,0);
    generalGrid->addWidget(orderNo,2,1);
    generalVBox->addLayout(generalGrid);
    generalVBox->addStretch();
    generalOptions->setLayout(generalVBox);

    //Setup Variables Options
    QGroupBox *variablesOptions = new QGroupBox(stackedWidget);
    variablesOptions->setTitle(optionList->item(1)->text());
    variablesOptions->setObjectName("variablesOptions");
    QStringList defaultVariables;
    for(int i = 0; i < 7; i++){
        defaultVariables.append("0");
    }

    QLabel *salaryPerMounthLabel = new QLabel(variablesOptions);
    salaryPerMounthLabel->setText("سقف سالیانه به ازای هر ماه");
    QDoubleSpinBox *salaryPerMounth = new QDoubleSpinBox(variablesOptions);
    salaryPerMounth->setObjectName("salaryPerMounth");
    salaryPerMounth->setValue(m_combosMap.value("variablesOptions",defaultVariables).at(0).toDouble());

    QLabel *estelajiPerMounthLabel = new QLabel(variablesOptions);
    estelajiPerMounthLabel->setText("سقف استعلاجی به ازای هر ماه");
    QDoubleSpinBox *estelajiPerMounth = new QDoubleSpinBox(variablesOptions);
    estelajiPerMounth->setObjectName("estelajiPerMounth");
    estelajiPerMounth->setValue(m_combosMap.value("variablesOptions",defaultVariables).at(1).toDouble());

    QLabel *maxTashvighiLabel = new QLabel(variablesOptions);
    maxTashvighiLabel->setText("سقف تشویقی قابل استفاده");
    QSpinBox *maxTashvighi = new QSpinBox(variablesOptions);
    maxTashvighi->setObjectName("maxTashvighi");
    maxTashvighi->setValue(m_combosMap.value("variablesOptions",defaultVariables).at(2).toInt());


    QLabel *maxSingleToRahiLabel = new QLabel(variablesOptions);
    maxSingleToRahiLabel->setText("سقف تعداد توراهی برای مجرد");
    QSpinBox *maxSingleToRahi = new QSpinBox(variablesOptions);
    maxSingleToRahi->setObjectName("maxSingleToRahi");
    maxSingleToRahi->setValue(m_combosMap.value("variablesOptions",defaultVariables).at(3).toInt());

    QLabel *maxMarriedToRahiLabel = new QLabel(variablesOptions);
    maxMarriedToRahiLabel->setText("سقف تعداد توراهی برای متاهل");
    QSpinBox *maxMarriedToRahi = new QSpinBox(variablesOptions);
    maxMarriedToRahi->setObjectName("maxMarriedToRahi");
    maxMarriedToRahi->setValue(m_combosMap.value("variablesOptions",defaultVariables).at(4).toInt());

    QLabel *gheireBomiGhanoniLabel = new QLabel(variablesOptions);
    gheireBomiGhanoniLabel->setText("خدمت قانونی غیربومی");
    QSpinBox *gheireBomiGhanoni = new QSpinBox(variablesOptions);
    gheireBomiGhanoni->setObjectName("gheireBomiGhanoni");
    gheireBomiGhanoni->setValue(m_combosMap.value("variablesOptions",defaultVariables).at(5).toInt());

    QLabel *bomiGhanoniLabel = new QLabel(variablesOptions);
    bomiGhanoniLabel->setText("خدمت قانونی بومی");
    QSpinBox *bomiGhanoni = new QSpinBox(variablesOptions);
    bomiGhanoni->setObjectName("bomiGhanoni");
    bomiGhanoni->setValue(m_combosMap.value("variablesOptions",defaultVariables).at(6).toInt());

    //Setup Multiplies Options Layout
    QVBoxLayout *variablesVBox = new QVBoxLayout;
    QGridLayout *variablesGrid = new QGridLayout;
    variablesGrid->addWidget(salaryPerMounthLabel,0,0);
    variablesGrid->addWidget(salaryPerMounth,0,1);
    variablesGrid->addWidget(estelajiPerMounthLabel,1,0);
    variablesGrid->addWidget(estelajiPerMounth,1,1);
    variablesGrid->addWidget(maxTashvighiLabel,2,0);
    variablesGrid->addWidget(maxTashvighi,2,1);
    variablesGrid->addWidget(maxSingleToRahiLabel,3,0);
    variablesGrid->addWidget(maxSingleToRahi,3,1);
    variablesGrid->addWidget(maxMarriedToRahiLabel,4,0);
    variablesGrid->addWidget(maxMarriedToRahi,4,1);
    variablesGrid->addWidget(gheireBomiGhanoniLabel,5,0);
    variablesGrid->addWidget(gheireBomiGhanoni,5,1);
    variablesGrid->addWidget(bomiGhanoniLabel,6,0);
    variablesGrid->addWidget(bomiGhanoni,6,1);
    variablesVBox->addLayout(variablesGrid);
    variablesVBox->addStretch();
    variablesOptions->setLayout(variablesVBox);

    //Setup Multiplies Options
    QGroupBox *multipliesOptions = new QGroupBox(stackedWidget);
    multipliesOptions->setTitle(optionList->item(2)->text());
    multipliesOptions->setObjectName("multipliesOptions");
    QStringList defaultMultiplies;
    for(int i = 0; i < 6; i++){
        defaultMultiplies.append("0");
    }

    QLabel *oldNahastLabel = new QLabel(multipliesOptions);
    oldNahastLabel->setText("ضریب نهست قدیمی (تا قبل از 1395/06/18)");
    QSpinBox *oldNahast = new QSpinBox(multipliesOptions);
    oldNahast->setValue(m_combosMap.value("multipliesOptions",defaultMultiplies).at(0).toInt());

    QLabel *nahastLabel = new QLabel(multipliesOptions);
    nahastLabel->setText("ضریب نهست");
    QSpinBox *nahast = new QSpinBox(multipliesOptions);
    nahast->setValue(m_combosMap.value("multipliesOptions",defaultMultiplies).at(1).toInt());

    QLabel *eidNahastLabel = new QLabel(multipliesOptions);
    eidNahastLabel->setText("ضریب نهست عید (از ۱ تا ۱۳ فروردین)");
    QSpinBox *eidNahast = new QSpinBox(multipliesOptions);
    eidNahast->setValue(m_combosMap.value("multipliesOptions",defaultMultiplies).at(2).toInt());

    QLabel *baKhedmatLabel = new QLabel(multipliesOptions);
    baKhedmatLabel->setText("ضریب بازداشت با خدمت");
    QSpinBox *baKhedmat = new QSpinBox(multipliesOptions);
    baKhedmat->setValue(m_combosMap.value("multipliesOptions",defaultMultiplies).at(3).toInt());

    QLabel *bedoneKhedmatLabel = new QLabel(multipliesOptions);
    bedoneKhedmatLabel->setText("ضریب بازداشت بدون خدمت");
    QSpinBox *bedoneKhedmat = new QSpinBox(multipliesOptions);
    bedoneKhedmat->setValue(m_combosMap.value("multipliesOptions",defaultMultiplies).at(4).toInt());

    QLabel *kasriAmaliatiLabel = new QLabel(multipliesOptions);
    kasriAmaliatiLabel->setText("ضریب کسری منطقه عملیاتی");
    QSpinBox *kasriAmaliati = new QSpinBox(multipliesOptions);
    kasriAmaliati->setValue(m_combosMap.value("multipliesOptions",defaultMultiplies).at(5).toInt());

    //Setup Multiplies Options Layout
    QVBoxLayout *multipliesVBox = new QVBoxLayout;
    QGridLayout *multipliesGrid = new QGridLayout;
    multipliesGrid->addWidget(oldNahastLabel,0,0);
    multipliesGrid->addWidget(oldNahast,0,1);
    multipliesGrid->addWidget(nahastLabel,1,0);
    multipliesGrid->addWidget(nahast,1,1);
    multipliesGrid->addWidget(eidNahastLabel,2,0);
    multipliesGrid->addWidget(eidNahast,2,1);
    multipliesGrid->addWidget(baKhedmatLabel,3,0);
    multipliesGrid->addWidget(baKhedmat,3,1);
    multipliesGrid->addWidget(bedoneKhedmatLabel,4,0);
    multipliesGrid->addWidget(bedoneKhedmat,4,1);
    multipliesGrid->addWidget(kasriAmaliatiLabel,5,0);
    multipliesGrid->addWidget(kasriAmaliati,5,1);
    multipliesVBox->addLayout(multipliesGrid);
    multipliesVBox->addStretch();
    multipliesOptions->setLayout(multipliesVBox);


    //Setup Lists Options
    QGroupBox* listsOptions = new QGroupBox(stackedWidget);
    listsOptions->setTitle(optionList->item(3)->text());
    QLabel *labelSelectList = new QLabel("انتخاب لیست:", listsOptions);
    QLabel *labelSearchText = new QLabel("جستجوی لیست:", listsOptions);
    textSearch = new QLineEdit(listsOptions);
    QPushButton *btnAdd = new QPushButton(listsOptions);
    btnAdd->setIcon(QIcon(":/icons/add.png"));
    btnAdd->setToolTip("افزودن به لیست");
    btnAdd->setDefault(true);
    QPushButton *btnRemove = new QPushButton(listsOptions);
    btnRemove->setIcon(QIcon(":/icons/minus.png"));
    btnRemove->setToolTip("حذف از لیست");
    QStringList labels = Soldier::titles();
    QStringList codes = Soldier::titles(true);
    labels.append("استخدام");
    labels.append("انتقالی");
    labels.append("کسری");
    labels.append("کسری منطقه عملیاتی");
    labels.append("صاحب نامه غیر قسمتی");
    labels.append("یگان اصلی مامور");
    codes.append("c");
    codes.append("c");
    codes.append("c");
    codes.append("c");
    codes.append("c");
    codes.append("c");
    listComboBox = new QComboBox(listsOptions);
    stackedLists = new QStackedWidget(listsOptions);
    int j = 0;
    for(int i = 0; i < labels.count(); i++){
        if(codes.at(i) == "c"){
            listComboBox->addItem(labels.at(i));
            QListWidget *listWidget = new QListWidget(stackedLists);
            listWidget->setObjectName(QString("list%0").arg(QString::number(j)));
            listWidget->addItems(m_combosMap.value(labels.at(i),QStringList()));
            connect(listWidget,&QListWidget::currentTextChanged,textSearch,&QLineEdit::setText);
            stackedLists->addWidget(listWidget);
            j++;
        }
    }
    //Setup lists Options Layout
    QGridLayout *listsOptionsGrid = new QGridLayout;
    listsOptionsGrid->addWidget(labelSelectList,0,0);
    listsOptionsGrid->addWidget(listComboBox,0,1,1,3);
    listsOptionsGrid->addWidget(labelSearchText,1,0);
    listsOptionsGrid->addWidget(textSearch,1,1);
    listsOptionsGrid->addWidget(btnAdd,1,2);
    listsOptionsGrid->addWidget(btnRemove,1,3);
    QVBoxLayout *listsOptionsVBox = new QVBoxLayout;
    listsOptionsVBox->addLayout(listsOptionsGrid);
    listsOptionsVBox->addWidget(stackedLists);
    listsOptions->setLayout(listsOptionsVBox);

    //Setup Ostan Options
    QGroupBox* ostansOptions = new QGroupBox(stackedWidget);
    ostansOptions->setTitle(optionList->item(4)->text());
    QLabel *ostanLabel = new QLabel(ostansOptions);
    QString ostanString = "برای تغییر مقادیر بر روی آیتم مورد نظر دوبار کلیک کنید";
    ostanString += "\n";
    ostanString += "مقادیر تغییر یافته با پس زمینه زرد رنگ مشخص می‌شوند";
    ostanLabel->setText(ostanString);
    ostanTable = new QTableWidget(m_combosMap.value("استان").size(),2,ostansOptions);
    QXlsx::Document doc(m_ostanFilename);
    doc.selectSheet(0);
    ostanTable->setHorizontalHeaderLabels(QStringList{doc.read(1,1).toString(),doc.read(1,2).toString()});
    for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
        QTableWidgetItem *item1 = new QTableWidgetItem;
        item1->setText(doc.read(r,1).toString());
        item1->setTextAlignment(Qt::AlignCenter);
        ostanTable->setItem(r-2,0,item1);
        QTableWidgetItem *item2 = new QTableWidgetItem;
        item2->setText(doc.read(r,2).toString());
        item2->setTextAlignment(Qt::AlignCenter);
        ostanTable->setItem(r-2,1,item2);
    }
    ostanTable->resizeColumnsToContents();
    ostanTable->setSelectionMode(QAbstractItemView::SingleSelection);
    ostanTable->setEditTriggers(QAbstractItemView::NoEditTriggers);
    ostanTable->setAlternatingRowColors(true);
    QScrollBar *customScrollBar = new QScrollBar(Qt::Horizontal,ostanTable);
    ostanTable->setHorizontalScrollBar(customScrollBar);
    customScrollBar->setLayoutDirection(Qt::LeftToRight);
    customScrollBar->setInvertedAppearance(true);
    //Setup Ostan Options Layout
    QVBoxLayout *ostansOptionsVBox = new QVBoxLayout;
    ostansOptionsVBox->addWidget(ostanLabel);
    ostansOptionsVBox->addWidget(ostanTable);
    ostansOptions->setLayout(ostansOptionsVBox);

    //Add Widgets to stack
    stackedWidget->addWidget(generalOptions);
    stackedWidget->addWidget(variablesOptions);
    stackedWidget->addWidget(multipliesOptions);
    stackedWidget->addWidget(listsOptions);
    stackedWidget->addWidget(ostansOptions);

    //Setup Layout
    QSplitter *hsplit = new QSplitter(Qt::Horizontal,this);
    hsplit->addWidget(optionList);
    hsplit->addWidget(stackedWidget);
    QVBoxLayout *vbox = new QVBoxLayout;
    vbox->addWidget(hsplit);
    vbox->addWidget(btnBox);
    this->setLayout(vbox);

    //Connecting Signal and Slots
    connect(btnBox,&QDialogButtonBox::rejected,this,&QDialog::close);
    connect(btnBox,&QDialogButtonBox::accepted,this,&OptionDialog::saveOptions);
    connect(optionList,&QListWidget::currentRowChanged,stackedWidget,&QStackedWidget::setCurrentIndex);
    connect(listComboBox,QOverload<int>::of(&QComboBox::currentIndexChanged),stackedLists,&QStackedWidget::setCurrentIndex);
    connect(btnAdd,&QPushButton::clicked,this,&OptionDialog::addToList);
    connect(btnRemove,&QPushButton::clicked,this,&OptionDialog::removeFromList);
    connect(ostanTable,&QTableWidget::cellDoubleClicked,this,&OptionDialog::cellDoubleClicked);
    connect(fontComboBox,&QFontComboBox::currentFontChanged,this,&OptionDialog::currentFontChanged);
    connect(fontSize,QOverload<int>::of(&QSpinBox::valueChanged),this,&OptionDialog::ivalueChanged);
    for(QSpinBox *spin:multipliesOptions->findChildren<QSpinBox*>(QString(), Qt::FindDirectChildrenOnly)){
        connect(spin,QOverload<int>::of(&QSpinBox::valueChanged),this,&OptionDialog::ivalueChanged);
    }
    for(QSpinBox *spin:variablesOptions->findChildren<QSpinBox*>(QString(), Qt::FindDirectChildrenOnly)){
        connect(spin,QOverload<int>::of(&QSpinBox::valueChanged),this,&OptionDialog::ivalueChanged);
    }
    for(QDoubleSpinBox *spin:variablesOptions->findChildren<QDoubleSpinBox*>(QString(), Qt::FindDirectChildrenOnly)){
        connect(spin,QOverload<double>::of(&QDoubleSpinBox::valueChanged),this,&OptionDialog::dvalueChanged);
    }
}

OptionDialog::~OptionDialog()
{
    delete ui;
}

QString OptionDialog::optionFilename() const
{
    return m_optionFilename;
}

QString OptionDialog::dbFilename() const
{
    return m_dbFilename;
}

QString OptionDialog::ostanFilename() const
{
    return m_ostanFilename;
}

QMap<QString, QStringList> OptionDialog::combosMap() const
{
    return m_combosMap;
}

void OptionDialog::saveOptions()
{
    QFile file(m_optionFilename);
    if(!file.open(QIODevice::WriteOnly)){
        QMessageBox::critical(this,"Error",file.errorString());
        return;
    }
    m_combosMap.insert("Order",QStringList(QString::number(orderNo->value())));
    QStringList files;
    files.append(m_optionFilename);
    files.append(m_dbFilename);
    files.append(m_ostanFilename);
    files.append(m_orderDir);
    files.append(m_statisticsDir);
    m_combosMap.insert("Files",files);
    QDataStream stream(&file);
    QStringList fontPeroperties;
    fontPeroperties.append(fontComboBox->currentFont().toString());
    fontPeroperties.append(QString::number(fontSize->value()));
    m_combosMap.insert("Font",fontPeroperties);
    QGroupBox *variablesOptions = stackedWidget->findChildren<QGroupBox*>("variablesOptions",Qt::FindDirectChildrenOnly).at(0);
    QGroupBox *multipliesOptions = stackedWidget->findChildren<QGroupBox*>("multipliesOptions",Qt::FindDirectChildrenOnly).at(0);
    QStringList variablesOptionsValues;
    QStringList multipliesOptionsValues;
    for(QWidget *widget: variablesOptions->findChildren<QWidget*>(QString(), Qt::FindDirectChildrenOnly)){
        if(QDoubleSpinBox *newWidget = qobject_cast<QDoubleSpinBox*>(widget)){
            variablesOptionsValues.append(QString::number(newWidget->value()));
        }else if(QSpinBox *newWidget = qobject_cast<QSpinBox*>(widget)){
            variablesOptionsValues.append(QString::number(newWidget->value()));
        }
    }
    for(QWidget *widget: multipliesOptions->findChildren<QWidget*>(QString(), Qt::FindDirectChildrenOnly)){
        if(QSpinBox *newWidget = qobject_cast<QSpinBox*>(widget)){
            multipliesOptionsValues.append(QString::number(newWidget->value()));
        }
    }
    m_combosMap.insert("variablesOptions",variablesOptionsValues);
    m_combosMap.insert("multipliesOptions",multipliesOptionsValues);
    QStringList vaziatHozor;
    vaziatHozor.append("حضور");
    vaziatHozor.append("نهست");
    vaziatHozor.append("ماموریت");
    vaziatHozor.append("مرخصی");
    vaziatHozor.append("فرار");
    m_combosMap.insert("وضعیت حضور",vaziatHozor);
    QStringList vaziatCard;
    vaziatCard.append("اقدام نگردیده");
    vaziatCard.append("اقدام گردیده");
    vaziatCard.append("بدهکار");
    vaziatCard.append("غیرقابل اقدام");
    m_combosMap.insert("وضعیت درخواست کارت",vaziatCard);
    stream << m_combosMap;
    file.close();
    QFile ostanFile(m_ostanFilename);
    if(!ostanFile.open(QIODevice::WriteOnly)){
        QMessageBox::critical(this,"Error Saving File",file.errorString());
        return;
    }
    ostanFile.close();
    QXlsx::Document doc(m_ostanFilename);
    doc.addSheet("Sheet1");
    doc.selectSheet("Sheet1");
    for(int h = 0; h < ostanTable->columnCount(); h++){
        doc.write(1,h+1,ostanTable->horizontalHeaderItem(h)->text());
        QXlsx::Format f = doc.columnFormat(h+1);
        f.setFont(this->font());
        f.setHorizontalAlignment(QXlsx::Format::HorizontalAlignment::AlignHCenter);
        f.setVerticalAlignment(QXlsx::Format::VerticalAlignment::AlignVCenter);
        doc.setColumnFormat(h+1,f);
        for(int r = 0; r < ostanTable->rowCount(); r++){
            doc.write(r+2,h+1,ostanTable->item(r,h)->text());
        }
    }
    doc.save();
    for(QString &sheetName: doc.sheetNames()){
        doc.selectSheet(sheetName);
        doc.autosizeColumnWidth();
    }
    doc.save();
    QString title = "ذخیره شد";
    QString message = "تنظیمات با موفقیت ذخیره سازی شد";
    QMessageBox::information(this,title,message);
    m_saved = true;
    accept();
}

void OptionDialog::loadOptions()
{
    QFile ostanFile (m_ostanFilename);
    if(!ostanFile.exists()){
        QFile::copy(":/options/ostanList.xlsx",m_ostanFilename);
        extern Q_CORE_EXPORT int qt_ntfs_permission_lookup;
        qt_ntfs_permission_lookup++;
        ostanFile.setPermissions(QFileDevice::WriteUser);
        qt_ntfs_permission_lookup--;
    }
    if(!ostanFile.exists()){
        QMessageBox::critical(this,"Error","فایل مربوط به استان ها یافت نشد");
        reject();
        return;
    }
    if(!ostanFile.open(QIODevice::ReadWrite)){
        QMessageBox::critical(this,"Error",ostanFile.errorString());
        reject();
        return;
    }
    ostanFile.close();
    QFile optionFile(m_optionFilename);
    if(!optionFile.exists()){
        QFile::copy(":/options/options.dat",m_optionFilename);
        extern Q_CORE_EXPORT int qt_ntfs_permission_lookup;
        qt_ntfs_permission_lookup++;
        optionFile.setPermissions(QFileDevice::WriteUser);
        qt_ntfs_permission_lookup--;
        if(!optionFile.exists()){
            QMessageBox::critical(this,"Error","فایل تنظیمات یافت نشد");
            reject();
            return;
        }
        QXlsx::Document doc(m_ostanFilename);
        doc.selectSheet(0);
        QStringList ostans;
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            ostans.append(doc.read(r,1).toString());
        }
        m_combosMap.insert("استان",ostans);
    }
    if(!optionFile.open(QIODevice::ReadOnly)){
        QMessageBox::critical(this,"Error",optionFile.errorString());
        reject();
        return;
    }
    QDataStream stream(&optionFile);
    stream >> m_combosMap;
    optionFile.close();
    QXlsx::Document doc(m_ostanFilename);
    doc.selectSheet(0);
    QStringList ostans;
    for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
        ostans.append(doc.read(r,1).toString());
    }
    m_combosMap.insert("استان",ostans);
    m_saved = true;
}

void OptionDialog::addToList()
{
    QStringList currentList = m_combosMap.value(listComboBox->currentText());
    QString currentText = textSearch->text().trimmed();
    if(currentText.isEmpty()){
        return;
    }
    if(currentList.contains(currentText)){
        return;
    }
    currentList.append(currentText);
    m_combosMap.insert(listComboBox->currentText(),currentList);
    for(QListWidget *list:stackedLists->findChildren<QListWidget*>(QString(), Qt::FindDirectChildrenOnly)){
        if (list->objectName().midRef(4).toInt() == stackedLists->currentIndex()) {
            list->addItem(currentText);
            if(listComboBox->currentText() == "استان"){
                ostanTable->insertRow(ostanTable->rowCount());
                QTableWidgetItem *item1 = new QTableWidgetItem(currentText);
                item1->setTextAlignment(Qt::AlignCenter);
                ostanTable->setItem(ostanTable->rowCount()-1,0,item1);
                QTableWidgetItem *item2 = new QTableWidgetItem(QString::number(0));
                item2->setTextAlignment(Qt::AlignCenter);
                ostanTable->setItem(ostanTable->rowCount()-1,1,item2);
            }
        }
    }
    m_saved = false;
}

void OptionDialog::removeFromList()
{
    QStringList currentList = m_combosMap.value(listComboBox->currentText());
    for(QListWidget *list:stackedLists->findChildren<QListWidget*>(QString(), Qt::FindDirectChildrenOnly)){
        if (list->objectName().midRef(4).toInt() == stackedLists->currentIndex()) {
            if(list->currentItem() == nullptr){
                return;
            }
            currentList.removeAt(list->currentRow());
            if(listComboBox->currentText() == "استان"){
                ostanTable->removeRow(list->currentRow());
            }
            list->takeItem(list->currentRow());
            list->removeItemWidget(list->currentItem());
            break;
        }
    }
    m_combosMap.insert(listComboBox->currentText(), currentList);
    m_saved = false;
}

void OptionDialog::cellDoubleClicked(int row, int column)
{
    bool ok = false;
    if(column == 0){
        QString title = "تغییر نام استان";
        QString label = "لطفا نام جدیدی برای استان";
        label += " " + ostanTable->item(row,0)->text() + " ";
        label += "وارد کنید";
        QString text = QInputDialog::getText(this,title,label,QLineEdit::Normal,ostanTable->item(row,column)->text(),&ok);
        if(ok){
            ostanTable->item(row,column)->setText(text);
            ostanTable->item(row,column)->setBackground(QBrush(QColor(Qt::yellow)));
            QStringList titles = Soldier::titles();
            QStringList codes = Soldier::titles(true);
            int j = 0;
            for(int i = 0; i < titles.count(); i++){
                if(codes.at(i) == "c"){
                    if(titles.at(i) == "استان"){
                        break;
                    } else {
                        j++;
                    }
                }
            }
            QListWidget *ostanList = qobject_cast<QListWidget *>(stackedLists->widget(j));
            ostanList->insertItem(row,text);
            delete ostanList->takeItem(row+1);
            QStringList ostans = m_combosMap.value("استان");
            ostans.replace(row,text);
            m_combosMap.insert("استان",ostans);
            m_saved = false;
        }
    } else {
        QString title = "تغییر ضریب توراهی استان";
        QString label = "لطفا ضریب توراهی برای استان";
        label += " " + ostanTable->item(row,0)->text() + " ";
        label += "وارد کنید";
        int value = QInputDialog::getInt(this,title,label,Soldier::PerToEngNo(ostanTable->item(row,column)->text()).toInt(),0,Soldier::maxSpinBox(),1,&ok);
        if(ok){
            ostanTable->item(row,column)->setText(QString::number(value));
            ostanTable->item(row,column)->setBackground(QBrush(QColor(Qt::yellow)));
            m_saved = false;
        }
    }
}

void OptionDialog::currentFontChanged(const QFont &font)
{
    m_saved = false;
}

void OptionDialog::ivalueChanged(int i)
{
    m_saved = false;
}

void OptionDialog::dvalueChanged(double d)
{
    m_saved = false;
}

QString OptionDialog::statisticsDir() const
{
    return m_statisticsDir;
}

QString OptionDialog::orderDir() const
{
    return m_orderDir;
}

void OptionDialog::closeEvent(QCloseEvent *event)
{
    if(!m_saved){
        QString message = "در تنظیمات تغییراتی اعمال کرده‌اید که ذخیره نشده است، آیا می‌خواهید تغییرات ذخیره شوند؟";
        QMessageBox::StandardButtons options = {QMessageBox::StandardButton::Save , QMessageBox::StandardButton::Discard , QMessageBox::StandardButton::Cancel};
        QMessageBox::StandardButton answer = QMessageBox::question(this, "File Not Saved",message, options);
        if(answer == QMessageBox::StandardButton::Save){
            saveOptions();
            event->accept();
        } else if (answer == QMessageBox::StandardButton::Discard){
            event->accept();
        } else {
            event->ignore();
        }
    } else {
        event->accept();
    }
}
