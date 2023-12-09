#include "addeditdialog.h"
#include "soldier.h"
#include "ui_addeditdialog.h"

AddEditDialog::AddEditDialog(const QMap<QString,QString> &codesMap, const QMap<QString, QStringList> &combosMap, const int &tab, const int &row, QWidget *parent, const QStringList &peroperties):
    QDialog(parent), m_person(peroperties),
    ui(new Ui::AddEditDialog)
{
    ui->setupUi(this);
    if(peroperties.isEmpty()){
        this->setWindowTitle("افزودن سرباز جدید");
        this->setWindowIcon(QIcon(":/icons/add.png"));
    } else {
        this->setWindowTitle("ویرایش اطلاعات");
        this->setWindowIcon(QIcon(":/icons/edit.png"));
    }
    this->setLayoutDirection(Qt::LayoutDirection::RightToLeft);
    m_codes = codesMap;
    m_tab = tab;
    m_row = row;
    this->setFont(parent->font());

    //Setup Buttons
    QDialogButtonBox *btnBox = new QDialogButtonBox(this);
    QPushButton *btnMethod = new QPushButton(this);
    if(!m_person.isEmpty()){
        btnMethod->setText("Edit");
    } else {
        btnMethod->setText("Add");
    }
    btnBox->addButton(btnMethod,QDialogButtonBox::ActionRole);
    btnBox->addButton(QDialogButtonBox::Cancel);


    //Get the Labels
    labels = Soldier::titles();
    QStringList codes = Soldier::titles(true);
    QStringList tabs = Soldier::tabs();
    QStringList tabsCodes = Soldier::tabs(true);
    labels = Soldier::addLastTitles(labels,tabs.at(tab),tabsCodes.at(tab));
    codes = Soldier::addLastTitles(codes,tabs.at(tab),tabsCodes.at(tab),true);

    //Setup the Grid and Contols
    QVBoxLayout *vbox = new QVBoxLayout;
    QGridLayout *gridLayout = new QGridLayout;
    for(int i = 0; i < labels.count(); i++){
        QLabel *label = new QLabel(labels.at(i),this);
        if(i%2 == 0){
            gridLayout->addWidget(label,i,0);
        } else {
            QLabel *empty = new QLabel("\t",this);
            gridLayout->addWidget(empty,i-1,2);
            gridLayout->addWidget(label,i-1,3);
        }
        if(codes.at(i) == "t"){
            QLineEdit *lineEdit = new QLineEdit(this);
            lineEdit->setObjectName("Row"+QString::number(i));
            if(!m_person.isEmpty()){
                lineEdit->setText(m_person.at(i));
            }
            if(i%2 == 0){
                gridLayout->addWidget(lineEdit,i,1);
            } else {
                gridLayout->addWidget(lineEdit,i-1,4);
            }
        } else if (codes.at(i) == "d") {
            QDateEdit *dateEdit = new QDateEdit(this);
            dateEdit->setObjectName("Row"+QString::number(i));
            dateEdit->setCalendar(Soldier::jalali());
            dateEdit->setCalendarPopup(true);
            Soldier::setCalenderWidgetFormat(dateEdit->calendarWidget());
            dateEdit->setDisplayFormat(Soldier::displayFormat());
            dateEdit->setLocale(Soldier::persian());
            dateEdit->setMinimumWidth(dateEdit->width());
            if(!m_person.isEmpty()){
                dateEdit->setDate(QDate::fromString(Soldier::PerToEngNo(m_person.at(i)),Soldier::displayFormat(true),Soldier::jalali()));
            } else {
                dateEdit->setDate(QDate::currentDate());
            }
            if(i%2 == 0){
                gridLayout->addWidget(dateEdit,i,1);
            } else {
                gridLayout->addWidget(dateEdit,i-1,4);
            }
        } else if (codes.at(i) == "c"){
            QComboBox *comboWidget = new QComboBox(this);
            comboWidget->setObjectName("Row"+QString::number(i));
            QStringList values;
            if(labels.at(i).contains("محل")){
                values = combosMap.value(labels.at(i).mid(4),QStringList());
            } else {
                values = combosMap.value(labels.at(i),QStringList());
            }
            comboWidget->addItems(values);
            if(!m_person.isEmpty()){
                comboWidget->setCurrentText(m_person.at(i));
            }
            if(i%2 == 0){
                gridLayout->addWidget(comboWidget,i,1);
            } else {
                gridLayout->addWidget(comboWidget,i-1,4);
            }
        } else if (codes.at(i) == "n") {
            QSpinBox *spinBox = new QSpinBox(this);
            spinBox->setMaximum(Soldier::maxSpinBox());
            spinBox->setObjectName("Row"+QString::number(i));
            if(!m_person.isEmpty()){
                spinBox->setValue(m_person.at(i).toInt());
            }
            if(i%2 == 0){
                gridLayout->addWidget(spinBox,i,1);
            } else {
                gridLayout->addWidget(spinBox,i-1,4);
            }
        }
    }
    vbox->addLayout(gridLayout);
    vbox->addWidget(btnBox);
    this->setLayout(vbox);

    //Connecting Signal and Slots
    if(m_person.isEmpty()){
        connect(btnMethod,&QPushButton::clicked,this,&AddEditDialog::btnAddClicked);
    } else {
        connect(btnMethod,&QPushButton::clicked,this,&AddEditDialog::btnEditClicked);
    }
    connect(btnBox,&QDialogButtonBox::rejected,this,&QDialog::close);
}

AddEditDialog::~AddEditDialog()
{
    delete ui;
}

void AddEditDialog::btnAddClicked()
{
    QStringList peroperties;
    peroperties.reserve(labels.size());
    std::fill_n(std::back_inserter(peroperties), labels.size(), "");
    for(QWidget *widget:this->findChildren<QWidget *>(QString(), Qt::FindDirectChildrenOnly)){
        if(widget->objectName().contains("Row")){
            if(QLineEdit* newWidget = qobject_cast<QLineEdit*>(widget)){
                peroperties[newWidget->objectName().midRef(3).toInt()] = newWidget->text();
            }
            if(QComboBox* newWidget = qobject_cast<QComboBox*>(widget)){
                peroperties[newWidget->objectName().midRef(3).toInt()] = newWidget->currentText();
            }
            if(QSpinBox* newWidget = qobject_cast<QSpinBox*>(widget)){
                peroperties[newWidget->objectName().midRef(3).toInt()] = newWidget->text();
            }
            if(QDateEdit* newWidget = qobject_cast<QDateEdit*>(widget)){
                peroperties[newWidget->objectName().midRef(3).toInt()] = newWidget->text();
            }
        }
    }
    if(peroperties.at(2).isEmpty()){
        QString title = "کد ملی";
        QString message = "کد ملی نباید خالی باشد";
        QMessageBox::warning(this,title,message);
        return;
    }
    peroperties.replace(2,Soldier::PerToEngNo(peroperties.at(2)));
    if(Soldier::isDuplicate(peroperties.at(2),m_codes,m_tab,m_row)){
        QString title = "تکراری";
        QString message = "سربازی با همین کد ملی موجود است";
        QMessageBox::warning(this,title,message);
        return;
    }
    QString title = "افزودن سرباز";
    QString message = "";
    message += "آیا تمایل به اضافه کردن سرباز با مشخصات زیر به آمار هستید؟";
    for(int i = 0; i < peroperties.size(); i++){
        message += "\n" + labels.at(i) + ": " + peroperties.at(i);
    }
    QMessageBox::StandardButtons btns = (QMessageBox::Yes | QMessageBox::No);
    QMessageBox *msg = new QMessageBox(QMessageBox::Question, title, message, btns, this);
    msg->setWindowIcon(QIcon(":/icons/add.png"));
    msg->setFont(this->font());
    msg->setTextInteractionFlags(Qt::TextSelectableByMouse|Qt::TextSelectableByKeyboard);
    int answer = msg->exec();
    delete msg;
    switch (answer) {
    case QMessageBox::Yes:
        m_person = peroperties;
        break;
    case QMessageBox::No:
        break;
    default:
        break;
    }
    accept();
}

void AddEditDialog::btnEditClicked()
{
    QStringList peroperties;
    peroperties.reserve(labels.size());
    std::fill_n(std::back_inserter(peroperties), labels.size(), "");
    for(QWidget *widget:this->findChildren<QWidget *>(QString(), Qt::FindDirectChildrenOnly)){
        if(widget->objectName().contains("Row")){
            if(QLineEdit* newWidget = qobject_cast<QLineEdit*>(widget)){
                peroperties[newWidget->objectName().midRef(3).toInt()] = newWidget->text();
            }
            if(QComboBox* newWidget = qobject_cast<QComboBox*>(widget)){
                peroperties[newWidget->objectName().midRef(3).toInt()] = newWidget->currentText();
            }
            if(QSpinBox* newWidget = qobject_cast<QSpinBox*>(widget)){
                peroperties[newWidget->objectName().midRef(3).toInt()] = newWidget->text();
            }
            if(QDateEdit* newWidget = qobject_cast<QDateEdit*>(widget)){
                peroperties[newWidget->objectName().midRef(3).toInt()] = newWidget->text();
            }
        }
    }
    if(peroperties.at(2).isEmpty()){
        QString title = "کد ملی";
        QString message = "کد ملی نباید خالی باشد";
        QMessageBox::warning(this,title,message);
        return;
    }
    peroperties.replace(2,Soldier::PerToEngNo(peroperties.at(2)));
    if(Soldier::isDuplicate(peroperties.at(2),m_codes,m_tab,m_row)){
        QString title = "تکراری";
        QString message = "سربازی با همین کد ملی موجود است";
        QMessageBox::warning(this,title,message);
        return;
    }
    QString title = "ویرایش سرباز";
    QString message = "";
    message += "آیا تمایل به ثبت ویرایش سرباز با مشخصات زیر هستید؟";
    for(int i = 0; i < peroperties.size(); i++){
        message += "\n" + labels.at(i) + ": " + peroperties.at(i);
    }
    QMessageBox::StandardButtons btns = (QMessageBox::Yes | QMessageBox::No);
    QMessageBox *msg = new QMessageBox(QMessageBox::Question, title, message, btns, this);
    msg->setWindowIcon(QIcon(":/icons/edit.png"));
    msg->setFont(this->font());
    msg->setTextInteractionFlags(Qt::TextSelectableByMouse|Qt::TextSelectableByKeyboard);
    int answer = msg->exec();
    delete msg;
    switch (answer) {
    case QMessageBox::Yes:
        m_person = peroperties;
        break;
    case QMessageBox::No:
        break;
    default:
        break;
    }
    accept();
}

QStringList AddEditDialog::person() const
{
    return m_person;
}

void AddEditDialog::closeEvent(QCloseEvent *event)
{
    QString title = "لغو کردن";
    QString message = "آیا مطمئن هستید که قصد لغو کردن را دارید؟";
    QMessageBox::StandardButton answer = QMessageBox::question(this,title,message);
    if (answer == QMessageBox::StandardButton::Yes){
        event->accept();
    } else {
        event->ignore();
    }
}
