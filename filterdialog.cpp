#include "filterdialog.h"
#include "soldier.h"
#include "ui_filterdialog.h"

FilterDialog::FilterDialog(const int &tabIndex, const QMap<QString, QStringList> &combosMap, QWidget *parent) :
    QDialog(parent),
    ui(new Ui::FilterDialog)
{
    ui->setupUi(this);
    this->setLayoutDirection(Qt::LayoutDirection::RightToLeft);
    this->setWindowTitle("فیلتر");
    this->setWindowIcon(QIcon(":/icons/filter.png"));
    this->setFont(parent->font());
    m_combosMap = combosMap;

    QLabel *selectColumnLabel = new QLabel(this);
    selectColumnLabel->setText("انتخاب ستون");
    m_selectColumnBox = new QComboBox(this);
    m_selectColumnBox->setObjectName("m_selectColumnBox");
    m_titles = Soldier::titles();
    m_titlesCodes = Soldier::titles(true);
    m_titles = Soldier::addLastTitles(m_titles,Soldier::tabs().at(tabIndex),Soldier::tabs(true).at(tabIndex));
    m_titlesCodes = Soldier::addLastTitles(m_titlesCodes,Soldier::tabs(true).at(tabIndex),Soldier::tabs(true).at(tabIndex),true);
    m_selectColumnBox->addItems(m_titles);

    QLabel *selectValueLabel = new QLabel(this);
    selectValueLabel->setText("انتخاب فیلترها");
    listWidget = new QListWidget(this);

    QDialogButtonBox *btnBox = new QDialogButtonBox(this);
    QPushButton *btnApply = new QPushButton("Apply",this);
    btnBox->addButton(btnApply,QDialogButtonBox::ActionRole);
    btnBox->addButton(QDialogButtonBox::Cancel);

    btnAdd = new QPushButton(this);
    btnAdd->setIcon(QIcon(":/icons/add.png"));
    btnAdd->setToolTip("افزودن به لیست");
    btnAdd->setDefault(true);
    btnRemove = new QPushButton(this);
    btnRemove->setIcon(QIcon(":/icons/minus.png"));
    btnRemove->setToolTip("حذف از لیست");

    QVBoxLayout *vbox = new QVBoxLayout;
    hbox = new QHBoxLayout;
    gridBox = new QGridLayout;
    gridBox->addWidget(selectColumnLabel,0,0);
    gridBox->addWidget(m_selectColumnBox,0,1);
    gridBox->addWidget(selectValueLabel,1,0);
    gridBox->addLayout(hbox,1,1);
    vbox->addLayout(gridBox);
    vbox->addWidget(listWidget);
    vbox->addStretch();
    vbox->addWidget(btnBox);
    this->setLayout(vbox);

    connect(btnBox,&QDialogButtonBox::rejected,this,&QDialog::close);
    connect(btnApply,&QPushButton::clicked,this,&FilterDialog::btnApplyClicked);
    connect(m_selectColumnBox,QOverload<int>::of(&QComboBox::currentIndexChanged),this,&FilterDialog::selectColumnsIndexChanged);
    connect(btnAdd,&QPushButton::clicked,this,&FilterDialog::addToList);
    connect(btnRemove,&QPushButton::clicked,this,&FilterDialog::removeFromList);
    selectColumnsIndexChanged(0);
}

FilterDialog::~FilterDialog()
{
    delete ui;
}

void FilterDialog::btnApplyClicked()
{
    m_filteredColumn = m_selectColumnBox->currentIndex();
    for(int i = 0; i < listWidget->count(); i++){
        m_filteredValues.append(listWidget->item(i)->text());
    }
    if(m_filteredValues.isEmpty()){
        QMessageBox::warning(this,"Filter","لطفا حداقل یک فیلتر اضافه کنید");
        return;
    }
    accept();
}

void FilterDialog::selectColumnsIndexChanged(int index)
{
    QList<QWidget*> valueWidgetList = this->findChildren<QWidget*>("valueWidget",Qt::FindDirectChildrenOnly);
    for(QWidget *widget: valueWidgetList){
        delete widget;
    }
    QList<QCheckBox*> checkBoxWidgetList = this->findChildren<QCheckBox*>(QString(), Qt::FindDirectChildrenOnly);
    for(QCheckBox *widget: checkBoxWidgetList){
        delete widget;
    }
    if(m_titlesCodes.at(index) == "c"){
        QComboBox *comboBox = new QComboBox(this);
        comboBox->setObjectName("valueWidget");
        comboBox->addItems(m_combosMap.value(m_selectColumnBox->currentText()));
        hbox->addWidget(comboBox,100);
    } else if(m_titlesCodes.at(index) == "d"){
        QDateEdit *dateEdit = new QDateEdit(this);
        dateEdit->setObjectName("valueWidget");
        dateEdit->setCalendar(Soldier::jalali());
        dateEdit->setCalendarPopup(true);
        Soldier::setCalenderWidgetFormat(dateEdit->calendarWidget());
        dateEdit->setDisplayFormat(Soldier::displayFormat());
        dateEdit->setLocale(Soldier::persian());
        dateEdit->setDate(QDate::currentDate());
        dateEdit->setMinimumWidth(dateEdit->width());
        QCheckBox *dayCheck = new QCheckBox(this);
        dayCheck->setText("روز");
        dayCheck->setChecked(true);
        dayCheck->setObjectName("Check2");
        QCheckBox *monthCheck = new QCheckBox(this);
        monthCheck->setText("ماه");
        monthCheck->setChecked(true);
        monthCheck->setObjectName("Check1");
        QCheckBox *yearCheck = new QCheckBox(this);
        yearCheck->setText("سال");
        yearCheck->setChecked(true);
        yearCheck->setObjectName("Check0");
        hbox->addWidget(dateEdit,100);
        hbox->addWidget(dayCheck);
        hbox->addWidget(monthCheck);
        hbox->addWidget(yearCheck);
    } else if(m_titlesCodes.at(index) == "t"){
        QLineEdit *lineEdit = new QLineEdit(this);
        lineEdit->setObjectName("valueWidget");
        hbox->addWidget(lineEdit,100);
    } else if(m_titlesCodes.at(index) == "n"){
        QSpinBox *spinBox = new QSpinBox(this);
        spinBox->setMaximum(Soldier::maxSpinBox());
        spinBox->setObjectName("valueWidget");
        hbox->addWidget(spinBox,100);
    }
    hbox->addWidget(btnAdd);
    hbox->addWidget(btnRemove);
}

void FilterDialog::addToList()
{
    QString newValue;
    QWidget *widget = this->findChildren<QWidget*>("valueWidget",Qt::FindDirectChildrenOnly).at(0);
    if(QComboBox *newWidget = qobject_cast<QComboBox*>(widget)){
        newValue = newWidget->currentText().trimmed();
    } else if(QDateEdit *newWidget = qobject_cast<QDateEdit*>(widget)){
        QCheckBox *yearCheck = this->findChildren<QCheckBox*>("Check0",Qt::FindDirectChildrenOnly).at(0);
        QCheckBox *monthCheck = this->findChildren<QCheckBox*>("Check1",Qt::FindDirectChildrenOnly).at(0);
        QCheckBox *dayCheck = this->findChildren<QCheckBox*>("Check2",Qt::FindDirectChildrenOnly).at(0);
        if(yearCheck->isChecked()){
            newValue += newWidget->text().trimmed().split("/").at(0);
        }
        if(monthCheck->isChecked()){
            newValue += "/";
            newValue += newWidget->text().trimmed().split("/").at(1);
            if(!dayCheck->isChecked() && !yearCheck->isChecked()){
                newValue += "/";
            }
        }
        if(dayCheck->isChecked()){
            newValue += "/";
            newValue += newWidget->text().trimmed().split("/").at(2);
        }
    } else if(QLineEdit *newWidget = qobject_cast<QLineEdit*>(widget)){
        newValue = newWidget->text().trimmed();
    } else if(QSpinBox *newWidget = qobject_cast<QSpinBox*>(widget)){
        newValue = QString::number(newWidget->value()).trimmed();
    }
    for(int i = 0; i < listWidget->count(); i++){
        if (listWidget->item(i)->text() == newValue){
            return;
        }
    }
    listWidget->addItem(newValue);
    m_selectColumnBox->setDisabled(true);
}

void FilterDialog::removeFromList()
{
    if(listWidget->count() == 0 || listWidget->currentItem() == nullptr){
        return;
    }
    delete listWidget->takeItem(listWidget->currentRow());
}

int FilterDialog::filteredColumn() const
{
    return m_filteredColumn;
}

QStringList FilterDialog::filteredValues() const
{
    return m_filteredValues;
}
