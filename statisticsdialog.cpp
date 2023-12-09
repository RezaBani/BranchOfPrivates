#include "statisticsdialog.h"
#include "ui_statisticsdialog.h"
#include "soldier.h"

StatisticsDialog::StatisticsDialog(const QString &filePath, const int &year, const int &month, QWidget *parent) :
    QDialog(parent),
    ui(new Ui::StatisticsDialog)
{
    ui->setupUi(this);
    m_filePath = filePath;
    this->setFont(parent->font());
    this->setWindowTitle("آمار پایه خدمتی");
    this->setWindowIcon(QIcon(":/icons/table.png"));
    this->setLayoutDirection(Qt::LayoutDirection::RightToLeft);
    tabWidget = new QTabWidget(this);
    tabWidget->setTabPosition(QTabWidget::South);

    QXlsx::Document doc(m_filePath);
    for(QString &sheetName:doc.sheetNames()){
        doc.selectSheet(sheetName);
        QTableWidget *table = new QTableWidget(tabWidget);
        table->setRowCount(doc.currentWorksheet()->dimension().rowCount());
        table->setColumnCount(doc.currentWorksheet()->dimension().columnCount());
        table->setSpan(0,1,1,12);
        table->setSpan(0,13,1,12);
        table->setSpan(0,25,1,12);
        table->setSpan(0,37,2,1);
        for(int r = 1; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            for(int c = 1; c <= doc.currentWorksheet()->dimension().columnCount(); c++){
                QTableWidgetItem *item = new QTableWidgetItem();
                item->setText(doc.read(r,c).toString());
                item->setTextAlignment(Qt::AlignCenter);
                table->setItem(r-1,c-1,item);
            }
        }
        table->setAlternatingRowColors(true);
        table->setEditTriggers(QAbstractItemView::NoEditTriggers);
        table->setSelectionMode(QAbstractItemView::SingleSelection);
        table->resizeColumnsToContents();
        tabWidget->addTab(table,sheetName);
        QScrollBar *customScrollBar = new QScrollBar(Qt::Horizontal,table);
        table->setHorizontalScrollBar(customScrollBar);
        customScrollBar->setLayoutDirection(Qt::LeftToRight);
        customScrollBar->setInvertedAppearance(true);
    }

    QLabel *infoLabel = new QLabel(this);
    QString infoText = "آمار پایه خدمتی";
    infoText += QString(" ") + Soldier::jalali().monthName(Soldier::persian(),month) + QString(" ");
    infoText += "ماه";
    infoText += QString(" ") + QString::number(year);
    infoLabel->setText(infoText);
    QPushButton *btnOpen = new QPushButton(this);
    btnOpen->setText("مشاهده فایل اکسل");
    QPushButton *btnClose = new QPushButton("Close",this);
    QHBoxLayout *hbox = new QHBoxLayout;
    hbox->addWidget(infoLabel,100);
    hbox->addWidget(btnOpen);
    hbox->addWidget(btnClose);
    QVBoxLayout *vbox = new QVBoxLayout;
    vbox->addWidget(tabWidget);
    vbox->addLayout(hbox);
    this->setLayout(vbox);

    connect(btnOpen,&QPushButton::clicked,this,&StatisticsDialog::btnOpenClicked);
    connect(btnClose,&QPushButton::clicked,this,&StatisticsDialog::close);
}

StatisticsDialog::~StatisticsDialog()
{
    delete ui;
}

void StatisticsDialog::btnOpenClicked()
{
    QDesktopServices::openUrl(QUrl::fromLocalFile(m_filePath));
}
