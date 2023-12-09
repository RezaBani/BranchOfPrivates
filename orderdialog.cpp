#include "orderdialog.h"
#include "ui_orderdialog.h"
#include "soldier.h"

OrderDialog::OrderDialog(const QString &filePath, const int &orderNo, QWidget *parent) :
    QDialog(parent),
    ui(new Ui::OrderDialog)
{
    ui->setupUi(this);
    m_orderNo = orderNo;
    m_filePath = filePath;
    this->setFont(parent->font());
    this->setWindowTitle("دستور");
    this->setWindowIcon(QIcon(":/icons/papers.png"));
    this->setLayoutDirection(Qt::LayoutDirection::RightToLeft);
    tabWidget = new QTabWidget(this);
    tabWidget->setTabPosition(QTabWidget::South);


    QStringList orders = Soldier::orders();
    QXlsx::Document doc(m_filePath);
    for(const QString &order:orders){
        doc.selectSheet(order);
        QTableWidget *table = new QTableWidget(tabWidget);
        table->setRowCount(doc.currentWorksheet()->dimension().rowCount()-1);
        table->setColumnCount(doc.currentWorksheet()->dimension().columnCount());
        QStringList header;
        for(int h = 1; h <= table->columnCount(); h++){
            header.append(doc.read(1,h).toString());
        }
        table->setHorizontalHeaderLabels(header);
        for(int r = 2; r <= doc.currentWorksheet()->dimension().rowCount(); r++){
            for(int c = 1; c <= doc.currentWorksheet()->dimension().columnCount(); c++){
                QTableWidgetItem *item = new QTableWidgetItem();
                item->setText(doc.read(r,c).toString());
                item->setTextAlignment(Qt::AlignCenter);
                table->setItem(r-2,c-1,item);
            }
        }
        table->setAlternatingRowColors(true);
        table->setEditTriggers(QAbstractItemView::NoEditTriggers);
        table->setSelectionMode(QAbstractItemView::SingleSelection);
        table->resizeColumnsToContents();
        tabWidget->addTab(table,order);
        QScrollBar *customScrollBar = new QScrollBar(Qt::Horizontal,table);
        table->setHorizontalScrollBar(customScrollBar);
        customScrollBar->setLayoutDirection(Qt::LeftToRight);
        customScrollBar->setInvertedAppearance(true);
    }
    QLabel *infoLabel = new QLabel(this);
    QString infoText;
    infoText = "شماره دستور فعلی";
    infoText += QString(" ") + QString::number(m_orderNo);
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

    connect(btnOpen,&QPushButton::clicked,this,&OrderDialog::btnOpenClicked);
    connect(btnClose,&QPushButton::clicked,this,&OrderDialog::close);
}

OrderDialog::~OrderDialog()
{
    delete ui;
}

void OrderDialog::btnOpenClicked()
{
    QDesktopServices::openUrl(QUrl::fromLocalFile(m_filePath));
}
