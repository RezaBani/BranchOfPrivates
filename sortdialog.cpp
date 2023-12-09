#include "sortdialog.h"
#include "ui_sortdialog.h"

SortDialog::SortDialog(QWidget *parent) :
    QDialog(parent),
    ui(new Ui::SortDialog)
{
    ui->setupUi(this);
    this->setLayoutDirection(Qt::LayoutDirection::RightToLeft);
    this->setWindowTitle("مرتب کردن");
    this->setWindowIcon(QIcon(":/icons/sort.png"));
}

SortDialog::~SortDialog()
{
    delete ui;
}
