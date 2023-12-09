#ifndef ORDERDIALOG_H
#define ORDERDIALOG_H

#include <QDialog>
#include <QTabWidget>
#include <QTableWidget>
#include <QLabel>
#include <QPushButton>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QMessageBox>
#include <QDir>
#include <QFileInfo>
#include <QScrollBar>
#include <QDesktopServices>
#include <xlsxdocument.h>
#include <xlsxchartsheet.h>
#include <xlsxcellrange.h>
#include <xlsxchart.h>
#include <xlsxrichstring.h>
#include <xlsxworkbook.h>

namespace Ui {
class OrderDialog;
}

class OrderDialog : public QDialog
{
    Q_OBJECT

public:
    explicit OrderDialog(const QString &filePath, const int &orderNo, QWidget *parent = nullptr);
    ~OrderDialog();

private slots:
    void btnOpenClicked();

private:
    Ui::OrderDialog *ui;
    int m_orderNo;
    QString m_filePath;
    QTabWidget *tabWidget;
};

#endif // ORDERDIALOG_H
