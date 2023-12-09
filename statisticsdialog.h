#ifndef STATISTICSDIALOG_H
#define STATISTICSDIALOG_H

#include <QDialog>
#include <QTabWidget>
#include <QTableWidget>
#include <QPushButton>
#include <QLabel>
#include <QPushButton>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QMessageBox>
#include <QDir>
#include <QFileInfo>
#include <QScrollBar>
#include <QDesktopServices>
#include <QPainter>
#include <xlsxdocument.h>
#include <xlsxchartsheet.h>
#include <xlsxcellrange.h>
#include <xlsxchart.h>
#include <xlsxrichstring.h>
#include <xlsxworkbook.h>

namespace Ui {
class StatisticsDialog;
}

class StatisticsDialog : public QDialog
{
    Q_OBJECT

public:
    explicit StatisticsDialog(const QString &filePath, const int &year, const int &month, QWidget *parent = nullptr);
    ~StatisticsDialog();

private slots:
    void btnOpenClicked();

private:
    Ui::StatisticsDialog *ui;
    QString m_filePath;
    QTabWidget *tabWidget;
};

#endif // STATISTICSDIALOG_H
