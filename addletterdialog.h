#ifndef LETTERDIALOG_H
#define LETTERDIALOG_H

#include <QDialog>
#include <QLabel>
#include <QGroupBox>
#include <QComboBox>
#include <QSpinBox>
#include <QStackedWidget>
#include <QTabWidget>
#include <QTableWidget>
#include <QListWidget>
#include <QScrollBar>
#include <QPushButton>
#include <QCloseEvent>
#include <QLineEdit>
#include <QDateEdit>
#include <QCheckBox>
#include <QLayout>
#include <QSplitter>
#include <QDir>
#include <QFile>
#include <QFileInfoList>
#include <QFileInfo>
#include <QDataStream>
#include <QMessageBox>
#include <xlsxdocument.h>
#include <xlsxchartsheet.h>
#include <xlsxcellrange.h>
#include <xlsxchart.h>
#include <xlsxrichstring.h>
#include <xlsxworkbook.h>
#include "soldier.h"

namespace Ui {
class LetterDialog;
}

class LetterDialog : public QDialog
{
    Q_OBJECT

public:
    explicit LetterDialog(const QString &tab, const QString &code, QWidget *parent = nullptr);
    ~LetterDialog();

signals:
    void changeColumnData(const int &columnIndex, const QString &columnNewData);
    void changeRowColor(const QString &tab, const int &rowIndex, const int &colorCode);
    void addLetter(const QString &tab, const QString &code, const QString &letterType, const QStringList &values);
    void removeLetter(const QString &tab, const QString &code, const int &letterIndex, const int &row);
    void moveToTable(const QString &tab, const QString &code, const QString &newTab, const QStringList &lastValues);
    void addToTable(const QString &tab, const QString &code, const QString &newTab, const QStringList &lastValues);
    void addToOrder(const QString &tab, const QString &code, const QString &sheetName, const QStringList &values);

private slots:
    void btnAddClicked();
    void btnRemoveClicked();
    void btnOldClicked();
    void actionSearchTriggered();
    void sendChangeColumnData(const int &columnIndex, const QString &columnNewData);
    void sendChangeRowColor(const QString &tab, const int &rowIndex, const int &colorCode);
    void showLetterWarning(const QString &message);
    void sendMoveToTable(const QString &tab, const QString &code, const QString &newTab, const QStringList &lastValues);
    void sendAddToTable(const QString &tab, const QString &code, const QString &newTab, const QStringList &lastValues);
    void sendAddToOrder(const QString &tab, const QString &code, const QString &sheetName, const QStringList &values);

private:
    Ui::LetterDialog *ui;
    QLineEdit *txtSearch;
    QLabel *greenLabel;
    QPushButton *btnRemove;
    QTabWidget *tabWidget;
    QStackedWidget *stackedWidget;
    QListWidget *listOptions;
    QMap<QString,QStringList> m_combosMap;
    QString m_tab;
    QString m_code;
    void loadOption();

    // QWidget interface
protected:
    void closeEvent(QCloseEvent *event) override;
};

#endif // LETTERDIALOG_H
