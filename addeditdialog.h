#ifndef ADDEDITDIALOG_H
#define ADDEDITDIALOG_H

#include <QDialog>
#include <QMessageBox>
#include <QLayout>
#include <QLabel>
#include <QPushButton>
#include <QDialogButtonBox>
#include <QComboBox>
#include <QDate>
#include <QCalendar>
#include <QLocale>
#include <QCalendarWidget>
#include <QLineEdit>
#include <QSpinBox>
#include <QDateEdit>
#include <QComboBox>
#include <QMap>
#include <QCloseEvent>

namespace Ui {
class AddEditDialog;
}

class AddEditDialog : public QDialog
{
    Q_OBJECT

public:
    explicit AddEditDialog(const QMap<QString,QString> &codesMap, const QMap<QString,QStringList> &combosMap, const int &tab = -1, const int &row = -1, QWidget *parent = nullptr, const QStringList &peroperties = QStringList());
    ~AddEditDialog();
    QStringList person() const;

private slots:
    void btnAddClicked();
    void btnEditClicked();

private:
    Ui::AddEditDialog *ui;
    QStringList m_person;
    QMap<QString,QString> m_codes;
    QMap<QString,QStringList> m_combosMap;
    QStringList labels;
    int m_tab;
    int m_row;

    // QWidget interface
protected:
    void closeEvent(QCloseEvent *event) override;
};

#endif // ADDEDITDIALOG_H
