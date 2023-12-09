#ifndef FILTERDIALOG_H
#define FILTERDIALOG_H

#include <QLayout>
#include <QFile>
#include <QDialog>
#include <QDialogButtonBox>
#include <QPushButton>
#include <QLabel>
#include <QComboBox>
#include <QDateEdit>
#include <QListWidget>
#include <QLineEdit>
#include <QMessageBox>
#include <QSpinBox>
#include <QCheckBox>

namespace Ui {
class FilterDialog;
}

class FilterDialog : public QDialog
{
    Q_OBJECT

public:
    explicit FilterDialog(const int &tabIndex, const QMap<QString,QStringList> &combosMap, QWidget *parent = nullptr);
    ~FilterDialog();
    void loadOptions();
    QStringList filteredValues() const;
    int filteredColumn() const;

private slots:
    void btnApplyClicked();
    void selectColumnsIndexChanged(int index);
    void addToList();
    void removeFromList();

private:
    Ui::FilterDialog *ui;
    QMap<QString,QStringList> m_combosMap;
    QGridLayout *gridBox;
    QHBoxLayout *hbox;
    QPushButton *btnAdd;
    QPushButton *btnRemove;
    QListWidget *listWidget;
    QComboBox *m_selectColumnBox;
    QStringList m_titles;
    QStringList m_titlesCodes;
    QStringList m_filteredValues;
    int m_filteredColumn;
};

#endif // FILTERDIALOG_H
