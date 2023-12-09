#ifndef OPTIONDIALOG_H
#define OPTIONDIALOG_H

#include <QDialog>
#include <QMessageBox>
#include <QPushButton>
#include <QDialogButtonBox>
#include <QLabel>
#include <QStackedWidget>
#include <QLineEdit>
#include <QFile>
#include <QDataStream>
#include <QSet>
#include <QStringList>
#include <QListWidget>
#include <QListView>
#include <QComboBox>
#include <QFontComboBox>
#include <QSpinBox>
#include <QGroupBox>
#include <QTableWidget>
#include <QScrollBar>
#include <QInputDialog>
#include <QLayout>
#include <QSplitter>
#include <QCloseEvent>
#include <xlsxdocument.h>
#include <xlsxchartsheet.h>
#include <xlsxcellrange.h>
#include <xlsxchart.h>
#include <xlsxrichstring.h>
#include <xlsxworkbook.h>

namespace Ui {
class OptionDialog;
}

class OptionDialog : public QDialog
{
    Q_OBJECT

public:
    explicit OptionDialog(QWidget *parent = nullptr, const bool &onlyGetFilename = false);
    ~OptionDialog();

    QString optionFilename() const;
    QString dbFilename() const;
    QString ostanFilename() const;
    QString orderDir() const;
    QString statisticsDir() const;
    QMap<QString, QStringList> combosMap() const;

private slots:
    void saveOptions();
    void loadOptions();
    void addToList();
    void removeFromList();
    void cellDoubleClicked(int row, int column);
    void currentFontChanged(const QFont &font);
    void ivalueChanged(int i);
    void dvalueChanged(double d);

private:
    Ui::OptionDialog *ui;
    QFontComboBox *fontComboBox;
    QSpinBox *fontSize;
    QSpinBox *orderNo;
    QLineEdit *textSearch;
    QComboBox *listComboBox;
    QStackedWidget *stackedWidget;
    QStackedWidget *stackedLists;
    QTableWidget *ostanTable;
    QString m_optionFilename;
    QString m_dbFilename;
    QString m_ostanFilename;
    QString m_orderDir;
    QString m_statisticsDir;
    QMap<QString,QStringList> m_combosMap;
    bool m_saved;

    // QWidget interface
protected:
    void closeEvent(QCloseEvent *event) override;
};

#endif // OPTIONDIALOG_H
