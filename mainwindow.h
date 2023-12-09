#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QMessageBox>
#include <QLabel>
#include <QLineEdit>
#include <QDateEdit>
#include <QToolBar>
#include <QTableWidget>
#include <QTableWidgetItem>
#include <QTabWidget>
#include <QScrollBar>
#include <QFile>
#include <QTextStream>
#include <QDataStream>
#include <QCloseEvent>
#include <QMultiMap>
#include <QFileDialog>
#include <QProgressBar>
#include <QInputDialog>
#include <QAction>
#include <QToolButton>
#include <QVariant>

//include QXlsx headers
#include <xlsxdocument.h>
#include <xlsxchartsheet.h>
#include <xlsxcellrange.h>
#include <xlsxchart.h>
#include <xlsxrichstring.h>
#include <xlsxworkbook.h>
#include <xlnt/xlnt.hpp>

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();
    QMessageBox::StandardButton saveRequest();
    void loadOptions();
    void updateRowColor(const int &tab, const int &row, const int &colorCode = -1, const QString &description = QString());


public slots:
    void changeColumnData(const int &columnIndex, const QString &columnNewData);
    void changeRowColor(const QString &tab, const int &rowIndex, const int &colorCode);
    void addLetter(const QString &tab, const QString &code, const QString &letterType, const QStringList &values);
    void removeLetter(const QString &tab, const QString &code, const int &letterIndex, const int &row);
    void moveToTable(const QString &tab, const QString &code, const QString &newTab, const QStringList &lastValues);
    void addToTable(const QString &tab, const QString &code, const QString &newTab, const QStringList &lastValues);
    void addToOrder(const QString &tab, const QString &code, const QString &sheetName, const QStringList &values);
    void changeInOrder(const QString &sheetName, const QStringList &values, const int &newValues);
    void removeOrder(const QString &sheetName, const QStringList &values);
    void summaryList(const QList<int> &list);

private slots:
    void actionSearchTriggered();
    void actionAddTriggered();
    void actionRemoveTriggered();
    void actionEditTriggered();
    void actionSaveTriggered();
    void actionLoadTriggered();
    void actionHideTriggered();
    void actionShowTriggered();
    void actionOptionTriggered();
    void actionLetterTriggered();
    void actionFilterTriggered();
    void actionClearTriggered();
    void actionHelpTriggered();
    void actionScreenTriggered();
    void actionSyncTriggered();
    void actionSyncAllTriggered();
    void actionSummaryTriggered();
    void actionOrderTriggered();
    void actionStatisticsTriggered();
    void updateStatusBar(int index);
    void cellDoubleClicked(int row, int column);

private:
    Ui::MainWindow *ui;
    QLineEdit *txtSearch;
    QTabWidget *tabWidget;
    QList<QTableWidget *> tables;
    QString m_dbFilename;
    QString m_ostanFilename;
    QString m_orderDir;
    QString m_statisticsDir;
    bool m_saved;
    bool m_skipSyncMessage;
    bool m_skipSummaryMessage;
    QMap<QString,QStringList> m_combosMap;
    QMap<QString,QString> m_codes;
    QMap<QString,QStringList> m_pendingLetter;
    QMap<QString,QStringList> m_pendingRemoveLetter;
    QMap<QString,QStringList> m_pendingMove;
    QMap<QString,QStringList> m_pendingRowColor;
    QMap<QString,QStringList> m_pendingOrder;
    QMap<QString,QStringList> m_pendingChangeOrder;
    QMap<QString,QStringList> m_pendingRemoveOrder;
    QMap<QString,QString> m_pendingCreatCodes;
    QStringList m_pendingChangeCode;
    QDateEdit *currentDate;
    QLabel *statusbarLabel;
    QList<int> summaryResult;
    QString m_tarkhis;
    QString m_ghanoni;
    QString m_mofid;

    // QWidget interface
protected:
    void closeEvent(QCloseEvent *event) override;
};

#endif // MAINWINDOW_H
