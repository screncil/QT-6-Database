/********************************************************************************
** Form generated from reading UI file 'mainwindow.ui'
**
** Created by: Qt User Interface Compiler version 6.7.0
**
** WARNING! All changes made in this file will be lost when recompiling UI file!
********************************************************************************/

#ifndef UI_MAINWINDOW_H
#define UI_MAINWINDOW_H

#include <QtCore/QVariant>
#include <QtGui/QAction>
#include <QtWidgets/QApplication>
#include <QtWidgets/QGroupBox>
#include <QtWidgets/QHeaderView>
#include <QtWidgets/QLineEdit>
#include <QtWidgets/QMainWindow>
#include <QtWidgets/QMenu>
#include <QtWidgets/QMenuBar>
#include <QtWidgets/QPushButton>
#include <QtWidgets/QStatusBar>
#include <QtWidgets/QTableView>
#include <QtWidgets/QToolBar>
#include <QtWidgets/QWidget>

QT_BEGIN_NAMESPACE

class Ui_MainWindow
{
public:
    QAction *actionSort;
    QAction *actionAverage;
    QAction *action_Excel;
    QAction *action_Word;
    QAction *action_Access;
    QAction *actionFullInfo;
    QAction *actionInform;
    QAction *actionClose;
    QAction *actionInfoApp;
    QAction *OpenProvidnyk;
    QWidget *centralwidget;
    QGroupBox *Average_2;
    QLineEdit *lineEdit;
    QPushButton *Average;
    QGroupBox *groupBox_2;
    QPushButton *OpenExcel;
    QPushButton *OpenWord;
    QPushButton *OpenAccess;
    QTableView *tableView;
    QGroupBox *Operations;
    QPushButton *Add;
    QPushButton *Delete;
    QPushButton *Download;
    QPushButton *Sort;
    QGroupBox *groupBox;
    QLineEdit *searchLine;
    QPushButton *Search;
    QMenuBar *menubar;
    QMenu *menuoptions;
    QMenu *menu;
    QMenu *menu_2;
    QStatusBar *statusbar;
    QToolBar *toolBar;

    void setupUi(QMainWindow *MainWindow)
    {
        if (MainWindow->objectName().isEmpty())
            MainWindow->setObjectName("MainWindow");
        MainWindow->resize(1440, 932);
        actionSort = new QAction(MainWindow);
        actionSort->setObjectName("actionSort");
        actionAverage = new QAction(MainWindow);
        actionAverage->setObjectName("actionAverage");
        action_Excel = new QAction(MainWindow);
        action_Excel->setObjectName("action_Excel");
        action_Word = new QAction(MainWindow);
        action_Word->setObjectName("action_Word");
        action_Access = new QAction(MainWindow);
        action_Access->setObjectName("action_Access");
        actionFullInfo = new QAction(MainWindow);
        actionFullInfo->setObjectName("actionFullInfo");
        actionInform = new QAction(MainWindow);
        actionInform->setObjectName("actionInform");
        actionClose = new QAction(MainWindow);
        actionClose->setObjectName("actionClose");
        actionInfoApp = new QAction(MainWindow);
        actionInfoApp->setObjectName("actionInfoApp");
        OpenProvidnyk = new QAction(MainWindow);
        OpenProvidnyk->setObjectName("OpenProvidnyk");
        centralwidget = new QWidget(MainWindow);
        centralwidget->setObjectName("centralwidget");
        Average_2 = new QGroupBox(centralwidget);
        Average_2->setObjectName("Average_2");
        Average_2->setGeometry(QRect(1050, 20, 221, 171));
        lineEdit = new QLineEdit(Average_2);
        lineEdit->setObjectName("lineEdit");
        lineEdit->setGeometry(QRect(0, 50, 221, 41));
        lineEdit->setReadOnly(true);
        Average = new QPushButton(Average_2);
        Average->setObjectName("Average");
        Average->setGeometry(QRect(0, 100, 221, 51));
        groupBox_2 = new QGroupBox(centralwidget);
        groupBox_2->setObjectName("groupBox_2");
        groupBox_2->setGeometry(QRect(780, 500, 501, 191));
        OpenExcel = new QPushButton(groupBox_2);
        OpenExcel->setObjectName("OpenExcel");
        OpenExcel->setGeometry(QRect(0, 40, 501, 41));
        OpenWord = new QPushButton(groupBox_2);
        OpenWord->setObjectName("OpenWord");
        OpenWord->setGeometry(QRect(0, 90, 501, 41));
        OpenAccess = new QPushButton(groupBox_2);
        OpenAccess->setObjectName("OpenAccess");
        OpenAccess->setGeometry(QRect(0, 140, 501, 41));
        tableView = new QTableView(centralwidget);
        tableView->setObjectName("tableView");
        tableView->setGeometry(QRect(20, 40, 731, 661));
        tableView->setContextMenuPolicy(Qt::DefaultContextMenu);
        tableView->setSortingEnabled(false);
        Operations = new QGroupBox(centralwidget);
        Operations->setObjectName("Operations");
        Operations->setGeometry(QRect(780, 200, 501, 231));
        Add = new QPushButton(Operations);
        Add->setObjectName("Add");
        Add->setGeometry(QRect(40, 170, 201, 41));
        QPalette palette;
        QBrush brush(QColor(0, 0, 0, 255));
        brush.setStyle(Qt::SolidPattern);
        palette.setBrush(QPalette::Active, QPalette::WindowText, brush);
        QBrush brush1(QColor(57, 147, 78, 255));
        brush1.setStyle(Qt::SolidPattern);
        palette.setBrush(QPalette::Active, QPalette::Button, brush1);
        QBrush brush2(QColor(86, 220, 117, 255));
        brush2.setStyle(Qt::SolidPattern);
        palette.setBrush(QPalette::Active, QPalette::Light, brush2);
        QBrush brush3(QColor(71, 183, 97, 255));
        brush3.setStyle(Qt::SolidPattern);
        palette.setBrush(QPalette::Active, QPalette::Midlight, brush3);
        QBrush brush4(QColor(28, 73, 39, 255));
        brush4.setStyle(Qt::SolidPattern);
        palette.setBrush(QPalette::Active, QPalette::Dark, brush4);
        QBrush brush5(QColor(38, 98, 52, 255));
        brush5.setStyle(Qt::SolidPattern);
        palette.setBrush(QPalette::Active, QPalette::Mid, brush5);
        palette.setBrush(QPalette::Active, QPalette::Text, brush);
        QBrush brush6(QColor(255, 255, 255, 255));
        brush6.setStyle(Qt::SolidPattern);
        palette.setBrush(QPalette::Active, QPalette::BrightText, brush6);
        palette.setBrush(QPalette::Active, QPalette::ButtonText, brush);
        palette.setBrush(QPalette::Active, QPalette::Base, brush6);
        palette.setBrush(QPalette::Active, QPalette::Window, brush1);
        palette.setBrush(QPalette::Active, QPalette::Shadow, brush);
        QBrush brush7(QColor(156, 201, 166, 255));
        brush7.setStyle(Qt::SolidPattern);
        palette.setBrush(QPalette::Active, QPalette::AlternateBase, brush7);
        QBrush brush8(QColor(255, 255, 220, 255));
        brush8.setStyle(Qt::SolidPattern);
        palette.setBrush(QPalette::Active, QPalette::ToolTipBase, brush8);
        palette.setBrush(QPalette::Active, QPalette::ToolTipText, brush);
        QBrush brush9(QColor(0, 0, 0, 127));
        brush9.setStyle(Qt::SolidPattern);
#if QT_VERSION >= QT_VERSION_CHECK(5, 12, 0)
        palette.setBrush(QPalette::Active, QPalette::PlaceholderText, brush9);
#endif
        palette.setBrush(QPalette::Active, QPalette::Accent, brush6);
        palette.setBrush(QPalette::Inactive, QPalette::WindowText, brush);
        palette.setBrush(QPalette::Inactive, QPalette::Button, brush1);
        palette.setBrush(QPalette::Inactive, QPalette::Light, brush2);
        palette.setBrush(QPalette::Inactive, QPalette::Midlight, brush3);
        palette.setBrush(QPalette::Inactive, QPalette::Dark, brush4);
        palette.setBrush(QPalette::Inactive, QPalette::Mid, brush5);
        palette.setBrush(QPalette::Inactive, QPalette::Text, brush);
        palette.setBrush(QPalette::Inactive, QPalette::BrightText, brush6);
        palette.setBrush(QPalette::Inactive, QPalette::ButtonText, brush);
        palette.setBrush(QPalette::Inactive, QPalette::Base, brush6);
        palette.setBrush(QPalette::Inactive, QPalette::Window, brush1);
        palette.setBrush(QPalette::Inactive, QPalette::Shadow, brush);
        palette.setBrush(QPalette::Inactive, QPalette::AlternateBase, brush7);
        palette.setBrush(QPalette::Inactive, QPalette::ToolTipBase, brush8);
        palette.setBrush(QPalette::Inactive, QPalette::ToolTipText, brush);
#if QT_VERSION >= QT_VERSION_CHECK(5, 12, 0)
        palette.setBrush(QPalette::Inactive, QPalette::PlaceholderText, brush9);
#endif
        palette.setBrush(QPalette::Inactive, QPalette::Accent, brush6);
        palette.setBrush(QPalette::Disabled, QPalette::WindowText, brush4);
        palette.setBrush(QPalette::Disabled, QPalette::Button, brush1);
        palette.setBrush(QPalette::Disabled, QPalette::Light, brush2);
        palette.setBrush(QPalette::Disabled, QPalette::Midlight, brush3);
        palette.setBrush(QPalette::Disabled, QPalette::Dark, brush4);
        palette.setBrush(QPalette::Disabled, QPalette::Mid, brush5);
        palette.setBrush(QPalette::Disabled, QPalette::Text, brush4);
        palette.setBrush(QPalette::Disabled, QPalette::BrightText, brush6);
        palette.setBrush(QPalette::Disabled, QPalette::ButtonText, brush4);
        palette.setBrush(QPalette::Disabled, QPalette::Base, brush1);
        palette.setBrush(QPalette::Disabled, QPalette::Window, brush1);
        palette.setBrush(QPalette::Disabled, QPalette::Shadow, brush);
        palette.setBrush(QPalette::Disabled, QPalette::AlternateBase, brush1);
        palette.setBrush(QPalette::Disabled, QPalette::ToolTipBase, brush8);
        palette.setBrush(QPalette::Disabled, QPalette::ToolTipText, brush);
        QBrush brush10(QColor(28, 73, 39, 127));
        brush10.setStyle(Qt::SolidPattern);
#if QT_VERSION >= QT_VERSION_CHECK(5, 12, 0)
        palette.setBrush(QPalette::Disabled, QPalette::PlaceholderText, brush10);
#endif
        QBrush brush11(QColor(74, 191, 101, 255));
        brush11.setStyle(Qt::SolidPattern);
        palette.setBrush(QPalette::Disabled, QPalette::Accent, brush11);
        Add->setPalette(palette);
        Delete = new QPushButton(Operations);
        Delete->setObjectName("Delete");
        Delete->setGeometry(QRect(270, 170, 201, 41));
        QPalette palette1;
        palette1.setBrush(QPalette::Active, QPalette::WindowText, brush);
        QBrush brush12(QColor(139, 54, 55, 255));
        brush12.setStyle(Qt::SolidPattern);
        palette1.setBrush(QPalette::Active, QPalette::Button, brush12);
        QBrush brush13(QColor(208, 81, 83, 255));
        brush13.setStyle(Qt::SolidPattern);
        palette1.setBrush(QPalette::Active, QPalette::Light, brush13);
        QBrush brush14(QColor(173, 67, 69, 255));
        brush14.setStyle(Qt::SolidPattern);
        palette1.setBrush(QPalette::Active, QPalette::Midlight, brush14);
        QBrush brush15(QColor(69, 27, 28, 255));
        brush15.setStyle(Qt::SolidPattern);
        palette1.setBrush(QPalette::Active, QPalette::Dark, brush15);
        QBrush brush16(QColor(93, 36, 37, 255));
        brush16.setStyle(Qt::SolidPattern);
        palette1.setBrush(QPalette::Active, QPalette::Mid, brush16);
        palette1.setBrush(QPalette::Active, QPalette::Text, brush);
        palette1.setBrush(QPalette::Active, QPalette::BrightText, brush6);
        palette1.setBrush(QPalette::Active, QPalette::ButtonText, brush);
        palette1.setBrush(QPalette::Active, QPalette::Base, brush6);
        palette1.setBrush(QPalette::Active, QPalette::Window, brush12);
        palette1.setBrush(QPalette::Active, QPalette::Shadow, brush);
        QBrush brush17(QColor(197, 154, 155, 255));
        brush17.setStyle(Qt::SolidPattern);
        palette1.setBrush(QPalette::Active, QPalette::AlternateBase, brush17);
        palette1.setBrush(QPalette::Active, QPalette::ToolTipBase, brush8);
        palette1.setBrush(QPalette::Active, QPalette::ToolTipText, brush);
#if QT_VERSION >= QT_VERSION_CHECK(5, 12, 0)
        palette1.setBrush(QPalette::Active, QPalette::PlaceholderText, brush9);
#endif
        palette1.setBrush(QPalette::Active, QPalette::Accent, brush6);
        palette1.setBrush(QPalette::Inactive, QPalette::WindowText, brush);
        palette1.setBrush(QPalette::Inactive, QPalette::Button, brush12);
        palette1.setBrush(QPalette::Inactive, QPalette::Light, brush13);
        palette1.setBrush(QPalette::Inactive, QPalette::Midlight, brush14);
        palette1.setBrush(QPalette::Inactive, QPalette::Dark, brush15);
        palette1.setBrush(QPalette::Inactive, QPalette::Mid, brush16);
        palette1.setBrush(QPalette::Inactive, QPalette::Text, brush);
        palette1.setBrush(QPalette::Inactive, QPalette::BrightText, brush6);
        palette1.setBrush(QPalette::Inactive, QPalette::ButtonText, brush);
        palette1.setBrush(QPalette::Inactive, QPalette::Base, brush6);
        palette1.setBrush(QPalette::Inactive, QPalette::Window, brush12);
        palette1.setBrush(QPalette::Inactive, QPalette::Shadow, brush);
        palette1.setBrush(QPalette::Inactive, QPalette::AlternateBase, brush17);
        palette1.setBrush(QPalette::Inactive, QPalette::ToolTipBase, brush8);
        palette1.setBrush(QPalette::Inactive, QPalette::ToolTipText, brush);
#if QT_VERSION >= QT_VERSION_CHECK(5, 12, 0)
        palette1.setBrush(QPalette::Inactive, QPalette::PlaceholderText, brush9);
#endif
        palette1.setBrush(QPalette::Inactive, QPalette::Accent, brush6);
        palette1.setBrush(QPalette::Disabled, QPalette::WindowText, brush15);
        palette1.setBrush(QPalette::Disabled, QPalette::Button, brush12);
        palette1.setBrush(QPalette::Disabled, QPalette::Light, brush13);
        palette1.setBrush(QPalette::Disabled, QPalette::Midlight, brush14);
        palette1.setBrush(QPalette::Disabled, QPalette::Dark, brush15);
        palette1.setBrush(QPalette::Disabled, QPalette::Mid, brush16);
        palette1.setBrush(QPalette::Disabled, QPalette::Text, brush15);
        palette1.setBrush(QPalette::Disabled, QPalette::BrightText, brush6);
        palette1.setBrush(QPalette::Disabled, QPalette::ButtonText, brush15);
        palette1.setBrush(QPalette::Disabled, QPalette::Base, brush12);
        palette1.setBrush(QPalette::Disabled, QPalette::Window, brush12);
        palette1.setBrush(QPalette::Disabled, QPalette::Shadow, brush);
        palette1.setBrush(QPalette::Disabled, QPalette::AlternateBase, brush12);
        palette1.setBrush(QPalette::Disabled, QPalette::ToolTipBase, brush8);
        palette1.setBrush(QPalette::Disabled, QPalette::ToolTipText, brush);
        QBrush brush18(QColor(69, 27, 28, 127));
        brush18.setStyle(Qt::SolidPattern);
#if QT_VERSION >= QT_VERSION_CHECK(5, 12, 0)
        palette1.setBrush(QPalette::Disabled, QPalette::PlaceholderText, brush18);
#endif
        QBrush brush19(QColor(181, 70, 72, 255));
        brush19.setStyle(Qt::SolidPattern);
        palette1.setBrush(QPalette::Disabled, QPalette::Accent, brush19);
        Delete->setPalette(palette1);
        Download = new QPushButton(Operations);
        Download->setObjectName("Download");
        Download->setGeometry(QRect(40, 40, 431, 51));
        Sort = new QPushButton(Operations);
        Sort->setObjectName("Sort");
        Sort->setGeometry(QRect(40, 100, 431, 61));
        QPalette palette2;
        palette2.setBrush(QPalette::Active, QPalette::WindowText, brush);
        QBrush brush20(QColor(26, 67, 162, 255));
        brush20.setStyle(Qt::SolidPattern);
        palette2.setBrush(QPalette::Active, QPalette::Button, brush20);
        QBrush brush21(QColor(39, 101, 243, 255));
        brush21.setStyle(Qt::SolidPattern);
        palette2.setBrush(QPalette::Active, QPalette::Light, brush21);
        QBrush brush22(QColor(32, 84, 202, 255));
        brush22.setStyle(Qt::SolidPattern);
        palette2.setBrush(QPalette::Active, QPalette::Midlight, brush22);
        QBrush brush23(QColor(13, 34, 81, 255));
        brush23.setStyle(Qt::SolidPattern);
        palette2.setBrush(QPalette::Active, QPalette::Dark, brush23);
        QBrush brush24(QColor(17, 45, 108, 255));
        brush24.setStyle(Qt::SolidPattern);
        palette2.setBrush(QPalette::Active, QPalette::Mid, brush24);
        palette2.setBrush(QPalette::Active, QPalette::Text, brush);
        palette2.setBrush(QPalette::Active, QPalette::BrightText, brush6);
        palette2.setBrush(QPalette::Active, QPalette::ButtonText, brush);
        palette2.setBrush(QPalette::Active, QPalette::Base, brush6);
        palette2.setBrush(QPalette::Active, QPalette::Window, brush20);
        palette2.setBrush(QPalette::Active, QPalette::Shadow, brush);
        QBrush brush25(QColor(140, 161, 208, 255));
        brush25.setStyle(Qt::SolidPattern);
        palette2.setBrush(QPalette::Active, QPalette::AlternateBase, brush25);
        palette2.setBrush(QPalette::Active, QPalette::ToolTipBase, brush8);
        palette2.setBrush(QPalette::Active, QPalette::ToolTipText, brush);
#if QT_VERSION >= QT_VERSION_CHECK(5, 12, 0)
        palette2.setBrush(QPalette::Active, QPalette::PlaceholderText, brush9);
#endif
        palette2.setBrush(QPalette::Active, QPalette::Accent, brush6);
        palette2.setBrush(QPalette::Inactive, QPalette::WindowText, brush);
        palette2.setBrush(QPalette::Inactive, QPalette::Button, brush20);
        palette2.setBrush(QPalette::Inactive, QPalette::Light, brush21);
        palette2.setBrush(QPalette::Inactive, QPalette::Midlight, brush22);
        palette2.setBrush(QPalette::Inactive, QPalette::Dark, brush23);
        palette2.setBrush(QPalette::Inactive, QPalette::Mid, brush24);
        palette2.setBrush(QPalette::Inactive, QPalette::Text, brush);
        palette2.setBrush(QPalette::Inactive, QPalette::BrightText, brush6);
        palette2.setBrush(QPalette::Inactive, QPalette::ButtonText, brush);
        palette2.setBrush(QPalette::Inactive, QPalette::Base, brush6);
        palette2.setBrush(QPalette::Inactive, QPalette::Window, brush20);
        palette2.setBrush(QPalette::Inactive, QPalette::Shadow, brush);
        palette2.setBrush(QPalette::Inactive, QPalette::AlternateBase, brush25);
        palette2.setBrush(QPalette::Inactive, QPalette::ToolTipBase, brush8);
        palette2.setBrush(QPalette::Inactive, QPalette::ToolTipText, brush);
#if QT_VERSION >= QT_VERSION_CHECK(5, 12, 0)
        palette2.setBrush(QPalette::Inactive, QPalette::PlaceholderText, brush9);
#endif
        palette2.setBrush(QPalette::Inactive, QPalette::Accent, brush6);
        palette2.setBrush(QPalette::Disabled, QPalette::WindowText, brush23);
        palette2.setBrush(QPalette::Disabled, QPalette::Button, brush20);
        palette2.setBrush(QPalette::Disabled, QPalette::Light, brush21);
        palette2.setBrush(QPalette::Disabled, QPalette::Midlight, brush22);
        palette2.setBrush(QPalette::Disabled, QPalette::Dark, brush23);
        palette2.setBrush(QPalette::Disabled, QPalette::Mid, brush24);
        palette2.setBrush(QPalette::Disabled, QPalette::Text, brush23);
        palette2.setBrush(QPalette::Disabled, QPalette::BrightText, brush6);
        palette2.setBrush(QPalette::Disabled, QPalette::ButtonText, brush23);
        palette2.setBrush(QPalette::Disabled, QPalette::Base, brush20);
        palette2.setBrush(QPalette::Disabled, QPalette::Window, brush20);
        palette2.setBrush(QPalette::Disabled, QPalette::Shadow, brush);
        palette2.setBrush(QPalette::Disabled, QPalette::AlternateBase, brush20);
        palette2.setBrush(QPalette::Disabled, QPalette::ToolTipBase, brush8);
        palette2.setBrush(QPalette::Disabled, QPalette::ToolTipText, brush);
        QBrush brush26(QColor(13, 34, 81, 127));
        brush26.setStyle(Qt::SolidPattern);
#if QT_VERSION >= QT_VERSION_CHECK(5, 12, 0)
        palette2.setBrush(QPalette::Disabled, QPalette::PlaceholderText, brush26);
#endif
        QBrush brush27(QColor(34, 87, 211, 255));
        brush27.setStyle(Qt::SolidPattern);
        palette2.setBrush(QPalette::Disabled, QPalette::Accent, brush27);
        Sort->setPalette(palette2);
        groupBox = new QGroupBox(centralwidget);
        groupBox->setObjectName("groupBox");
        groupBox->setGeometry(QRect(780, 20, 221, 171));
        searchLine = new QLineEdit(groupBox);
        searchLine->setObjectName("searchLine");
        searchLine->setGeometry(QRect(0, 50, 221, 41));
        searchLine->setReadOnly(false);
        Search = new QPushButton(groupBox);
        Search->setObjectName("Search");
        Search->setGeometry(QRect(0, 100, 221, 51));
        MainWindow->setCentralWidget(centralwidget);
        menubar = new QMenuBar(MainWindow);
        menubar->setObjectName("menubar");
        menubar->setGeometry(QRect(0, 0, 1440, 25));
        menuoptions = new QMenu(menubar);
        menuoptions->setObjectName("menuoptions");
        menu = new QMenu(menuoptions);
        menu->setObjectName("menu");
        menu_2 = new QMenu(menuoptions);
        menu_2->setObjectName("menu_2");
        MainWindow->setMenuBar(menubar);
        statusbar = new QStatusBar(MainWindow);
        statusbar->setObjectName("statusbar");
        MainWindow->setStatusBar(statusbar);
        toolBar = new QToolBar(MainWindow);
        toolBar->setObjectName("toolBar");
        MainWindow->addToolBar(Qt::ToolBarArea::TopToolBarArea, toolBar);

        menubar->addAction(menuoptions->menuAction());
        menuoptions->addSeparator();
        menuoptions->addAction(actionSort);
        menuoptions->addAction(actionAverage);
        menuoptions->addAction(menu->menuAction());
        menuoptions->addAction(menu_2->menuAction());
        menuoptions->addSeparator();
        menuoptions->addAction(actionClose);
        menu->addAction(action_Excel);
        menu->addAction(action_Word);
        menu->addAction(action_Access);
        menu_2->addAction(actionFullInfo);
        menu_2->addAction(actionInform);
        menu_2->addAction(actionInfoApp);

        retranslateUi(MainWindow);

        QMetaObject::connectSlotsByName(MainWindow);
    } // setupUi

    void retranslateUi(QMainWindow *MainWindow)
    {
        MainWindow->setWindowTitle(QCoreApplication::translate("MainWindow", "MainWindow", nullptr));
        actionSort->setText(QCoreApplication::translate("MainWindow", "\320\241\320\276\321\200\321\202\321\203\320\262\320\260\321\202\320\270 \320\267\320\260 \321\206\321\226\320\275\320\276\321\216", nullptr));
#if QT_CONFIG(shortcut)
        actionSort->setShortcut(QCoreApplication::translate("MainWindow", "Ctrl+X", nullptr));
#endif // QT_CONFIG(shortcut)
        actionAverage->setText(QCoreApplication::translate("MainWindow", "\320\241\320\265\321\200\320\265\320\264\320\275\321\217 \321\206\321\226\320\275\320\260", nullptr));
#if QT_CONFIG(shortcut)
        actionAverage->setShortcut(QCoreApplication::translate("MainWindow", "Ctrl+A", nullptr));
#endif // QT_CONFIG(shortcut)
        action_Excel->setText(QCoreApplication::translate("MainWindow", "\320\222\321\226\320\264\320\272\321\200\320\270\321\202\320\270 Excel", nullptr));
        action_Word->setText(QCoreApplication::translate("MainWindow", "\320\222\321\226\320\264\320\272\321\200\320\270\321\202\320\270 Word", nullptr));
        action_Access->setText(QCoreApplication::translate("MainWindow", "\320\222\321\226\320\264\320\272\321\200\320\270\321\202\320\270 Access", nullptr));
        actionFullInfo->setText(QCoreApplication::translate("MainWindow", "\320\237\320\276\320\262\320\275\320\260 \321\226\320\275\321\204\320\276\321\200\320\274\320\260\321\206\321\226\321\217 \320\277\321\200\320\276 \321\200\320\276\320\267\321\200\320\276\320\261\320\275\320\270\320\272\320\260", nullptr));
        actionInform->setText(QCoreApplication::translate("MainWindow", "\320\206\320\275\321\204\320\276\321\200\320\274\320\260\321\206\321\226\321\217 \320\277\321\200\320\276 \320\272\320\265\321\200\321\226\320\262\320\275\320\270\320\272\320\260 ", nullptr));
        actionClose->setText(QCoreApplication::translate("MainWindow", "\320\222\320\270\321\205\321\226\320\264", nullptr));
        actionInfoApp->setText(QCoreApplication::translate("MainWindow", "\320\206\320\275\321\204\320\276\321\200\320\274\320\260\321\206\321\226\321\217 \320\277\321\200\320\276 \320\264\320\276\320\264\320\260\321\202\320\276\320\272", nullptr));
        OpenProvidnyk->setText(QCoreApplication::translate("MainWindow", "\320\222\321\226\320\264\320\272\321\200\320\270\321\202\320\270 \320\277\321\200\320\276\320\262\321\226\320\264\320\275\320\270\320\272", nullptr));
#if QT_CONFIG(shortcut)
        OpenProvidnyk->setShortcut(QCoreApplication::translate("MainWindow", "Ctrl+A", nullptr));
#endif // QT_CONFIG(shortcut)
        Average_2->setTitle(QCoreApplication::translate("MainWindow", "\320\241\320\265\321\200\320\265\320\264\320\275\321\217 \321\206\321\226\320\275\321\217", nullptr));
        Average->setText(QCoreApplication::translate("MainWindow", "\320\241\320\265\321\200\320\265\320\264\320\275\321\217 \321\206\321\226\320\275\320\260", nullptr));
        groupBox_2->setTitle(QCoreApplication::translate("MainWindow", "\320\224\320\276\320\264\320\260\321\202\320\272\320\276\320\262\321\226 \320\276\320\277\321\206\321\226\321\227", nullptr));
        OpenExcel->setText(QCoreApplication::translate("MainWindow", "\320\222\321\226\320\264\320\272\321\200\320\270\321\202\320\270 Excel", nullptr));
        OpenWord->setText(QCoreApplication::translate("MainWindow", "\320\222\321\226\320\264\320\272\321\200\320\270\321\202\320\270 Word", nullptr));
        OpenAccess->setText(QCoreApplication::translate("MainWindow", "\320\222\321\226\320\264\320\272\321\200\320\270\321\202\320\270 Access", nullptr));
        Operations->setTitle(QCoreApplication::translate("MainWindow", "\320\236\320\277\320\265\321\200\320\260\321\206\321\226\321\227 \320\267 \321\202\320\260\320\261\320\273\320\270\321\206\320\265\321\216", nullptr));
        Add->setText(QCoreApplication::translate("MainWindow", "\320\224\320\276\320\264\320\260\321\202\320\270", nullptr));
        Delete->setText(QCoreApplication::translate("MainWindow", "\320\222\320\270\320\264\320\260\320\273\320\270\321\202\320\270", nullptr));
        Download->setText(QCoreApplication::translate("MainWindow", "\320\237\320\276\320\272\320\260\320\267\320\260\321\202\320\270 \321\202\320\260\320\261\320\273\320\270\321\206\321\216", nullptr));
        Sort->setText(QCoreApplication::translate("MainWindow", "\320\241\320\276\321\200\321\202\321\203\320\262\320\260\321\202\320\270 \320\267\320\260 \321\206\321\226\320\275\320\276\321\216", nullptr));
        groupBox->setTitle(QCoreApplication::translate("MainWindow", "\320\237\320\276\321\210\321\203\320\272", nullptr));
        Search->setText(QCoreApplication::translate("MainWindow", "\320\237\320\276\321\210\321\203\320\272", nullptr));
        menuoptions->setTitle(QCoreApplication::translate("MainWindow", "\320\236\320\277\321\206\321\226\321\227", nullptr));
        menu->setTitle(QCoreApplication::translate("MainWindow", "\320\224\320\276\320\264\320\260\321\202\320\272\320\276\320\262\321\226 \320\276\320\277\321\206\321\226\321\227", nullptr));
        menu_2->setTitle(QCoreApplication::translate("MainWindow", "\320\206\320\275\321\204\320\276\321\200\320\274\320\260\321\206\321\226\321\217", nullptr));
        toolBar->setWindowTitle(QCoreApplication::translate("MainWindow", "toolBar", nullptr));
    } // retranslateUi

};

namespace Ui {
    class MainWindow: public Ui_MainWindow {};
} // namespace Ui

QT_END_NAMESPACE

#endif // UI_MAINWINDOW_H
