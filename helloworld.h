#ifndef HELLOWORLD_H
#define HELLOWORLD_H

#include <QtGui/QMainWindow>
#include "ui_helloworld.h"
#include <iostream>
using namespace std;

class Helloworld : public QMainWindow
{
	Q_OBJECT

public:
	Helloworld(QWidget *parent = 0, Qt::WFlags flags = 0);
	~Helloworld();
signals:
	void signalsTest(int test);
public slots:
	void slot1TestClik();

	void slot2Test(int test);

	void slot1OpenXls();
	void slot1ReadXls();
	void slot1WriteXls();

private:
	Ui::HelloworldClass ui;

public:
	string GetExcelDriver();
	void ReadFromExcel();
	void WriteToExcel();
};

#endif // HELLOWORLD_H
