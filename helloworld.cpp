#include "helloworld.h"
#include <QMessageBox>

//#undef UNICODE
#define SQL_NOUNICODEMAP
#include <afxdb.h>
#include <odbcinst.h>

//#define  _T(a)  a
//#include "CSpreadSheet.h"

Helloworld::Helloworld(QWidget *parent, Qt::WFlags flags)
	: QMainWindow(parent, flags)
{
	ui.setupUi(this);

	QObject::connect(this, SIGNAL(signalsTest(int)), this, SLOT(slot2Test(int)));
}

Helloworld::~Helloworld()
{

}

void Helloworld::slot1TestClik()
{
	emit signalsTest(100);
	cout<<"test"<<endl;
}

void Helloworld::slot2Test(int test)
{
	cout<<test<<endl;
}

void Helloworld::slot1OpenXls()
{
	QString  strTite = "test";
	QString  strinfo = "test open";
	string str = strTite.toStdString();

	QMessageBox::information(this,strTite, strinfo);

	str = GetExcelDriver();
	QMessageBox::information(this, "SQL", str.c_str());
}

void Helloworld::slot1ReadXls()
{
	ReadFromExcel();
}
void Helloworld::slot1WriteXls()
{
	WriteToExcel();
}

string Helloworld::GetExcelDriver()
{
	char szBuf[2001];
	unsigned short cbBufMax = 2000;
	unsigned short cbBufOut;
	char *pszBuf = szBuf;
	string sDriver = "";

	// 获取已安装驱动的名称(涵数在odbcinst.h里)
	if (!SQLGetInstalledDrivers(szBuf, cbBufMax, &cbBufOut))
		return "";

	// 检索已安装的驱动是否有Excel...
	do
	{
		if (strstr(pszBuf, "Excel") != 0)
		{
			//发现 !
			sDriver = pszBuf;
			break;
		}
		pszBuf = strchr(pszBuf, '\0') + 1;
	}
	while (pszBuf[1] != '\0');

	return sDriver;
}

void Helloworld::ReadFromExcel()
{
	CDatabase database;
	wstring sSql;
	CString sItem1, sItem2;
	wstring sDriver;
	wstring sDsn;
	wstring sFile = _T("J:\\demo.xls"); 			// 将被读取的Excel文件名

	// 检索是否安装有Excel驱动 "Microsoft Excel Driver (*.xls)" 
	string temp = GetExcelDriver();
	sDriver = _T("Microsoft Excel Driver (*.xls)");//temp.c_str();
	if (sDriver.empty())
	{
		// 没有发现Excel驱动
		QMessageBox::information(this, "xls","没有安装Excel驱动!");
		return;
	}

	// 创建进行存取的字符串
	
	//sDsn.Format("ODBC;DRIVER={%s};DSN='''';DBQ=%s", sDriver, sFile);
	sDsn = _T("ODBC;DRIVER={") + sDriver + _T("};DSN='''';DBQ=") + sFile;

	TRY
	{
		// 打开数据库(既Excel文件)
		database.Open(NULL, false, false, sDsn.c_str());

		CRecordset recset(&database);

		// 设置读取的查询语句.
		sSql = _T("SELECT Name, Age ")       
			_T("FROM demo ")                 
			_T("ORDER BY Name ");

		// 执行查询语句
		recset.Open(CRecordset::forwardOnly, sSql.c_str(), CRecordset::readOnly);

		// 获取查询结果
		while (!recset.IsEOF())
		{
			//读取Excel内部数值
			recset.GetFieldValue(_T("Name"), sItem1);
			recset.GetFieldValue(_T("Age"), sItem2);

			// 移到下一行
			recset.MoveNext();
		}

		// 关闭数据库
		database.Close();

	}
	CATCH(CDBException, e)
	{
		// 数据库操作产生异常时...
		AfxMessageBox("数据库错误: " + e->m_strError);
	}
	END_CATCH;
}

void Helloworld::WriteToExcel()
{
	CDatabase database;
	wstring sDriver = _T("MICROSOFT EXCEL DRIVER (*.XLS)"); // Excel安装驱动
	wstring sExcelFile = _T("J:\\demo.xls");                // 要建立的Excel文件
	wstring sSql;

	TRY
	{
		// 创建进行存取的字符串
		//sSql.Format("DRIVER={%s};DSN='''';FIRSTROWHASNAMES=1;READONLY=FALSE;CREATE_DB=\"%s\";DBQ=%s",sDriver, sExcelFile, sExcelFile);

		sSql = _T("DRIVER={") + sDriver + _T("};DSN='''';FIRSTROWHASNAMES=1;READONLY=FALSE;CREATE_DB=\"") + sExcelFile + _T("\";DBQ=") + sExcelFile;

		// 创建数据库 (既Excel表格文件)
		if( database.OpenEx(sSql.c_str(),CDatabase::noOdbcDialog) )
		{
			// 创建表结构(姓名、年龄)
			//sSql = _T("CREATE TABLE  demo (Name TEXT,Age NUMBER)");  
			//sSql = _T("CREATE TABLE  demo IF NOT EXISTS");
			//database.ExecuteSQL(sSql.c_str());

			// 插入数值
			sSql = _T("INSERT INTO demo (Name,Age) VALUES ('徐景周AASDASDASDASDASDASDASDD',26)");
			database.ExecuteSQL(sSql.c_str());

			sSql = _T("INSERT INTO demo (Name,Age) VALUES ('徐志慧',22)");
			database.ExecuteSQL(sSql.c_str());

			sSql = _T("INSERT INTO demo (Name,Age) VALUES ('郭徽',27)");
			database.ExecuteSQL(sSql.c_str());
		}      
		{
			// 创建表结构(姓名、年龄)
			sSql = _T("CREATE TABLE DC (Districts TEXT,constituency TEXT)");
			database.ExecuteSQL(sSql.c_str());

			// 插入数值
			sSql = _T("INSERT INTO DC (Districts,constituency) VALUES ('hunan','hengy')");
			database.ExecuteSQL(sSql.c_str());

			sSql = _T("INSERT INTO DC (Districts,constituency) VALUES ('jianxi','nanchang')");
			database.ExecuteSQL(sSql.c_str());

			sSql = _T("INSERT INTO DC (Districts,constituency) VALUES ('zhejian','hanzhou')");
			database.ExecuteSQL(sSql.c_str());
		}

	}
	CATCH_ALL(e)
	{
		//TRACE1("Excel ERROR: %s", e->[CDBException].m_strError);
		CString  strError =  ((CDBException*)e)->m_strError;
		//QMessageBox::Information(this,);
		AfxMessageBox(strError);
	}
	END_CATCH_ALL;
	// 关闭数据库
	database.Close();
}