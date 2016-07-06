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

	// ��ȡ�Ѱ�װ����������(������odbcinst.h��)
	if (!SQLGetInstalledDrivers(szBuf, cbBufMax, &cbBufOut))
		return "";

	// �����Ѱ�װ�������Ƿ���Excel...
	do
	{
		if (strstr(pszBuf, "Excel") != 0)
		{
			//���� !
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
	wstring sFile = _T("J:\\demo.xls"); 			// ������ȡ��Excel�ļ���

	// �����Ƿ�װ��Excel���� "Microsoft Excel Driver (*.xls)" 
	string temp = GetExcelDriver();
	sDriver = _T("Microsoft Excel Driver (*.xls)");//temp.c_str();
	if (sDriver.empty())
	{
		// û�з���Excel����
		QMessageBox::information(this, "xls","û�а�װExcel����!");
		return;
	}

	// �������д�ȡ���ַ���
	
	//sDsn.Format("ODBC;DRIVER={%s};DSN='''';DBQ=%s", sDriver, sFile);
	sDsn = _T("ODBC;DRIVER={") + sDriver + _T("};DSN='''';DBQ=") + sFile;

	TRY
	{
		// �����ݿ�(��Excel�ļ�)
		database.Open(NULL, false, false, sDsn.c_str());

		CRecordset recset(&database);

		// ���ö�ȡ�Ĳ�ѯ���.
		sSql = _T("SELECT Name, Age ")       
			_T("FROM demo ")                 
			_T("ORDER BY Name ");

		// ִ�в�ѯ���
		recset.Open(CRecordset::forwardOnly, sSql.c_str(), CRecordset::readOnly);

		// ��ȡ��ѯ���
		while (!recset.IsEOF())
		{
			//��ȡExcel�ڲ���ֵ
			recset.GetFieldValue(_T("Name"), sItem1);
			recset.GetFieldValue(_T("Age"), sItem2);

			// �Ƶ���һ��
			recset.MoveNext();
		}

		// �ر����ݿ�
		database.Close();

	}
	CATCH(CDBException, e)
	{
		// ���ݿ���������쳣ʱ...
		AfxMessageBox("���ݿ����: " + e->m_strError);
	}
	END_CATCH;
}

void Helloworld::WriteToExcel()
{
	CDatabase database;
	wstring sDriver = _T("MICROSOFT EXCEL DRIVER (*.XLS)"); // Excel��װ����
	wstring sExcelFile = _T("J:\\demo.xls");                // Ҫ������Excel�ļ�
	wstring sSql;

	TRY
	{
		// �������д�ȡ���ַ���
		//sSql.Format("DRIVER={%s};DSN='''';FIRSTROWHASNAMES=1;READONLY=FALSE;CREATE_DB=\"%s\";DBQ=%s",sDriver, sExcelFile, sExcelFile);

		sSql = _T("DRIVER={") + sDriver + _T("};DSN='''';FIRSTROWHASNAMES=1;READONLY=FALSE;CREATE_DB=\"") + sExcelFile + _T("\";DBQ=") + sExcelFile;

		// �������ݿ� (��Excel����ļ�)
		if( database.OpenEx(sSql.c_str(),CDatabase::noOdbcDialog) )
		{
			// ������ṹ(����������)
			//sSql = _T("CREATE TABLE  demo (Name TEXT,Age NUMBER)");  
			//sSql = _T("CREATE TABLE  demo IF NOT EXISTS");
			//database.ExecuteSQL(sSql.c_str());

			// ������ֵ
			sSql = _T("INSERT INTO demo (Name,Age) VALUES ('�쾰��AASDASDASDASDASDASDASDD',26)");
			database.ExecuteSQL(sSql.c_str());

			sSql = _T("INSERT INTO demo (Name,Age) VALUES ('��־��',22)");
			database.ExecuteSQL(sSql.c_str());

			sSql = _T("INSERT INTO demo (Name,Age) VALUES ('����',27)");
			database.ExecuteSQL(sSql.c_str());
		}      
		{
			// ������ṹ(����������)
			sSql = _T("CREATE TABLE DC (Districts TEXT,constituency TEXT)");
			database.ExecuteSQL(sSql.c_str());

			// ������ֵ
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
	// �ر����ݿ�
	database.Close();
}