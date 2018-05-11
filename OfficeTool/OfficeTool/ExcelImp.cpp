#include "stdafx.h"
#include "ExcelImp.h"
#include<comdef.h>

CExcelImp::CExcelImp(void)
{
	CoInitialize(NULL);
	m_App.CreateDispatch(_T("Excel.Application"));
	m_strPath		= _T("");
}

CExcelImp::~CExcelImp(void)
{
	Close();
	CoUninitialize();
}

bool CExcelImp::Open(bool bCreate, CString strPath)
{
	m_strPath	= strPath;
	if ( bCreate )
	{
		m_App.SetScreenUpdating(FALSE);
		m_App.SetDisplayAlerts(FALSE);
		m_WorkBooks.AttachDispatch(m_App.GetWorkbooks(),TRUE);
		m_WorkBooks.Add(COleVariant((long)DISP_E_PARAMNOTFOUND, VT_ERROR));
		m_WorkSheets.AttachDispatch(m_App.GetWorksheets(),TRUE);
		int nCount = m_WorkSheets.GetCount();
		return true;
	}
	else
		return false;
}

void CExcelImp::Close()
{	
	int i = 1;
	int nCount = m_WorkSheets.GetCount();
	for ( i = 1; i <= nCount; i++ )
	{
		MS_EXCEL_2007::_Workbook mWorkBook = m_WorkSheets.GetItem(COleVariant(long(i)));
		mWorkBook.ReleaseDispatch();
	}
	m_WorkSheets.ReleaseDispatch();

	nCount = m_WorkBooks.GetCount();
	for ( i = 1; i <= nCount; i++ )
	{
		MS_EXCEL_2007::_Workbook mWorkBook = m_WorkBooks.GetItem(COleVariant(long(i)));
		mWorkBook.ReleaseDispatch();
	}
	m_WorkBooks.Close();
	m_WorkBooks.ReleaseDispatch();

	m_App.ReleaseDispatch();
	m_App.Quit();
}

bool CExcelImp::Save()
{
	int nCount = m_WorkBooks.GetCount();
	for ( int i = 1; i <= nCount; i++ )
	{
		MS_EXCEL_2007::_Workbook mWorkBook = m_WorkBooks.GetItem(COleVariant(long(i)));
		//mWorkBook.Save();
		COleVariant mOpt = COleVariant((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
		mWorkBook.SaveAs(COleVariant(m_strPath), mOpt,	mOpt, mOpt,	mOpt, mOpt,	0, mOpt, mOpt, mOpt, mOpt, mOpt);
	}
	return true;
}

void CExcelImp::SetRangeText(int nSheetId, int nRow, int nCol, CString strText)
{

}

CString	CExcelImp::GetRangeText(int nSheetId, int nRow, int nCol)
{
	return _T("");
}

bool CExcelImp::AddWorkSheet(CString strName)
{
	MS_EXCEL_2007::_Worksheet mWorkSheet = m_WorkSheets.Add(vtMissing,_variant_t(m_WorkSheets.GetItem(COleVariant(m_WorkSheets.GetCount()))),COleVariant(long(1)),vtMissing);
	mWorkSheet.SetName(strName);
	return m_WorkSheets.Add(vtMissing,vtMissing,_variant_t((long)1),vtMissing);;
}

long CExcelImp::GetWorkSheetCount()
{
	return m_WorkSheets.GetCount();
}