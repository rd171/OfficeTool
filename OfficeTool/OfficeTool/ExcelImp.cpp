#include "stdafx.h"
#include "ExcelImp.h"
#include<comdef.h>

CExcelImp::CExcelImp(void)
{
	CoInitialize(NULL);
	m_App.CreateDispatch(_T("Excel.Application"));
	m_strPath		= _T("");
	m_nRow			= 0;
	m_nCol			= 0;
	m_pCol			= NULL;
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
		m_App.SetSheetsInNewWorkbook(long(1));
		m_WorkBooks.AttachDispatch(m_App.GetWorkbooks(),TRUE);
		m_WorkBooks.Add(COleVariant((long)DISP_E_PARAMNOTFOUND, VT_ERROR));
		m_WorkSheets.AttachDispatch(m_App.GetWorksheets(),TRUE);
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
	CString strData;
	for ( int i = 0; i < m_nRow; i++ )
	{
		for ( int j = 0; j < m_nCol; j++ )
			strData += m_pCol[i].GetAt(j);
		strData	+= _T("\r\n");
	}
	Paste(1, strData);

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
	if ( nRow > m_nRow || nCol > m_nCol )
		return;
	m_pCol[nRow-1].SetAt(m_nCol-1, strText);
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

bool CExcelImp::SetWorkSheetName(long nSheetId, CString strName)
{
	if ( nSheetId < 1 || nSheetId > m_WorkSheets.GetCount() )
		return false;
	MS_EXCEL_2007::_Worksheet WorkSheet = m_WorkSheets.GetItem(_variant_t(nSheetId));
	WorkSheet.SetName(strName);
	return true;
}

void CExcelImp::SetRowAndCol(int nRow, int nCol)
{
	if ( NULL != m_pCol )
		return;
	m_nRow		= nRow;
	m_nCol		= nCol;
	m_pCol		= new CStringArray[m_nRow];
	for ( int i = 0; i < m_nRow; i++ )
	{
		for ( int j = 0; j < m_nCol; j++ )
		{
			CString strTemp =  _T("\t");
			m_pCol[i].Add(strTemp);
		}
	}
}

void CExcelImp::GetRowAndCol(int& nRow, int& nCol)
{

}

CString CExcelImp::Cell(long nItem,long nCol)
{
	CString strCell = _T("");
	nCol--;
	if ( nCol == 0 )
	{
		strCell.Format(_T("A%d"),nItem);
		return strCell;
	}
	TCHAR szData[2048];
	memset(szData,0,2048);
	int  nDataPos = -1;
	while ( nCol != 0 )
	{
		nDataPos++;
		szData[nDataPos] += (nCol % 26 + 65);
		nCol = nCol / 26;
	}
	int i = 0;
	for ( i = 1; i <= nDataPos; i++)
	{
		szData[i] -= 1;
	}
	char cTemp;
	for ( i = 0; i < (nDataPos + 1)/2; i++ )
	{
		cTemp = szData[i];
		szData[i] = szData[nDataPos - i];
		szData[nDataPos - i] = cTemp;
	}
	szData[nDataPos + 1] = 0;
	strCell.Format(_T("%s%d"),szData,nItem);
	return strCell;
}

void CExcelImp::Paste(long nSheetId, CString& strText)
{
	if (OpenClipboard(NULL))   //如果能打开剪贴板  
	{  
		::EmptyClipboard();  //清空剪贴板，使该窗口成为剪贴板的拥有者   
		HGLOBAL hClip;  
		hClip = ::GlobalAlloc(GMEM_MOVEABLE, (strText.GetLength() * 2) + 2); //判断要是文本数据，分配内存时多分配一个字符  
		TCHAR *pBuf;  
		pBuf = (TCHAR *)::GlobalLock(hClip);//锁定剪贴板  
		lstrcpy(pBuf, strText);//把CString转换  
		::GlobalUnlock(hClip);//解除锁定剪贴板  
		::SetClipboardData(CF_UNICODETEXT, hClip);//把文本数据发送到剪贴板  CF_UNICODETEXT为Unicode编码  
		::CloseClipboard();//关闭剪贴板  
	}  

	MS_EXCEL_2007::_Worksheet mWorkSheet = m_WorkSheets.GetItem(COleVariant(nSheetId));
	COleVariant var1, var2;

	VARIANT vDest;
	CString strCell = Cell(1,1);
	Range mRange = mWorkSheet.GetRange( COleVariant( strCell ),COleVariant( strCell ) );

	VARIANT vLink;
	vLink.vt	= VT_BOOL;
	vLink.boolVal = false;

	vDest.vt	= VT_DISPATCH;
	vDest.pdispVal= mWorkSheet.GetRange( COleVariant( strCell ),COleVariant( strCell ) );

	mWorkSheet.Paste(vDest, vLink);
}