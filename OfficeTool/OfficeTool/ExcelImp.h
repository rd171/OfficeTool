#pragma once
#include "excel_2007.h"
using namespace MS_EXCEL_2007;

class CExcelImp
{
public:
	CExcelImp(void);
	~CExcelImp(void);

	bool	Open(bool bCreate, CString strPath);
	void	Close();
	bool	Save();

	void	SetRangeText(int nSheetId, int nRow, int nCol, CString strText);
	CString	GetRangeText(int nSheetId, int nRow, int nCol);

	void	SetRowAndCol(int nRow, int nCol);
	void	GetRowAndCol(int& nRow, int& nCol);

	bool	AddWorkSheet(CString strName);
	long	GetWorkSheetCount();
	bool	SetWorkSheetName(long nSheetId, CString strName);

private:
	void		Paste(long nSheetId, CString& strText);
	CString		Cell(long nItem,long nCol);

private:
	MS_EXCEL_2007::_Application	m_App;
	MS_EXCEL_2007::Worksheets	m_WorkSheets;
	MS_EXCEL_2007::Workbooks	m_WorkBooks;
	CString						m_strPath;
	int							m_nRow;
	int							m_nCol;
	CStringArray*				m_pCol;
};

