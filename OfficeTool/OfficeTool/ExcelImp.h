#pragma once
#include "excel_2007.h"
using namespace MS_EXCEL_2007;


class CSheetData
{
public:
	CSheetData(int nRow, int nCol);
	~CSheetData();

	bool	SetRangeText(int nRow, int nCol, CString strText);
	CString	GetRangeText(int nRow, int nCol);
	CString GetAllText();

private:
	int				m_nRow;		// Sheet行数(索引从1开始)
	int				m_nCol;		// Sheet列数(索引从1开始)
	CStringArray*	m_pRowText;	// 行数据
};

class CExcelImp
{
public:
	CExcelImp(void);
	~CExcelImp(void);

	bool	Open(bool bCreate, CString strPath);
	void	Close();
	bool	Save();

	BOOL	SetRangeText(int nSheetId, int nRow, int nCol, CString strText);
	CString	GetRangeText(int nSheetId, int nRow, int nCol);

	bool	AddWorkSheet(CString strName, int nRow, int nCol);
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

	CPtrList					m_listSheetData;
};

