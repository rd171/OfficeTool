#pragma once
#include "excel_2007.h"
using namespace MS_EXCEL_2007;


class CSheetData
{
public:
	CSheetData(long nRow, long nCol);
	~CSheetData();

	bool	SetRangeText(long nRow, long nCol, CString strText);
	CString	GetRangeText(long nRow, long nCol);
	CString GetAllText();

private:
	int				m_nRow;		// Sheet行数(索引从1开始)
	int				m_nCol;		// Sheet列数(索引从1开始)
	CStringArray*	m_pRowText;	// 行数据
};

class CExcelImp
{
public:
	enum RangeStyle
	{
		RS_NORMAL,		// 常规
		RS_NUMBER,		// 数值(两位小数)
		RS_STRING,		// 文本
		RS_DATE,		// 日期(yyyy/m/d)
		RS_TIME,		// 日期(hh:mm:ss)
	};

public:
	CExcelImp(void);
	~CExcelImp(void);

	bool	Open(bool bCreate, CString strPath);
	void	Close();
	bool	Save();

	bool	AddWorkSheet(CString strName, long nRow, long nCol);
	long	GetWorkSheetCount();
	bool	SetWorkSheetName(long nSheetId, CString strName);

	BOOL	SetRangeText(long nSheetId, long nRow, long nCol, CString strText);
	CString	GetRangeText(long nSheetId, long nRow, long nCol);

	void	SetRangeStyle(long nSheetId, long nRow1, long nCol1, long nRow2, long nCol2, RangeStyle style);

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

