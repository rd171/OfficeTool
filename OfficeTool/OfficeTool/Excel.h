#pragma once

class CExcelImp;
class CExcel
{
public:
	CExcel(void);
	~CExcel(void);

	bool	Open(bool bCreate, CString strPath);
	void	Close();
	bool	Save();

	void	SetRangeText(int nSheetId, int nRow, int nCol, CString strText);
	CString	GetRangeText(int nSheetId, int nRow, int nCol);

	bool	AddWorkSheet(CString strName, int nRow, int nCol);
	long	GetWorkSheetCount();
	bool	SetWorkSheetName(long nSheetId, CString strName);

private:
	CExcelImp*	m_pImp;
};

