#pragma once

class CExcelImp;
class CExcel
{
public:
	enum RangeStyle
	{
		RS_NORMAL,		// ����
		RS_NUMBER,		// ��ֵ(��λС��)
		RS_STRING,		// �ı�
		RS_DATE,		// ����(yyyy/m/d)
		RS_TIME,		// ����(hh:mm:ss)
	};

public:
	CExcel(void);
	~CExcel(void);

	bool	Open(bool bCreate, CString strPath);
	void	Close();
	bool	Save();

	bool	AddWorkSheet(CString strName, int nRow, int nCol);
	long	GetWorkSheetCount();
	bool	SetWorkSheetName(long nSheetId, CString strName);

	void	SetRangeText(int nSheetId, int nRow, int nCol, CString strText);
	CString	GetRangeText(int nSheetId, int nRow, int nCol);

	void	SetRangeStyle(int nSheetId, int nRow1, int nCol1, int nRow2, int nCol2, RangeStyle style);

private:
	CExcelImp*	m_pImp;
};

