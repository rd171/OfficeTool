#include "stdafx.h"
#include "Excel.h"
#include "ExcelImp.h"

CExcel::CExcel(void)
{
	m_pImp	= new CExcelImp;
}

CExcel::~CExcel(void)
{
	delete	m_pImp;
}

bool CExcel::Open(bool bCreate, CString strPath)
{
	return m_pImp->Open(bCreate, strPath);
}

void CExcel::Close()
{
	m_pImp->Close();
}

bool CExcel::Save()
{
	return m_pImp->Save();
}

void CExcel::SetRangeText(int nSheetId, int nRow, int nCol, CString strText)
{
	m_pImp->SetRangeText(nSheetId, nRow, nCol, strText);
}

CString	CExcel::GetRangeText(int nSheetId, int nRow, int nCol)
{
	return m_pImp->GetRangeText(nSheetId, nRow, nCol);
}

bool CExcel::AddWorkSheet(CString strName, int nRow, int nCol)
{
	return m_pImp->AddWorkSheet(strName, nRow, nCol);
}

long CExcel::GetWorkSheetCount()
{
	return m_pImp->GetWorkSheetCount();
}

bool CExcel::SetWorkSheetName(long nSheetId, CString strName)
{
	return m_pImp->SetWorkSheetName(nSheetId, strName);
}

void CExcel::SetRangeStyle(int nSheetId, int nRow1, int nCol1, int nRow2, int nCol2, RangeStyle style)
{
	m_pImp->SetRangeStyle(nSheetId, nRow1, nCol1, nRow2, nCol2, (CExcelImp::RangeStyle)style);
}