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

bool CExcel::AddWorkSheet(CString strName)
{
	return m_pImp->AddWorkSheet(strName);
}

long CExcel::GetWorkSheetCount()
{
	return m_pImp->GetWorkSheetCount();
}

bool CExcel::SetWorkSheetName(long nSheetId, CString strName)
{
	return m_pImp->SetWorkSheetName(nSheetId, strName);
}

void CExcel::SetRowAndCol(int nRow, int nCol)
{
	m_pImp->SetRowAndCol(nRow, nCol);
}

void CExcel::GetRowAndCol(int& nRow, int& nCol)
{
	m_pImp->GetRowAndCol(nRow, nCol);
}