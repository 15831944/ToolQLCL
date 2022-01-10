#pragma once
#include "stdafx.h"

// Dlg_Luu_DMCV dialog
class Dlg_Luu_DMCV : public CDialog
{
public:
	Dlg_Luu_DMCV();
	
	enum { IDD = DLG_LUU_DMCV};

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);  

protected:
	virtual BOOL OnInitDialog();
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);

public:
	CButton _nRad1,_nRad2;
	CComboBox _cbb0;
	int xde,_nChk0;
	int _ncot[10];
	int nvtri[1000];
	CConnection ObjConn;
	CString msg;
	_WorksheetPtr psDMCV;
	RangePtr prDMCV,PRS;
	_bstr_t shDMCV;

public:
	afx_msg void f_kiemtra_sheet();
	afx_msg void f_Load_giaidoan();
	afx_msg int f_Luu_1_danhmuc(int _vtri);

	afx_msg void OnBnClickedB1();
	afx_msg void OnBnClickedB2();
	afx_msg void OnBnClickedR1();
	afx_msg void OnBnClickedR2();
};
