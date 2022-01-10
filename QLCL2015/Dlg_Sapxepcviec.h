#pragma once
#include "stdafx.h"

class Dlg_Sapxepcviec : public CDialog
{
	DECLARE_DYNAMIC(Dlg_Sapxepcviec)

public:
	Dlg_Sapxepcviec(CWnd* pParent = NULL);   // standard constructor
	virtual ~Dlg_Sapxepcviec();

	// Dialog Data
	enum { IDD = DLG_SORTCV };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual BOOL OnInitDialog();
	DECLARE_MESSAGE_MAP()

public:
	CButton m_btn, m_chk;
	CComboBox m_cbb1, m_cbb2, m_cbb3;
	_WorksheetPtr pSheet;
	RangePtr pRange, PRS;
	_bstr_t _bsName, _bsCNM;
	CString sNameSh;

	int iNgay[10];
	int _iCot[30];
	int Msrt[1000], Msgd[1000];
	int iEnd, ibd, ikt;

public:
	afx_msg void xl_xdcot_DMCV(_WorksheetPtr pSh);
	afx_msg void f_Load_combobx();
	afx_msg void f_xac_dinh_HM();
	afx_msg void f_xacdinh_gdoan(int num1, int num2);
	afx_msg void f_Sap_xep(int iNum1, int iNum2, int iCSort);

	afx_msg void OnBnClickedB2();
	afx_msg void OnBnClickedB1();
	afx_msg void OnCbnSelchangeCb2();
	afx_msg void OnCbnSelchangeCb3();
	afx_msg void OnCbnSelchangeCb1();
};
