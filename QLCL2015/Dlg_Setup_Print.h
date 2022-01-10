#pragma once
#include "stdafx.h"

// Dlg_Setup_Print dialog
class Dlg_Setup_Print : public CDialog
{
public:
	Dlg_Setup_Print();
	
	enum { IDD = DLG_SETTUP_PRT};

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);  

protected:
	virtual BOOL OnInitDialog();
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);

public:
	CEdit txt[6];
	CComboBox fcbb[15];
	CButton fchk[15];
	CToolTipCtrl m_ToolTips;

	CString msg;
	LPTSTR lp;
	int iLen;
	int _ick[3];

	CString fchon[13];
	afx_msg void f_thietlap_macdinh();
	afx_msg void f_load_combobox();
	afx_msg int f_setCombobox(int n);
	afx_msg CString f_Gettext(int n);
	afx_msg void f_Load_ToolTip();
	afx_msg void f_Enable_private(int nType);

	afx_msg void OnBnClickedB1();
	afx_msg void OnBnClickedB2();
	afx_msg void OnBnClickedCk6();
	afx_msg void OnBnClickedCk7();
	afx_msg void OnBnClickedCk8();
	afx_msg void OnBnClickedCk9();
	afx_msg void OnBnClickedCk10();
	afx_msg void OnEnKillfocusT3();
	afx_msg void OnBnClickedCk11();
	afx_msg void OnEnKillfocusT4();
	afx_msg void OnBnClickedCk12();
	afx_msg void OnBnClickedCk1();
	afx_msg void OnBnClickedB3();
	afx_msg void OnBnClickedCk13();
	afx_msg void OnCbnSelchangeC12();
};
