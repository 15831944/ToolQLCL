#pragma once
#include "stdafx.h"

// Dlg_Cpys dialog
class Dlg_Cpys : public CDialog
{
public:
	Dlg_Cpys();
	
	enum { IDD = DLG_CPYSHET};

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);  

protected:
	virtual BOOL OnInitDialog();
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);

public:
	CListCtrl m_list55;
	CButton m_ck1;

	afx_msg void f_Style_list();
	afx_msg void f_Load_dsanh();
	afx_msg void f_SetCheck(int n);
	afx_msg void OnBnClickedB1();
	afx_msg void OnBnClickedB2();
	afx_msg void OnBnClickedCk1();
};

