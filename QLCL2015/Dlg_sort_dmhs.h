#pragma once
#include "stdafx.h"

// Dlg_sort_dmhs dialog
class Dlg_sort_dmhs : public CDialog
{
public:
	Dlg_sort_dmhs();
	
	enum { IDD = DLG_SRTDMHS};

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);  

protected:
	virtual BOOL OnInitDialog();
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);

public:
	CButton Ckhb1,Crd1,Crd2;
	CComboBox Cbb1;

	int iCheckTenTT;

public:
	afx_msg void OnBnClickedButton2();
	afx_msg void OnBnClickedB1();
	afx_msg void OnBnClickedCk1();
	afx_msg void OnBnClickedCk2();
	afx_msg void OnCbnSelchangeCb1();
};


