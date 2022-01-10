#pragma once
#include "stdafx.h"

// Dlg_sapxepIn dialog
class Dlg_sapxepIn : public CDialog
{
	DECLARE_DYNAMIC(Dlg_sapxepIn)

public:
	Dlg_sapxepIn(CWnd* pParent = NULL);   // standard constructor
	virtual ~Dlg_sapxepIn();

// Dialog Data
	enum { IDD = DLG_SAPXEPIN };

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()

public:
	CListCtrl m_list1_VL,m_list2_CV,m_list3_GD;
	int nSheet;

	afx_msg void f_GetAllSheet();
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	afx_msg void OnBnClickedB1();
	afx_msg void OnBnClickedB4();
	afx_msg void OnBnClickedB7();
	afx_msg void OnBnClickedB2();
	afx_msg void OnBnClickedB5();
	afx_msg void OnBnClickedB8();
	afx_msg void OnBnClickedB3();
	afx_msg void OnBnClickedB6();
	afx_msg void OnBnClickedB9();
};
