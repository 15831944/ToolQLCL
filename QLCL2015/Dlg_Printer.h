#pragma once
#include "stdafx.h"

class Dlg_Printer : public CDialog
{

DECLARE_EASYSIZE
public:
	Dlg_Printer(CWnd* pParent = NULL);	// standard constructor
	
	enum { IDD = DLG_PRINTER};

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);  

protected:
	virtual BOOL OnInitDialog();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnSizing(UINT fwSide, LPRECT pRect);
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);

public:
	CListCtrl m_List;
	CButton m_Check;

	_bstr_t sh1;
	_WorksheetPtr pSh1;

public:
	afx_msg void f_GetAllSheet();
	afx_msg void f_StyleList();
	afx_msg void f_SetCheck_List(BOOLEAN f);
	afx_msg void OnBnClickedCk1();
	afx_msg void OnBnClickedB1();
	afx_msg void OnBnClickedB3();
	afx_msg void OnBnClickedB2();
	afx_msg void OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult);
};

