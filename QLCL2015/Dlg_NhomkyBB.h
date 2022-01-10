#pragma once
#include "stdafx.h"

// Dlg_NhomkyBB dialog
class Dlg_NhomkyBB : public CDialogEx
{
	DECLARE_DYNAMIC(Dlg_NhomkyBB)

public:
	Dlg_NhomkyBB(CWnd* pParent = NULL);   // standard constructor
	virtual ~Dlg_NhomkyBB();

// Dialog Data
	enum { IDD = DLG_NHOMKBB };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual BOOL OnInitDialog();
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);

public:
	CListCtrl m_list;

	afx_msg void _LoadDS();
	afx_msg void OnBnClickedB1();
	afx_msg void OnBnClickedB2();
	afx_msg void OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnLvnKeydownL1(NMHDR *pNMHDR, LRESULT *pResult);
};
