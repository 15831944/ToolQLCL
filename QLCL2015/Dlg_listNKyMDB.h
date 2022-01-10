#pragma once
#include "stdafx.h"

// Dlg_listNKyMDB dialog
class Dlg_listNKyMDB : public CDialogEx
{
	DECLARE_DYNAMIC(Dlg_listNKyMDB)

public:
	Dlg_listNKyMDB(CWnd* pParent = NULL);   // standard constructor
	virtual ~Dlg_listNKyMDB();

// Dialog Data
	enum { IDD = DLG_NKYMDB };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual BOOL OnInitDialog();
	DECLARE_MESSAGE_MAP()

	private:
		afx_msg BOOL PreTranslateMessage(MSG* pMsg);
		afx_msg void OnSysCommand(UINT nID, LPARAM lParam);

public:
	CListCtrl m_list;
	CFileFinder	_finder;

	afx_msg void f_Add_Csdl(CString m_sBaseFolder, CString m_sFileMask);
	afx_msg void OnBnClickedB1();
	afx_msg void OnBnClickedB2();
	afx_msg void OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult);
};
