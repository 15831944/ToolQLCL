#pragma once
#include "stdafx.h"

// Dlg_Create_Tieude dialog
class Dlg_Create_Tieude : public CDialog
{
	DECLARE_EASYSIZE
public:
	Dlg_Create_Tieude(CWnd* pParent = NULL);	// standard constructor

	enum { IDD = DLG_CREATE_TIEUDE };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);
	virtual BOOL OnInitDialog();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnSizing(UINT fwSide, LPRECT pRect);
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnContextMenu(CWnd* /*pWnd*/, CPoint /*point*/);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);

public:
	afx_msg void CreateList();
	afx_msg void OnBnClickedBtnCancel();
	afx_msg void OnBnClickedBtnOk();

public:
	int sl;
	int sldong;
	int nItem, nSubItem;
	CString szTitle;
	CString szAfter, szBefore;

public:
	afx_msg void OnT40065();
	afx_msg void OnT40066();
	afx_msg void OnNMClickList(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMRClickList(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnEnSetfocusEdit();
	afx_msg void OnEnKillfocusEdit();
};
