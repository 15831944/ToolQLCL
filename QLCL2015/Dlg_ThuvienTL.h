#pragma once
#include "stdafx.h"
#include "Path.h"

class Dlg_ThuvienTL : public CDialog
{
	DECLARE_EASYSIZE
public:
	Dlg_ThuvienTL(CWnd* pParent = NULL);	// standard constructor
	
	enum { IDD = DLG_TVTLIEU};

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);  

protected:
	virtual BOOL OnInitDialog();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnSizing(UINT fwSide, LPRECT pRect);
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	CFileFinder	_finder;

public:
	CListCtrlEx list1;
	CEdit txtKey;
	CStatic sttic;

	int slTL;
	CString szFolder,sList[2000];

public:
	afx_msg void f_Add_Csdl(CString m_sBaseFolder, CString m_sFileMask);
	afx_msg void f_Style_tieuchuan();
	afx_msg void f_delete_list();
	afx_msg void f_Load_list_tieuchuan();

	afx_msg void OnBnClickedB4();
	afx_msg void OnBnClickedB2();
	afx_msg void OnBnClickedB3();
	afx_msg void OnBnClickedB1();
	afx_msg void OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult);
};
