#pragma once
#include "stdafx.h"

class Dlg_NameSheet : public CDialog
{
	DECLARE_EASYSIZE
public:
	Dlg_NameSheet(CWnd* pParent = NULL);	// standard constructor
	
	enum { IDD = DLG_NAMESHEET};

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
	CColorStatic m_huongdan;
	CButton m_btnKhoiphuc;

public:
	CConnection ObjConn;
	_WorksheetPtr pSheet;
	RangePtr pRange;
	CString sNSheet;
	CString sPathCF;
	CString smsg1,smsg2,smsg3;
	CString sBefore,sAfter;

	int totalName;
	CString sName[2000];
	CString sVtri[2000];

	afx_msg int _CheckName(CString snm);
	afx_msg void _GetAllName();
	afx_msg void _GetAllNameSheet();
	afx_msg void _StyleList();
	afx_msg void OnBnClickedB2();
	afx_msg void OnBnClickedB1();
	afx_msg void OnNMClickL1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnEnSetfocusE1();
	afx_msg void OnEnKillfocusE1();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
};
