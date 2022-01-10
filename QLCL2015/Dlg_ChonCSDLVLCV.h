#pragma once
#include "stdafx.h"

class Dlg_ChonCSDLVLCV : public CDialogEx
{
	DECLARE_DYNAMIC(Dlg_ChonCSDLVLCV)

public:
	Dlg_ChonCSDLVLCV(CWnd* pParent = NULL);   // standard constructor
	virtual ~Dlg_ChonCSDLVLCV();

// Dialog Data
	enum { IDD = DLG_CHONDLCVVL };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual BOOL OnInitDialog();
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);

public:
	CListCtrl m_list;
	CFileFinder	_finder;

	afx_msg void _LoadDS();
	afx_msg void f_Add_Csdl(CString m_sBaseFolder, CString m_sFileMask);

	afx_msg void OnBnClickedB2();
	afx_msg void OnBnClickedB1();

	afx_msg void OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnLvnKeydownL1(NMHDR *pNMHDR, LRESULT *pResult);
};
