#pragma once
#include "stdafx.h"

// Dlg_QCLMBetong dialog
class Dlg_QCLMBetong : public CDialog
{
	DECLARE_DYNAMIC(Dlg_QCLMBetong)

public:
	Dlg_QCLMBetong(CWnd* pParent = NULL);   // standard constructor
	virtual ~Dlg_QCLMBetong();

// Dialog Data
	enum { IDD = IDD_DLG_MTN };

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);

public:
	CStatic m_stt[6];
	CEdit m_txt[30];

	afx_msg void OnBnClickedMtnOk();
	afx_msg void OnBnClickedMtnCancel();
	afx_msg void f_Tack_chuoi(CString sKey);

	afx_msg void OnBnClickedMtnOk2();
	afx_msg void OnBnClickedMtnOk3();
	afx_msg void OnBnClickedMtnOk4();
	afx_msg void OnBnClickedMtnOk5();

	afx_msg void OnBnClickedAutoCheck();
};
