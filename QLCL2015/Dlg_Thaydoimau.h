#pragma once
#include "stdafx.h"

class Dlg_Thaydoimau : public CDialogEx
{
	DECLARE_DYNAMIC(Dlg_Thaydoimau)

public:
	Dlg_Thaydoimau(CWnd* pParent = NULL);   // standard constructor
	virtual ~Dlg_Thaydoimau();

// Dialog Data
	enum { IDD = DLG_CHGMAU };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual BOOL OnInitDialog();
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);

public:
	afx_msg void OnBnClickedB1();
	afx_msg void OnBnClickedB2();
};
