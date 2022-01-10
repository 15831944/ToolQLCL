#pragma once
#include "stdafx.h"

// Dlg_Nky_BangNLTB dialog

class Dlg_Nky_BangNLTB : public CDialogEx
{
	DECLARE_DYNAMIC(Dlg_Nky_BangNLTB)

public:
	Dlg_Nky_BangNLTB(CWnd* pParent = nullptr);   // standard constructor
	virtual ~Dlg_Nky_BangNLTB();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = DLG_NKY_BANGNLTB };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButton3();
};
