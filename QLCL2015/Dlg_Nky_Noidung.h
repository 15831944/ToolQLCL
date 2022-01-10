#pragma once
#include "stdafx.h"

// Dlg_Nky_Noidung dialog

class Dlg_Nky_Noidung : public CDialogEx
{
	DECLARE_DYNAMIC(Dlg_Nky_Noidung)

public:
	Dlg_Nky_Noidung(CWnd* pParent = nullptr);   // standard constructor
	virtual ~Dlg_Nky_Noidung();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = DLG_NKY_NOIDUNG };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButton3();
};
