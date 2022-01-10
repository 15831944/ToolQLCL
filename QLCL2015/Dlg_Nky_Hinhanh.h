#pragma once
#include "stdafx.h"

// Dlg_Nky_Hinhanh dialog

class Dlg_Nky_Hinhanh : public CDialogEx
{
	DECLARE_DYNAMIC(Dlg_Nky_Hinhanh)

public:
	Dlg_Nky_Hinhanh(CWnd* pParent = nullptr);   // standard constructor
	virtual ~Dlg_Nky_Hinhanh();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = DLG_NKY_HINHANH };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButton4();
};
