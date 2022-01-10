#pragma once
#include "stdafx.h"

// Dlg_Nky_Thoitiet dialog
class Dlg_Nky_Thoitiet : public CDialogEx
{
	DECLARE_DYNAMIC(Dlg_Nky_Thoitiet)

public:
	Dlg_Nky_Thoitiet(CWnd* pParent = nullptr);   // standard constructor
	virtual ~Dlg_Nky_Thoitiet();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = DLG_NKY_THOITIET };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButton3();
};
