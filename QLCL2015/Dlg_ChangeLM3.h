#pragma once
#include "stdafx.h"

// Dlg_ChangeLM3 dialog
class Dlg_ChangeLM3 : public CDialog
{
	DECLARE_DYNAMIC(Dlg_ChangeLM3)

public:
	Dlg_ChangeLM3(CWnd* pParent = NULL);   // standard constructor
	virtual ~Dlg_ChangeLM3();

// Dialog Data
	enum { IDD = DLG_CGLM3 };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual BOOL OnInitDialog();
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);

public:
	CEdit m_t1,m_t2;
	CConnection ObjConn;

	afx_msg void OnBnClickedButton7();
	afx_msg void OnBnClickedButton8();
	afx_msg void OnBnClickedNex1();
	afx_msg void OnBnClickedPrv1();
};
