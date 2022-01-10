#pragma once
#include "stdafx.h"

// Dlg_open_progress dialog

class Dlg_open_progress : public CDialogEx
{
	DECLARE_DYNAMIC(Dlg_open_progress)

public:
	Dlg_open_progress(CWnd* pParent = nullptr);   // standard constructor
	virtual ~Dlg_open_progress();

	enum { IDD = IDD_DIALOG_PROGRESS };

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	DECLARE_MESSAGE_MAP()

public:
	afx_msg void OnTimer(UINT_PTR nIDEvent);

	int demTimer;

};
