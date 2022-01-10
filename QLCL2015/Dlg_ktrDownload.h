#pragma once
#include "stdafx.h"

// Dlg_ktrDownload dialog
class Dlg_ktrDownload : public CDialogEx
{
	DECLARE_DYNAMIC(Dlg_ktrDownload)

public:
	Dlg_ktrDownload(CWnd* pParent = NULL);   // standard constructor
	virtual ~Dlg_ktrDownload();

// Dialog Data
	enum { IDD = DLG_CHECKDOWN };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual BOOL OnInitDialog();
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);

public:
	CStatic m_stt1;
	CEdit m_txt2;

	CString sDapan;
	int iDem, iSai;

	afx_msg void _SetRandomQuit();
	afx_msg void _CheckKetqua();
	afx_msg void OnBnClickedButton1();
};
