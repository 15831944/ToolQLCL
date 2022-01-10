#pragma once
#include "stdafx.h"

// Dlg_Chonngay_paste dialog
class Dlg_Chonngay_paste : public CDialog
{
	DECLARE_DYNAMIC(Dlg_Chonngay_paste)

public:
	Dlg_Chonngay_paste(CWnd* pParent = NULL);   // standard constructor
	virtual ~Dlg_Chonngay_paste();

// Dialog Data
	enum { IDD = DLG_CHNGAYP };

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	DECLARE_MESSAGE_MAP()

public:
	CStatic stt1,stt2;
	CString _ngbd,_ngkt;
	CString sNgayle[500];
	CString sGchu[500];
	CString sSaochepngay;

	afx_msg long f_Read_file_CSV(CString pathCSV);
	afx_msg long f_GetLineFile(CString szPath);
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	afx_msg void OnBnClickedCancel2();
};
