#pragma once
#include "stdafx.h"


// Dlg_print_nhatky dialog
class Dlg_print_nhatky : public CDialog
{
public:
	Dlg_print_nhatky();
	
	enum { IDD = DLG_PRTNKY};

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);  

protected:
	virtual BOOL OnInitDialog();
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);

public:
	CConnection ObjConn;
	CEdit txt1;
	CStatic stt1;
	CButton btn1,fdkem,rd1,rd2,rd3,rd4,rd5,rd6;
	CDateTimeCtrl dt1,dt2;
	int _ird,_irp,_islpage;
	CString numin;
	CString _ngaythang;

	CString msg;
	LPTSTR lp;
	int iLen;
	
public:
	_WorksheetPtr psNDNK;
	RangePtr prNDNK,PRS;
	_bstr_t shNDNK;

	void f_Load_data();

	afx_msg CString f_ngaythang();
	afx_msg void f_set_datetime(CString _sTime);
	afx_msg void f_in_sheetpic(_WorksheetPtr ps1);
	afx_msg void f_set_enable(int type);
	afx_msg void OnBnClickedB2();
	afx_msg void OnBnClickedB3();
	afx_msg void OnBnClickedB1();
	afx_msg void OnBnClickedR1();
	afx_msg void OnBnClickedR2();
	afx_msg void OnBnClickedCk1();
	afx_msg void OnEnKillfocusT1();
	afx_msg void OnBnClickedB4();
	afx_msg void OnBnClickedB5();
	afx_msg void OnBnClickedR3();
	afx_msg void OnBnClickedR4();
	afx_msg void OnBnClickedR5();
	afx_msg void OnBnClickedR6();
};

