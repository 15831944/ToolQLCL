#pragma once
#include "stdafx.h"

// Dlg_Chen_dong dialog
class Dlg_Chen_dong : public CDialog
{
public:
	Dlg_Chen_dong();
	
	enum { IDD = DLG_CHENDONG};

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);  

protected:
	virtual BOOL OnInitDialog();
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);

public:
	CEdit txt1;
	CButton btn1,btn2;
	
public:
	_WorksheetPtr pSheet;
	RangePtr pRange,PRS;

	afx_msg void OnBnClickedCdgbtn2();
	afx_msg void OnBnClickedCdgbtn1();
};
