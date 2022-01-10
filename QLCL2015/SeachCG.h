#pragma once
#include "stdafx.h"

struct TienLuongCG{
	CString macg;
	CString hoten;
	CString chucd;
	CString hsl;
	CString hspc;
	CString cpnc;
	int	index;
};

class SeachCG : public CDialog
{
public:
	SeachCG();
	
	enum { IDD = DLG_CHGIA};

	int num,inp;
	CListCtrl listTCG;
	TienLuongCG CGArr[200];
	TienLuongCG pCGArr[200];
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);  

protected:
	virtual BOOL OnInitDialog();
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);

public:
	int readCG();
	void fillcell(int irn);
	afx_msg void OnBnClickedBtnchg1();
	afx_msg void OnBnClickedBtnchg2();
	afx_msg void OnLvnKeydownListchg1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMDblclkListchg1(NMHDR *pNMHDR, LRESULT *pResult);
};