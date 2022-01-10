#pragma once
#include "stdafx.h"


class Dlg_all_mauvl : public CDialog
{

DECLARE_EASYSIZE
public:
	Dlg_all_mauvl(CWnd* pParent = NULL);	// standard constructor
	
	enum { IDD = DLG_ALL_MVL};

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);  

protected:
	virtual BOOL OnInitDialog();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnSizing(UINT fwSide, LPRECT pRect);
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);

public:
	CConnection ObjConn;
	CEdit m_Txt,m_MaLm,m_Pplm,m_Tslm,m_Dvi;
	CButton m_btn[5];
	CStatic m_Group;
	CListCtrl m_List;
	CToolTipCtrl m_ToolTips;

	int iClose,iTang;
	CString sFullPath;

public:
	void f_Tack_tukhoa(CString sKey);
	void f_Tao_truyvan();
	void f_Truy_van_DL();

	afx_msg int f_Tim_kiem_MVL();
	afx_msg void f_StyleList();
	afx_msg void f_Load_ToolTip();
	afx_msg void f_SaveWidth();
	afx_msg void f_Xoa_mauvl();
	afx_msg void f_Load_DL();
	
	afx_msg void OnBnClickedB1();
	afx_msg void OnBnClickedB2();
	afx_msg void OnBnClickedB4();
	afx_msg void OnBnClickedB3();
	afx_msg void OnBnClickedB5();
	afx_msg void OnBnClickedCk1();
	afx_msg void OnLvnKeydownL1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMClickL1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult);
};


