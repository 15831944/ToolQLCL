#pragma once
#include "stdafx.h"


class Dlg_all_tieuchuan : public CDialog
{

DECLARE_EASYSIZE
public:
	Dlg_all_tieuchuan(CWnd* pParent = NULL);	// standard constructor
	
	enum { IDD = DLG_ALL_TC};

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);  

protected:
	virtual BOOL OnInitDialog();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnSizing(UINT fwSide, LPRECT pRect);
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnContextMenu(CWnd* /*pWnd*/, CPoint /*point*/);

public:
	CConnection ObjConn;
	CEdit m_Txt,m_Matc,m_Tentc,m_Loaitc;
	CButton m_btn[5];
	CStatic m_Group;
	CListCtrl m_List;
	CToolTipCtrl m_ToolTips;

	int iVtTC,iClose,iTang;
	CString sFullPath;

	int iTCclick;
	CString sDownTC,sFullTC;

public:
	void f_Tack_tukhoa(CString sKey);
	void f_Tao_truyvan();
	void f_Truy_van_DL();

	afx_msg int f_Tim_kiem_TC();
	afx_msg void f_StyleList();
	afx_msg void f_Load_ToolTip();
	afx_msg void f_SaveWidth();
	afx_msg void f_Xoa_tieuchuan();
	afx_msg void f_Load_DL();

	afx_msg void OnBnClickedB1();
	afx_msg void OnBnClickedB2();
	afx_msg void OnBnClickedB4();
	afx_msg void OnBnClickedB3();
	afx_msg void OnBnClickedB5();
	afx_msg void OnBnClickedCk1();

	afx_msg CString _GetMaTieuchuan();
	afx_msg void OnTi40039();
	afx_msg void OnTi40040();

	afx_msg void OnLvnKeydownL1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMClickL1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult);	
	afx_msg void OnNMRClickL1(NMHDR *pNMHDR, LRESULT *pResult);	
};

