#pragma once
#include "stdafx.h"

class Dlg_Manager_wb : public CDialog
{
	DECLARE_EASYSIZE
	public:
		Dlg_Manager_wb(CWnd* pParent = NULL);	// standard constructor
	
		enum { IDD = DLG_QLWB};

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
		int nSheet,iCode;

	public:
		CToolTipCtrl m_toolTips;
		CButton m_check,m_btn[10];
		_WorksheetPtr pSheet;

		afx_msg void f_GetAllSheet();
		afx_msg void f_Load_AllSheet();
		afx_msg void f_Load_ToolTip();
		afx_msg void f_Style_List();
		afx_msg void f_Delete_List();
		afx_msg void f_Ktra_sheet();
		afx_msg void f_FindSheet();
		afx_msg void f_SetCheck_List(BOOLEAN f);

		afx_msg void OnBnClickedB1();
		afx_msg void OnBnClickedB2();
		afx_msg void OnBnClickedK1();
		afx_msg void OnBnClickedB3();
		afx_msg void OnEnKillfocusT1();
		afx_msg void OnBnClickedB4();
		afx_msg void OnBnClickedB5();
		afx_msg void OnBnClickedB6();
		afx_msg void OnBnClickedB7();
		afx_msg void OnLvnKeydownL1(NMHDR *pNMHDR, LRESULT *pResult);
		afx_msg void OnNMClickL1(NMHDR *pNMHDR, LRESULT *pResult);
		afx_msg void OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult);
		afx_msg void OnNMRClickL1(NMHDR *pNMHDR, LRESULT *pResult);
};



