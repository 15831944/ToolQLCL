#pragma once
#include "stdafx.h"

class Dlg_TH_Cptuvan : public CDialog
{
	DECLARE_EASYSIZE
	public:
		Dlg_TH_Cptuvan(CWnd* pParent = NULL);	// standard constructor
	
		enum { IDD = DLG_THCPTV};

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
		CListCtrl m_list;
		_WorksheetPtr pSheet;
		RangePtr pRange,PRS;

	public:
		afx_msg void f_Get_size();
		afx_msg void f_Get_width();
		afx_msg void f_Style_List();
		afx_msg void f_Load_data();
		afx_msg long f_GetLineFile(CString szPath);
		afx_msg void OnBnClickedB1();
		afx_msg void OnBnClickedB2();
		afx_msg void OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult);
		afx_msg void OnLvnKeydownL1(NMHDR *pNMHDR, LRESULT *pResult);
		afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
};



