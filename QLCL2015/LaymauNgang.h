#pragma once
#include "stdafx.h"

class LaymauNgang : public CDialog
{
	DECLARE_EASYSIZE
	public:
		LaymauNgang(CWnd* pParent = NULL);	// standard constructor
	
		enum { IDD = DLG_LAYMAU3};

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
		afx_msg void OnSysCommand(UINT nID, LPARAM lParam);

	public:
		CComboBoxExt m_cbb1;
		CButton m_btn1,m_btn2,m_btn4;
		CButton m_chk1,m_chk2,m_chk_dong, m_chk_nthu;
		CEdit m_txt1,m_txt2,m_txt3;
		CNumSpinCtrl m_spin1;
		CToolTipCtrl m_ToolTips;

		CConnection ObjConn;
		int nItem, nSubItem;
		
		int iKeyX;
		CString pKeyX[10];

		_WorksheetPtr pSheet;
		RangePtr pRange,PRS;
		int iRow,iC2,iC4;		
		CString sCodeName,sDuoi;

		afx_msg void f_Load_tooltip();
		afx_msg void f_Get_size();
		afx_msg void _SetEnableNewMau(int itype);
		afx_msg void _ClearAllList();
		afx_msg void f_Load_dulieu();
		afx_msg void _StyleListLM3(CString sTenmau);
		afx_msg void _AutoColumnName();
		afx_msg void f_Luu_mau_ngang(CString sn);

		afx_msg void OnBnClickedB4();
		afx_msg void OnBnClickedCk1();
		afx_msg void OnBnClickedB1();		
		afx_msg void OnBnClickedB2();
		afx_msg void OnBnClickedB3();
		
		afx_msg void OnNMClickL1(NMHDR *pNMHDR, LRESULT *pResult);
		afx_msg void OnNMRClickL1(NMHDR *pNMHDR, LRESULT *pResult);
		
		afx_msg void OnCbnSelchangeCb1();
		afx_msg void OnBnClickedCk2();
		
		afx_msg void OnMau();			// Đổi tên mẫu
		afx_msg void OnMNML_Ten();		// Đổi tên cột
		afx_msg void OnThemdong();		// Thêm dòng
		afx_msg void OnThemcot();		// Thêm cột
		afx_msg void OnXoacot();		// Xóa cột
		afx_msg void OnXdong();			// Xóa dòng

		afx_msg void OnBnClickedB5();
		afx_msg void OnBnClickedB8();
		afx_msg void OnBnClickedCk3();
		afx_msg void OnBnClickedCk4();
		afx_msg void OnEnSetfocusE1();
		afx_msg void OnEnSetfocusT2();
		afx_msg void OnEnSetfocusT1();
		
		afx_msg void OnEnKillfocusE1();
		afx_msg void OnEnSetfocusT3();
		afx_msg void OnEnKillfocusT3();
		afx_msg void OnEnKillfocusT1();

		afx_msg void OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult);
		afx_msg void OnDeltaposSp1(NMHDR *pNMHDR, LRESULT *pResult);
		afx_msg void OnLvnKeydownL1(NMHDR *pNMHDR, LRESULT *pResult);		
		
		afx_msg void OnEnChangeT3();
		afx_msg void ChangeColumnWidth(CString szColumn);
};


