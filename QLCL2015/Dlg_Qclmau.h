#pragma once
#include "stdafx.h"


class Dlg_Qclmau : public CDialog
{

DECLARE_EASYSIZE
public:
	Dlg_Qclmau(CWnd* pParent = NULL);	// standard constructor
	
	enum { IDD = DLG_QCLM};

	protected:
	virtual void DoDataExchange(CDataExchange* pDX); 

protected:
	virtual BOOL OnInitDialog();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnSizing(UINT fwSide, LPRECT pRect);
	afx_msg void OnStateChanged(NMHDR* pNMHDR, LRESULT* pResult);
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);

public:
	_WorksheetPtr pSheet;
	RangePtr pRange;
	int iRow,iCol;

	CButton m_chek;
	CComboBox cbb1;
	CColorEdit m_txt;
	CToolTipCtrl m_ToolTips;	
	CConnection ObjConn;
	
	int iC2,iC4;
	CString sCodeName;
	CString svlTen[500],svlDvi[500],svlPP[500],svlGtr[500];
	int tongRw;
	int iRClick,iCClick;	

public:
	afx_msg void f_Load_tooltip();
	afx_msg void f_Tack_quycach();
	afx_msg void f_Get_size();
	afx_msg void f_Get_width();
	afx_msg void f_Style_dsach();
	afx_msg void f_Delete_list();
	afx_msg void f_Load_dulieu(CString stype);
	afx_msg int f_Check_update();
	afx_msg int f_Update_CSDL();

	afx_msg void OnContextMenu(CWnd* /*pWnd*/, CPoint /*point*/);
	afx_msg void OnTi40002();		// Thêm mới
	afx_msg void OnTi40003();		// Xóa
	afx_msg void OnTi40001();		// F4: Thêm chỉ tiêu từ mẫu
	afx_msg void OnBnClickedB2();
	afx_msg void OnBnClickedB1();
	afx_msg void OnEnKillfocusT1();		// Chỉnh sửa text
	afx_msg void OnCbnSelchangeCb1();
	afx_msg void OnBnClickedB3();
	afx_msg void OnBnClickedB4();
	afx_msg void OnBnClickedB5();
	afx_msg void OnLvnKeydownL1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnBnClickedCk1();
	afx_msg void OnNMClickL1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnBnClickedB6();
	afx_msg void OnEnSetfocusT2();
	afx_msg void OnEnSetfocusT1();
	afx_msg void OnBnClickedB8();
	afx_msg void OnBnClickedB7();
	afx_msg void OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult);
};


