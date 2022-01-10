#pragma once
#include "stdafx.h"

// Dlg_dsach_ctieu dialog
class Dlg_dsach_ctieu : public CDialogEx
{
	DECLARE_DYNAMIC(Dlg_dsach_ctieu)

public:
	Dlg_dsach_ctieu(CWnd* pParent = NULL);   // standard constructor
	virtual ~Dlg_dsach_ctieu();

// Dialog Data
	enum { IDD = DLG_QCDSLM };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual BOOL OnInitDialog();
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);

public:
	CListCtrl m_list;
	CConnection ObjConn;
	CButton m_check;
	CEdit m_txt,m_t1,m_t2;

	afx_msg void f_Style_dsach();
	afx_msg void f_Delete_list();
	afx_msg void f_Load_dulieu();
	afx_msg void f_Delete_item();
	afx_msg void f_Tack_tukhoa(CString sKey);
	afx_msg void f_Tao_truyvan(CString sBosung);

	afx_msg void OnBnClickedButton3();
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedButton4();
	afx_msg void OnLvnKeydownL1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnBnClickedButton5();
	afx_msg void OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult);
};
