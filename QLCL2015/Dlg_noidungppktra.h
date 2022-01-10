#pragma once
#include "stdafx.h"

// Dlg_noidungppktra dialog
class Dlg_noidungppktra : public CDialogEx
{
	DECLARE_DYNAMIC(Dlg_noidungppktra)

public:
	Dlg_noidungppktra(CWnd* pParent = NULL);   // standard constructor
	virtual ~Dlg_noidungppktra();

// Dialog Data
	enum { IDD = DLG_NDPPKTRA };

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()

public:
	CConnection ObjConn;
	CEdit m_Txt,m_Them;
	CListCtrl m_List;

	int iClose,iTang;
	CString sFullPath;

	void f_Tack_tukhoa(CString sKey);
	void f_Tao_truyvan();
	void f_Truy_van_DL();

	afx_msg int f_Tim_kiem_NDPP();
	afx_msg void f_Load_DL();
	afx_msg void OnBnClickedB2();
	afx_msg void OnBnClickedB3();
	afx_msg void OnBnClickedB1();
	afx_msg void OnBnClickedChk1();
	afx_msg void OnBnClickedB4();
	afx_msg void OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult);
};
