#pragma once
#include "stdafx.h"

// Dlg_NhomBangmau dialog
class Dlg_NhomBangmau : public CDialogEx
{
	DECLARE_EASYSIZE
	DECLARE_DYNAMIC(Dlg_NhomBangmau)

public:
	Dlg_NhomBangmau(CWnd* pParent = NULL);   // standard constructor
	virtual ~Dlg_NhomBangmau();

	// Dialog Data
	enum { IDD = DLG_NHOMBMAU };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual BOOL OnInitDialog();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnSizing(UINT fwSide, LPRECT pRect);
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);

public:
	_bstr_t shBMau;
	_WorksheetPtr psBMau;
	RangePtr prBMau;

	CListCtrlEx m_list;
	CEdit m_txt_search;
	CComboBoxExt m_cbb_cviec, m_cbb_ctrinh;

	int pRowEnd;
	_WorksheetPtr pShDMuc;
	int iRowActive, iColumnActive;
	CString szTenmau;

	int slmau;
	struct MyStrDSMau
	{
		int irow;
		CString szten;
		CString szmota;
		CString szcviec;
		CString szctrinh;
	};

	vector<MyStrDSMau> vecMSDSM;

	afx_msg void _LoadDanhSach();
	afx_msg void _Timkiemdulieu();

	afx_msg	int _LocCongviec(int nIndex, int vt);
	afx_msg	int _LocCongtrinh(int nIndex, int vt);
	afx_msg int _LocNoidung(CString szNoidung, int vt);

	afx_msg void OnBnClickedB1();
	afx_msg void OnBnClickedB2();
	afx_msg void OnEnChangeTxt1();
	afx_msg void OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnLvnKeydownL1(NMHDR *pNMHDR, LRESULT *pResult);	
	afx_msg void OnCbnSelchangeCbb1();
	afx_msg void OnCbnSelchangeCbb2();
	afx_msg void OnBnClickedB5();
	afx_msg void OnBnClickedB4();
	afx_msg void OnNMClickL1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnBnClickedCheck1();
	afx_msg void f_Get_size();
	afx_msg void OnBnClickedB6();
};
