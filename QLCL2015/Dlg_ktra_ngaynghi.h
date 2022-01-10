#pragma once
#include "stdafx.h"

// Dlg_ktra_ngaynghi dialog
class Dlg_ktra_ngaynghi : public CDialog
{
	DECLARE_DYNAMIC(Dlg_ktra_ngaynghi)

public:
	Dlg_ktra_ngaynghi(CWnd* pParent = NULL);   // standard constructor
	virtual ~Dlg_ktra_ngaynghi();

// Dialog Data
	enum { IDD = DLG_KTNTG };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual BOOL OnInitDialog();
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);

public:
	CButton m_chk1,m_chk2,m_chk3,m_chk4;
	CButton m_rad1,m_rad2;
	CColorStatic m_stt1;

	_WorksheetPtr pSheet;
	RangePtr pRange;
	_bstr_t _code;
	int _irow,jbd,jkt;
	int jrdAll;
	CString sNgayle[1000];
	CString sGchu[1000];
	CString szCheckNgay;

	afx_msg long f_Read_Nhatky();
	afx_msg long f_Read_file_CSV(CString pathCSV, long num);
	afx_msg long f_GetLineFile(CString szPath);
	afx_msg void f_Kiemtra_ngay();
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedButton3();
	afx_msg void OnBnClickedButton2();
	afx_msg void OnBnClickedButton6();
	afx_msg void OnBnClickedRd1();
	afx_msg void OnBnClickedRd2();
	afx_msg void OnBnClickedCk1();
	afx_msg void OnBnClickedCk2();
	afx_msg void OnBnClickedCk3();
	afx_msg void OnBnClickedCk4();
};
