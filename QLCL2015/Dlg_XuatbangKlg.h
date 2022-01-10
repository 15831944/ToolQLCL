#pragma once
#include "stdafx.h"

// Dlg_XuatbangKlg dialog
class Dlg_XuatbangKlg : public CDialog
{
	DECLARE_DYNAMIC(Dlg_XuatbangKlg)

public:
	Dlg_XuatbangKlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~Dlg_XuatbangKlg();

// Dialog Data
	enum { IDD = DLG_XUATKLGD };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual BOOL OnInitDialog();
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);

public:
	CComboBox m_cbb1,m_cbb2;
	CButton m_chk;

	_WorksheetPtr pSheet;
	RangePtr pRange,PRS;
	_bstr_t _bsName,_bsCNM;
	CString sNameSh;

	int iEnd,ibd,ikt;
	int Msrt[1000],Msgd[1000];
	int iNgay[10];
	int _iCot[30];

public:
	afx_msg void xl_xdcot_DMCV(_WorksheetPtr pSh);
	afx_msg void f_Load_hangmuc();
	afx_msg void f_xacdinh_gdoan(int num1,int num2);
	afx_msg void OnBnClickedB1();
	afx_msg void OnBnClickedB2();
};
