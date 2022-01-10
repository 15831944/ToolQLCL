#pragma once
#include "stdafx.h"

class Dlg_trauu_tieuchuan : public CDialog
{
	DECLARE_EASYSIZE
public:
	Dlg_trauu_tieuchuan(CWnd* pParent = NULL);	// standard constructor
	
	enum { IDD = DLG_TCTC};

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);  

protected:
	virtual BOOL OnInitDialog();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnSizing(UINT fwSide, LPRECT pRect);
	DECLARE_MESSAGE_MAP()

private:
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnContextMenu(CWnd* /*pWnd*/, CPoint /*point*/);
public:
	CListCtrl list1;
	CEdit txtKey,txtLTC;
	CStatic sttic;
	
	wchar_t *txt_timkiem;
	int iKey,iLTC;
	wchar_t pKey[20][500],pLTC[20][500];
	CString xl_tukhoa,xl_LTC;
	int iRow,iColumn;
	long iKiemtra;
	CString val1,val2,val3;
	CString _matc[5000],_nametc[5000];
	int i1,i2,i3,xde;
	int cotHyperlink;

	CString msg;
	LPTSTR lp;
	int iSL;

	int tslkq,lanshow,ibuocnhay;
	int iStopload;

public:
	CConnection ObjConn;
	_WorksheetPtr pSheet;
	RangePtr pRange,PRS;
	_bstr_t _code;

	int iTCclick;
	CString sDownTC,sFullTC;

public:
	afx_msg void xl_timkiem_tieuchuan();
	afx_msg long xl_kiemtra_dulieu();
	afx_msg void xl_xacdinh_sheet(int sl);
	afx_msg void f_Style_tieuchuan();
	afx_msg void f_xacdinh_SQL(CString _sql2);
	afx_msg CString f_xacdinh_sqlLTC();
	afx_msg void f_Tack_tu_khoa(CString tukhoa);
	afx_msg void f_Tack_loai_tieuchuan(CString tukhoa);
	afx_msg void f_Load_list_tieuchuan();
	afx_msg void f_delete_list();
	afx_msg void OnBnClickedB2();
	afx_msg void OnBnClickedB1();
	afx_msg void OnBnClickedB3();
	afx_msg void OnBnClickedB4();
	afx_msg void OnBnClickedB5();	

	afx_msg CString _GetMaTieuchuan();
	afx_msg void OnTi40039();
	afx_msg void OnTi40040();

	afx_msg void OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnLvnKeydownL1(NMHDR *pNMHDR, LRESULT *pResult);	
	afx_msg void OnNMClickL1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMRClickL1(NMHDR *pNMHDR, LRESULT *pResult);

	
	afx_msg void OnLvnEndScrollL1(NMHDR *pNMHDR, LRESULT *pResult);
};
