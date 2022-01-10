#pragma once
#include "stdafx.h"

class Dg_In_ho_so : public CDialog
{

DECLARE_EASYSIZE
public:
	Dg_In_ho_so(CWnd* pParent = NULL);	// standard constructor
	
	enum { IDD = DLG_PRINT};

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);  

protected:
	virtual BOOL OnInitDialog();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnSizing(UINT fwSide, LPRECT pRect);
	DECLARE_MESSAGE_MAP()

private:
	afx_msg void OnTimer(UINT_PTR nIDEvent);
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);

public:
	CListCtrl list1;
	CEdit txt[3];
	CButton btn[8],chk[9];
	CButton btn55,btn77;
	CComboBox cbb[8];
	CToolTipCtrl m_ToolTips;
	CStatic kqgrp;
	CColorStatic sttHuongdan;
	
public:
	_WorksheetPtr psPrint;
	RangePtr prPrint,PRS;
	_bstr_t shPrint;
	
	CString msg;
	LPTSTR lp;
	int iLen;
	int iSec,iSetTm;

	int slTTVL,slTTCV,slTTGD;
	CString sDotTT_VL[200],sDotTT_CV[200],sDotTT_GD[200];
	CString fchon[15];
	int igt[8];
	wchar_t *txt_timkiem;
	wchar_t pKey[150][10];
	int iKey,_itrang;
	int iKieuIn,iSLBanIn;
	CString _bs1,_bs2;

public:	
	afx_msg void OnBnClickedB3();
	afx_msg void OnBnClickedB2();
	afx_msg void OnBnClickedB1();
	afx_msg void OnCbnSelchangeC1();
	afx_msg void OnCbnSelchangeC2();
	afx_msg void OnCbnSelchangeC3();
	afx_msg void OnBnClickedCk1();
	afx_msg void OnBnClickedCk2();

	afx_msg void f_Load_Hoso();
	afx_msg void f_Load_Bienban();
	afx_msg void f_Load_ToolTip();
	afx_msg void f_CheckBtn55(int a1);
	afx_msg void f_CheckNUM1(int a1,int a2);	
	afx_msg void f_Load_DotThanhtoan(CString sdotTT);
	afx_msg void _SetLecongiviec();

	afx_msg void f_thietlap_giatri();
	afx_msg void f_Style_ListCV();
	afx_msg void f_Delete_ListCV();
	afx_msg void f_Add_congviec();
	afx_msg int f_Xac_dinh_sl_In();
	afx_msg void f_xacdinh_sl_bban();
	afx_msg void f_Xdinh_GiaiDoan(int istyle);
	afx_msg void f_Load_GiaiDoan(int igdoan);
	afx_msg void f_Load_ListVLGD(_bstr_t pSheet);
	afx_msg int f_Num_SLBB(_bstr_t pSheet, CString NameColumn,int istart);
	afx_msg CString f_Gettext(int n);
	afx_msg void f_Set_Duieu(int fset);
	afx_msg void f_KTLaymau_VL_CV();
	afx_msg void f_Load_all_sheet();
	afx_msg void f_Dot_thanh_toanVL();
	afx_msg void f_Dot_thanh_toanCV();
	afx_msg void f_Dot_thanh_toanGD();
	afx_msg void f_CodeNameActivate(_WorksheetPtr ps1);
	afx_msg void f_Kiem_tra_inum(CString inum,int icb1, int icb2);
	afx_msg void f_Kiem_tra_bsung(int istyle,CString inum,int iBanIn);
	afx_msg void f_SetCheck_VLCV(int ick1, int ick2, int ick3, int ick4, int ick5, int ick6);
	
	afx_msg void In_chitiet_bienban(_bstr_t sPrint,int ithem,int num_cviec);
	afx_msg void In_toan_bo_ho_so();
	afx_msg void In_toan_hs_bsung(int istyle,int iBanIn);
	afx_msg void f_lua_chon_bienban(int num_cviec, CString sCodename, int icb1, int icb2);
	afx_msg void f_in_danh_muc_hs(int icbb);
	afx_msg void In_chitiet_danhmuc(CString sPrint);

	afx_msg void OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnLvnKeydownL1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnBnClickedCk3();
	afx_msg void OnEnKillfocusT2();
	afx_msg void OnBnClickedB4();
	afx_msg void OnBnClickedCk5();
	afx_msg void OnBnClickedCk6();
	afx_msg void OnBnClickedCk7();
	afx_msg void OnBnClickedB5();
	afx_msg void OnBnClickedB6();
	afx_msg void OnBnClickedCk4();
	afx_msg void OnBnClickedB7();	
	afx_msg void OnBnClickedCk8();
	afx_msg void OnCbnSelchangeC4();
};