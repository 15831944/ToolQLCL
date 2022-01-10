#pragma once
#include "stdafx.h"

// Dlg_luu_Hso dialog
class Dlg_luu_Hso : public CDialog
{
public:
	Dlg_luu_Hso();
	
	enum { IDD = DLL_LUUHS1};

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);  

protected:
	virtual BOOL OnInitDialog();
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);

public:
	CEdit Txt1;
	CToolTipCtrl m_ToolTips;
	CButton Crd1,Crd2,Chk0;
	
	CString sPathF,sCreate,sPFull;
	CString sDM,sBB,sVL,sCV,sGD;
	CString sRDay,sRTime;
	CConnection ObjConn;

	int iType,iChk,_itrang;
	int _demIn;

	_WorksheetPtr psPrint,psTS;
	RangePtr prPrint,prTS,PRS;
	_bstr_t shPrint,shTS;
	CString _bs1,_bs2;
	CString _ngaythang;

	BOOL Create_Folder(CString directoryPath);
	void f_GetTimeNow();
	void Luu_toan_bo_ho_so();
	void Luu_toan_hs_bsung();
	void f_lua_chon_bienban(int num_cviec, CString sCodename, int icb1, int icb2);	
	void Luu_chitiet_bienban(_bstr_t sPrint,int ithem,int num_cviec);
	void f_CodeNameActivate(_WorksheetPtr ps1,int num_cv);
	void f_Kiem_tra_inum(CString sPt,CString inum,int icb1, int icb2);
	void f_Kiem_tra_bsung(CString sPt,CString inum);
	void f_luu_danh_muc_hs(int icbb);
	void Luu_chitiet_danhmuc(_bstr_t sPrint);
	void xl_luu_nhat_ky();
	void f_Load_data();
	void f_luu_sheetpic(_WorksheetPtr ps1,CString sth,CString sduoi);

	afx_msg void f_Load_ToolTip();
	afx_msg void OnBnClickedR1();
	afx_msg void OnBnClickedR2();
	afx_msg void OnBnClickedCk1();
	afx_msg void OnBnClickedB1();
	afx_msg void OnBnClickedB2();
	afx_msg void OnBnClickedB3();
};

