#pragma once
#include "stdafx.h"

class Dlg_Open_DMCV : public CDialog
{

DECLARE_EASYSIZE
public:
	Dlg_Open_DMCV(CWnd* pParent = NULL);	// standard constructor
	
	enum { IDD = DLG_DSCV};

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);  

protected:
	virtual BOOL OnInitDialog();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnSizing(UINT fwSide, LPRECT pRect);
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnContextMenu(CWnd* /*pWnd*/, CPoint /*point*/);

public:
	CColorEdit m_Txt;
	CComboBox m_cbbFile;
	CComboBoxExt m_cbbox;	
	CColorEdit m_txtTChon,m_txtColor,m_Group;
	CButton m_btn[2], m_chkDMCVCon,m_chkLuuTC;
	CButton m_cbk1,m_cbk2,m_cbk3;
	CButton bxx[9],m_btnPre,m_btnNxt;
	CButton m_chkColor,m_btnColor,m_chkBTC;
	CToolTipCtrl m_ToolTips;
	
	CConnection ObjConn,ObjConn2;
	_bstr_t shDMCV;
	_WorksheetPtr psDMCV;
	RangePtr prDMCV,PRS,PRS2;

	_bstr_t shBieugia,shBTC;
	_WorksheetPtr psBieugia,psBTC;
	RangePtr prBieugia,prBTC;

	int iTCclick;
	CString sDownTC,sFullTC;

	int iRow,iColumn,xde;
	int vtnhom;
	int iCpy;
	int nClick;
	int _iTab;
	int _isl1,_isl2;
	int _iCot[50];
	int _iCotEND,_iCotMaDM;

	CString sFullPath,sDDanfile;
	CString sFiledulieu[100];
	CString sFpathdulieu[100];
	int slfileDL;

	int iStopload;
	int tslkq,lanshow,ibuocnhay;

	int slMaTC;
	CString sKTMaTC[100];
	CString sKTTenTC[100];
	CString sLuuMaCT,sLuuTenCT,sLuuDVT,sLuuLink;
	CString sBefore,sAfter;
	CString szfDMNhancong,szfTNMay;
	CString szconDMNC[100], szconTNMay[100];

	CString sNgayle[1000];
	CString sGchu[1000];
	CString szCheckNgay;

	struct MyStructCVC
	{
		CString szMaDM;
		CString szMahieu;
		CString szNoidung;
		CString szDVT;
		CString szPath;
	};
	MyStructCVC MSCVC[100];

public:
	void f_Tao_truyvan();
	void f_Tack_tukhoa(CString sKey);
	void xl_xdcot_DMCV(_WorksheetPtr pSh);

	afx_msg void f_Tim_kiem_DMCV();
	afx_msg void f_Truy_van_DL();
	afx_msg void f_Danh_sachfile();
	afx_msg void f_Load_DL();
	afx_msg int f_Get_change();

	afx_msg void f_Get_size();
	afx_msg void f_StyleList();
	afx_msg void f_StyleList2();
	afx_msg void f_StyleList3();
	afx_msg void f_Load_ToolTip();
	afx_msg void f_SaveWidth();	
	afx_msg int f_Check_tieuchuan(CString sKtraTC);
	afx_msg void f_Load_all_bieugia();
	afx_msg void _EnterProject();
	afx_msg void _Change_Start(RangePtr prDMCV, int iRow, int jTime, int jStart, int jFinish, int jLienhe, int jTongnghi, int jColumnEnd);
	afx_msg	void _Change_Finish(RangePtr prDMCV, int iRow, int jTime, int jStart, int jFinish, int jLienhe, int jTongnghi, int jColumnEnd);
	afx_msg void _Check_nghile(RangePtr pr1, int ir, int ic);
	afx_msg long f_Read_Nhatky();
	afx_msg long f_Read_file_CSV(CString pathCSV, long num);
	afx_msg long f_GetLineFile(CString szPath);
	afx_msg int _CheckMultiItem();
	afx_msg void _getDinhmucNhancong(CString szMaCV, CString szLLk);

	afx_msg CString f_ten_tchuan(CString sVal);
	afx_msg void f_tk_tchuan(CString sVal);
	afx_msg void f_tk_ndung(CString sVal);
	afx_msg void f_UpDown_listTC(int iStyle);
	afx_msg void f_UpDown_listND(int iStyle);
	afx_msg void _UpdateDL(int iItemR, int itype);
	afx_msg void _GetDanhsachnhom();
	afx_msg void _LoadlistFile();
	afx_msg int _CheckTongnghaynghi(CString szDSNgay);

	afx_msg void OnBnClickedB1();
	afx_msg void OnBnClickedB2();	// <-- button OK
	afx_msg void OnBnClickedB3();	
	afx_msg void OnBnClickedB4();
	afx_msg void OnBnClickedB5();
	afx_msg void OnBnClickedB6();
	afx_msg void OnBnClickedB7();
	afx_msg void OnBnClickedB8();
	afx_msg void OnBnClickedB9();
	afx_msg void OnBnClickedB10();
	afx_msg void OnBnClickedB11();
	afx_msg void OnBnClickedCk1();
	
	afx_msg void OnEnKillfocusEd1();
	afx_msg void OnEnKillfocusEd2();
	afx_msg void OnEnKillfocusEd3();
	
	afx_msg void OnEnSetfocusEd1();
	afx_msg void OnEnSetfocusT1();
	afx_msg void OnEnSetfocusEd2();
	afx_msg void OnEnSetfocusEd3();
	afx_msg void OnBnClickedCbs1();
	afx_msg void OnBnClickedCbs2();
	afx_msg void OnBnClickedCbs3();
	afx_msg void OnBnClickedCkluu();
	afx_msg void OnBnClickedCbs4();
	afx_msg void OnBnClickedB13();
	afx_msg void OnBnClickedCbs5();
	afx_msg void OnBnClickedCk2();
	afx_msg void OnCbnSelchangeCbbfl();
	afx_msg void OnBnClickedPre();
	afx_msg void OnBnClickedNxt();

	afx_msg CString _GetMaTieuchuan();
	afx_msg void OnTi40039();	// Đọc
	afx_msg void OnTi40040();	// Google
	afx_msg void OnTi40051();

	afx_msg void OnNMClickL1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMClickL2(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMClickL3(NMHDR *pNMHDR, LRESULT *pResult);

	afx_msg void OnLvnKeydownL1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnLvnKeydownL2(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnLvnKeydownL3(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMDblclkL3(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMDblclkL2(NMHDR *pNMHDR, LRESULT *pResult);	
	afx_msg void OnNMRClickL2(NMHDR *pNMHDR, LRESULT *pResult);	
	afx_msg void OnNMRClickL1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnLvnEndScrollL1(NMHDR *pNMHDR, LRESULT *pResult);	
	afx_msg void OnBnClickedCkdmcvcon();
};

