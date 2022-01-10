#pragma once
#include "stdafx.h"


class Dlg_Open_DSVL : public CDialog
{

DECLARE_EASYSIZE
public:
	Dlg_Open_DSVL(CWnd* pParent = NULL);	// standard constructor
	
	enum { IDD = DLG_DSVL};

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
	CEdit m_txtTenBTC;
	CColorEdit m_Txt,m_txtTChon;
	CComboBox m_cbbFile;
	CComboBoxExt m_cbbox;
	CButton bxx3,m_cbk1,m_cbk2,m_chkBTC,m_chkLMau,m_chkLuuTC;
	CToolTipCtrl m_ToolTips;
	CColorEdit m_txtColor,m_Group2;
	CButton m_chkColor,m_btnColor,m_btnPre,m_btnNxt;

	CConnection ObjConn,ObjConn2;
	_bstr_t shDMVL,shBTC;
	_WorksheetPtr psDMVL,psBTC;
	RangePtr prDMVL,prBTC,PRS,PRS2;

	int iTCclick;
	CString sDownTC,sFullTC;

	int iRow,iColumn,xde;
	int vtnhom;
	int iCpy;
	int nClick;
	int _iTab;
	int _isl1,_isl2;
	int _iCot[30];
	int _iCotEND;
	CString sMau[3];

	CString sFullPath,sDDanfile;
	CString sFiledulieu[100];
	CString sFpathdulieu[100];
	int slfileDL;

	int tslkq,lanshow,ibuocnhay;
	int iStopload;
	int slMaTC;
	CString sKTMaTC[100];
	CString sKTTenTC[100];
	CString sLuuMaCT,sLuuTenCT,sLuuDVT,sLuuLink;
	CString sBefore,sAfter;

public:
	void f_Tao_truyvan();
	void f_Tack_tukhoa(CString sKey);
	void xl_xdcot_DMVL(_WorksheetPtr pSh);

	afx_msg void f_Tim_kiem_DMVL();
	afx_msg void f_Truy_van_DL();
	afx_msg void f_Load_DL();
	afx_msg void f_Danh_sachfile();
	afx_msg int f_Get_change();
	afx_msg int f_Check_tieuchuan(CString sKtraTC);

	afx_msg void f_Get_size();
	afx_msg void f_Load_ToolTip();
	afx_msg void f_SaveWidth();
	afx_msg void f_StyleList();
	afx_msg void f_StyleList2();
	afx_msg void f_StyleList3();

	afx_msg CString f_ten_tchuan(CString sVal);
	afx_msg void f_tk_tchuan(CString sVal);
	afx_msg void f_tk_ndung(CString sVal);	
	afx_msg void f_ten_mauvl(CString sVal);
	afx_msg void f_UpDown_listTC(int iStyle);
	afx_msg void f_UpDown_listND(int iStyle);
	afx_msg void _UpdateDL(int iItemR, int itype);
	afx_msg void _GetDanhsachnhom();
	afx_msg void _LoadlistFile();

	afx_msg void OnBnClickedB1();
	afx_msg void OnBnClickedB2();
	afx_msg void OnBnClickedB4();
	afx_msg void OnBnClickedB5();
	afx_msg void OnBnClickedB3();
	afx_msg void OnBnClickedB6();
	afx_msg void OnBnClickedB7();
	afx_msg void OnBnClickedB8();
	afx_msg void OnBnClickedB9();
	afx_msg void OnBnClickedB10();
	afx_msg void OnBnClickedB11();

	afx_msg void OnEnKillfocusEd1();
	afx_msg void OnEnKillfocusEd2();
	afx_msg void OnEnKillfocusEd3();
	
	afx_msg void OnEnSetfocusT1();
	afx_msg void OnBnClickedCbs1();
	afx_msg void OnBnClickedCbs2();
	afx_msg void OnEnSetfocusEd1();
	afx_msg void OnEnSetfocusEd2();
	afx_msg void OnEnSetfocusEd3();
	afx_msg void OnBnClickedCbs3();
	afx_msg void OnBnClickedB12();
	afx_msg void OnBnClickedCbs4();

	afx_msg void OnBnClickedPre();
	afx_msg void OnBnClickedNxt();
	afx_msg void OnCbnSelchangeCbbfl();
	afx_msg void OnBnClickedCkluu();

	afx_msg CString _GetMaTieuchuan();
	afx_msg void OnTi40039();
	afx_msg void OnTi40040();
	afx_msg void OnTi40052();

	afx_msg void OnLvnKeydownL1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnLvnKeydownL2(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnLvnKeydownL3(NMHDR *pNMHDR, LRESULT *pResult);

	afx_msg void OnNMClickL1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMClickL2(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMClickL3(NMHDR *pNMHDR, LRESULT *pResult);	
	afx_msg void OnNMRClickL2(NMHDR *pNMHDR, LRESULT *pResult);	
	
	afx_msg void OnNMDblclkL2(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMDblclkL3(NMHDR *pNMHDR, LRESULT *pResult);	
	afx_msg void OnLvnEndScrollL1(NMHDR *pNMHDR, LRESULT *pResult);

};

