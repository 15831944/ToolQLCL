#pragma once
#include "stdafx.h"

// Dlg_AutoSTT_HSNT dialog
class Dlg_AutoSTT_HSNT : public CDialogEx
{
	DECLARE_DYNAMIC(Dlg_AutoSTT_HSNT)

public:
	Dlg_AutoSTT_HSNT(CWnd* pParent = NULL);   // standard constructor
	virtual ~Dlg_AutoSTT_HSNT();

// Dialog Data
	enum { IDD = AUTO_HSNT_STT };

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);

public:
	CEdit m_txtSTT,m_txtBD,m_txtKT;
	CNumSpinCtrl m_spinSTT,m_spinBD,m_spinKT;
	CButton m_ckTT,m_ckHT;
	CButton m_rdSel,m_rdHM,m_rdGD;
	CComboBox m_cbbHM;
	CComboBoxExt m_cbbTen;

	_WorksheetPtr pSheet;
	RangePtr pRange;

	int iUpdate,iSLTT,iSLHT;
	int ixdEnd;
	int slHM,slGD;
	int iVBatdau,iVKetthuc;
	int iCotID,iCotMaCV,iCotMaHS,iCotND;
	int iVtrHM[4000],iVtrGD[4000];
	CString sTenHM[4000],sTenGD[4000];

public:

	afx_msg void _XoaCBB();
	afx_msg void _Xacdinhvitri();	
	afx_msg void _LoadMaHSNT(int itype);
	afx_msg void _SetEnableBDKT(int itype);
	afx_msg void _CheckMaHSNT(CString sTen);

	afx_msg void OnBnClickedB1();
	afx_msg void OnBnClickedB2();

	afx_msg void OnBnClickedRd2();
	afx_msg void OnBnClickedRd3();
	afx_msg void OnBnClickedRd1();

	afx_msg void OnBnClickedCk1();
	afx_msg void OnBnClickedCk2();
};
