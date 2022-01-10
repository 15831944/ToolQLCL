#pragma once
#include "stdafx.h"

// Dlg_move_dulieu dialog
class Dlg_move_dulieu : public CDialog
{
	DECLARE_DYNAMIC(Dlg_move_dulieu)

public:
	Dlg_move_dulieu(CWnd* pParent = NULL);   // standard constructor
	virtual ~Dlg_move_dulieu();

// Dialog Data
	enum { IDD = DLG_MOVEDLIEU };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual BOOL OnInitDialog();
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);

public:
	CComboBox m_cbb1,m_cbb2;
	CButton m_chk1,m_chk2,m_chk3, m_chk4;

public:
	_bstr_t shDMCV;
	_WorksheetPtr psDMCV;
	RangePtr prDMCV;

	int vthm,vtgd,vtketiep;
	int iEnd,iCdm1,iCdm4,iCdm7,iCdm12;

	afx_msg void f_Xac_dinh_hangmuc();

	afx_msg void OnBnClickedB1();
	afx_msg void OnBnClickedB2();

	afx_msg void OnCbnSelchangeCb1();
	afx_msg void OnCbnSelchangeCb2();
};
