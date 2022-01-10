#pragma once
#include "stdafx.h"

// Dlg_TimeDM dialog
class Dlg_TimeDM : public CDialogEx
{
	DECLARE_DYNAMIC(Dlg_TimeDM)

public:
	Dlg_TimeDM(CWnd* pParent = NULL);   // standard constructor
	virtual ~Dlg_TimeDM();

// Dialog Data
	enum { IDD = DLG_TMIEDM };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);  

protected:
	virtual BOOL OnInitDialog();
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);

public:
	_WorksheetPtr pSh1;
	RangePtr pRg1;
	int _irow,_icol;

	CButton m_chk1,m_chk2;
	CListBox m_list[2];
	CNumSpinCtrl m_spin[4];
	CEdit m_txt[4];

	int iChk,iNext;
	int irts,icts;

	afx_msg void f_Check_spin(int type, int num);
	afx_msg void f_Check_int(int type, int num);
	afx_msg void f_Tack_thoigian(CString stime);
	afx_msg void f_Tack_time2(CString stime);
	afx_msg void f_Auto_time();
	afx_msg void f_Get_time(int num);
	afx_msg void f_Delete_list(int num);
	afx_msg void f_Load_listbox();
	afx_msg void f_Set_spin();
	afx_msg void f_Set_time();
	afx_msg void f_Set_location();
	afx_msg void f_Luu_time_TS(CString sLuu);

	afx_msg void OnBnClickedB1();
	afx_msg void OnBnClickedB2();
	afx_msg void OnDeltaposSp1V1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnDeltaposSp1V2(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnDeltaposSp1V3(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnDeltaposSp1V4(NMHDR *pNMHDR, LRESULT *pResult);
	
	afx_msg void OnEnKillfocusT1V1();
	afx_msg void OnEnKillfocusT1V3();
	afx_msg void OnEnKillfocusT1V2();
	afx_msg void OnEnKillfocusT1V4();
	afx_msg void OnLbnDblclkL1();
	afx_msg void OnLbnDblclkL2();
	afx_msg void OnLbnSelchangeL1();
	afx_msg void OnLbnSelchangeL2();
	afx_msg void OnBnClickedS2();
};
