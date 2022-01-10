#pragma once
#include "stdafx.h"

// Dlg_Timer dialog
class Dlg_Timer : public CDialogEx
{
	DECLARE_DYNAMIC(Dlg_Timer)

public:
	Dlg_Timer(CWnd* pParent = NULL);   // standard constructor
	virtual ~Dlg_Timer();

// Dialog Data
	enum { IDD = DLG_TIMER };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);  

protected:
	virtual BOOL OnInitDialog();
	afx_msg void OnTimer(UINT_PTR nIDEvent);
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);

public:
	CButton cck1;
	CMonthCalCtrl cmc1;
	int iDbclick;
	CString sDbcl[3];
	CString sNgayle[1000];
	CString sGchu[1000];
	CString szCheckNgay;

	_WorksheetPtr pSh1;

	afx_msg long f_Read_Nhatky();
	afx_msg long f_Read_file_CSV(CString pathCSV, long num);
	afx_msg long f_GetLineFile(CString szPath);
	afx_msg int f_SetGtriDate(RangePtr pRx,CString csdate,int iR,int iC,int iTG,int type);
	afx_msg void f_Autocheck_date(RangePtr pr1,int ir,int ic);
	afx_msg void f_Check_nghile(RangePtr pr1,int ir,int ic);
	afx_msg void f_set_date();
	afx_msg void f_Set_location();
	afx_msg void OnBnClickedB2();
	afx_msg void OnBnClickedB1();
	afx_msg void _Change_Start(RangePtr prDMCV, int iRow, int jTime, int jStart, int jFinish, int jLienhe, int jTongnghi, int jColumnEnd);
	afx_msg void _Change_Finish(RangePtr prDMCV, int iRow, int jTime, int jStart, int jFinish, int jLienhe, int jTongnghi, int jColumnEnd);
	afx_msg void OnMcnSelectMc1(NMHDR *pNMHDR, LRESULT *pResult);
};
