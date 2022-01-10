#pragma once
#include "stdafx.h"
#include "PictureCtrl.h"
#include "MultiFileOpenDialog.h"

class Dlg_Qly_image_nky : public CDialog
{

DECLARE_EASYSIZE
public:
	Dlg_Qly_image_nky(CWnd* pParent = NULL);	// standard constructor
	
	enum { IDD = DLG_QLIMAGE};

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);  

protected:
	virtual BOOL OnInitDialog();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnSizing(UINT fwSide, LPRECT pRect);
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);


public:
	CEdit txt1;
	CStatic stt1,stt2;
	CButton btn[5];
	CListCtrl listIMG;
	CToolTipCtrl m_ToolTips;
	CPictureCtrl m_picCtrl;
	CString msg,_fullpath;
	CConnection ObjConn;
	CString sNgay;
	int nClick;

public:
	afx_msg void f_Load_ToolTip();
	afx_msg void f_Style_ListIMG();
	afx_msg void f_Load_data();

	afx_msg void OnBnClickedB8();
	afx_msg void OnBnClickedB7();
	afx_msg void OnNMClickL1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnEnKillfocusT1();
	afx_msg void OnBnClickedB2();
	afx_msg void OnBnClickedB1();
	afx_msg void OnBnClickedB3();
	afx_msg void OnBnClickedB4();
};
