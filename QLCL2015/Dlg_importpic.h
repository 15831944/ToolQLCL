#pragma once
#include "stdafx.h"
#include "Path.h"
#include "PictureCtrl.h"
#include "FolderDlg.h"

class Dlg_importpic : public CDialog
{

DECLARE_EASYSIZE
public:
	Dlg_importpic(CWnd* pParent = NULL);	// standard constructor
	
	enum { IDD = DLG_IMPPIC};

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);  

protected:
	virtual BOOL OnInitDialog();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnSizing(UINT fwSide, LPRECT pRect);
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	CFileFinder	_finder;

public:
	CEdit txt1;
	CStatic stt1,stt2;
	CButton btn[5];
	CListCtrl list1;
	CToolTipCtrl m_ToolTips;
	CPictureCtrl m_picCtrl;
	CString msg,_fullpath;

public:
	afx_msg void f_Delete_List();
	afx_msg void f_Load_List();
	afx_msg void f_Load_ToolTip();
	afx_msg void f_Add_Csdl(CString m_sBaseFolder, CString m_sFileMask);

	afx_msg void OnBnClickedB1();
	afx_msg void OnBnClickedB2();
	afx_msg void OnBnClickedB3();
	afx_msg void OnNMClickL1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnBnClickedB4();
};
