#pragma once
#include "stdafx.h"""

// Dlg_Viewbangmau dialog

class Dlg_Viewbangmau : public CDialogEx
{
	DECLARE_DYNAMIC(Dlg_Viewbangmau)

	public:
		Dlg_Viewbangmau(CWnd* pParent = nullptr);   // standard constructor
		virtual ~Dlg_Viewbangmau();

	// Dialog Data
	#ifdef AFX_DESIGN_TIME
		enum { IDD = DLG_VIEWBMAU };
	#endif

	protected:
		virtual BOOL OnInitDialog();
		virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
		DECLARE_MESSAGE_MAP()

	private:
		afx_msg BOOL PreTranslateMessage(MSG* pMsg);
		afx_msg void OnSysCommand(UINT nID, LPARAM lParam);

	public:
		CString szNumRow;
		CString szTenbangmau;
		CListCtrlEx m_list;

		struct MyStrDScot
		{
			int icol;
			int iwidth;
			CString szten;
		};

		MyStrDScot MSDSC[50];

		afx_msg void LoadList();
		afx_msg void SetLocation();

		afx_msg void OnBnClickedButton1();
};
