#pragma once
#include "stdafx.h"

// Dlg_Themtieuchuan dialog

class Dlg_Themtieuchuan : public CDialogEx
{
	DECLARE_DYNAMIC(Dlg_Themtieuchuan)

public:
	Dlg_Themtieuchuan(CWnd* pParent = NULL);   // standard constructor
	virtual ~Dlg_Themtieuchuan();

// Dialog Data
	enum { IDD = DLG_THEMNHOMTC };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual BOOL OnInitDialog();
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg BOOL OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message);
	afx_msg void OnContextMenu(CWnd* /*pWnd*/, CPoint /*point*/);

public:
	CComboBox m_cbbox;
	CComboBox m_cbbFile;
	CColorStatic m_title;
	CEdit m_txtTieuchuan,m_txtCongviec;
	CListCtrl m_listTC,m_listThem,m_listViewTC;
	CListCtrlEx m_listCongviec;
	CButton m_btnKqua,m_check,m_cksltc;
	CButton m_btnPre,m_btnNxt;
	CTextProgressCtrl m_progress;
	CConnection ObjConn,ObjConn2;
	
	CString sFullPath,sDDanfile;
	CString sFiledulieu[100];
	CString sFpathdulieu[100];
	int slfileDL;

	CString szlMaTC;
	int nItem,nListClick;
	int tslkq,lanshow,ibuocnhay;

	int iTCclick;
	CString sDownTC,sFullTC;

public:
	afx_msg void _OnExit();
	afx_msg void _TruyvanSQL();
	afx_msg void _TruyvanSQLCongviec();
	afx_msg void _SetStyleList();
	afx_msg void _LoadAllTC();
	afx_msg void _LoadAllCongviec();
	afx_msg int _CheckItemTC(CString szCheck);
	afx_msg void _ViewTC(int irow);
	afx_msg void _SetCheckList(int check);
	afx_msg void _XoaTieuchuan();
	afx_msg void _XoaTCOfCongviec();
	afx_msg void _XoaALLTCOfCongviec();
	afx_msg void f_Danh_sachfile();
	afx_msg void _LoadlistFile();

	afx_msg void OnBnClickedB1();
	afx_msg void OnBnClickedB2();
	afx_msg void OnBnClickedB3();
	afx_msg void OnBnClickedCk1();
	afx_msg void OnBnClickedB4();
	afx_msg void OnBnClickedB7();
	afx_msg void OnBnClickedB8();
	afx_msg void OnEnSetfocusT1();	

	afx_msg void OnBnClickedB5();
	afx_msg void OnBnClickedB6();
	afx_msg void OnBnClickedCk2();

	afx_msg CString _GetMaTieuchuan();

	afx_msg void OnTh40029();	// Tích
	afx_msg void OnTh40030();	// Bỏ tích
	afx_msg void OnTh40031();	// Thêm TC
	afx_msg void OnTh40032();	// Xóa TC
	afx_msg void OnTh40033();	// Thêm vào DS
	afx_msg void OnTh40034();	// Áp dụng TC vào CV
	afx_msg void OnTh40041();	// Đọc tiêu chuẩn
	afx_msg void OnTh40042();	// Google

	afx_msg void OnNMClickL2(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMRClickL2(NMHDR *pNMHDR, LRESULT *pResult);	
	afx_msg void OnNMClickL3(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMClickL1(NMHDR *pNMHDR, LRESULT *pResult);
	
	afx_msg void OnNMClickL4(NMHDR *pNMHDR, LRESULT *pResult);	
	afx_msg void OnNMRClickL1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMRClickL3(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMRClickL4(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult);	
	afx_msg void OnCbnSelchangeCb1();
	afx_msg void OnBnClickedPre();
	afx_msg void OnBnClickedNxt();
	afx_msg void OnCbnSelchangeCbbfl();
};
