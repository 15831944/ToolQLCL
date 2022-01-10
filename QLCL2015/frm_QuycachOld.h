#pragma once
#include "stdafx.h"

// frm_QuycachOld dialog
class frm_QuycachOld : public CDialogEx
{
	DECLARE_DYNAMIC(frm_QuycachOld)

	public:
		frm_QuycachOld(CWnd* pParent = NULL);   // standard constructor
		virtual ~frm_QuycachOld();

	// Dialog Data
		enum { IDD = DLG_LAYMAUVL_OLD };

	protected:
		virtual void DoDataExchange(CDataExchange* pDX);  
		virtual BOOL OnInitDialog();
		DECLARE_MESSAGE_MAP()

	private:
		afx_msg BOOL PreTranslateMessage(MSG* pMsg);
		afx_msg void OnSysCommand(UINT nID, LPARAM lParam);

	public:
		CEdit m_txt[7];
		CButton m_chk_nthu;
		CComboBox m_cbb_ten,m_cbb_qcach;
		CScrollEdit m_txt_gc,m_txt_hd;
		CConnection ObjConn;

		_WorksheetPtr pSheet;
		RangePtr pRange;
		int iRow,iCol;
		int icolSTT,icotMAVL,icolLM,icolNdung;
		CString szKichthuoc;
		CString szTansuat;

		int sslvl;		
		struct MyStDSVL
		{
			CString szMa;
			CString szTen;
		}
		MSVlieu[100];

		int sslpp;
		struct MyStPP
		{
			CString szMa;
			CString szTenPP;
			CString szTSuat;
			CString szDVT;
		}
		MSPP[100];

	public:
		afx_msg void XacdinhQuycach();
		afx_msg void Loaddulieu(CString szMahieu);
		afx_msg CString AutoTinhSLMau(CString szKLNhap, CString szTansuat);

		afx_msg void OnCbnSelchangeCombo1();
		afx_msg void OnBnClickedBtnOk();
		afx_msg void OnBnClickedBtnCancel();
		afx_msg void OnBnClickedBtnMaukhac();
		afx_msg void OnCbnSelchangeCbb2();
		afx_msg void OnEnSetfocusTxt6();
		afx_msg void OnEnChangeTxt1();
};
