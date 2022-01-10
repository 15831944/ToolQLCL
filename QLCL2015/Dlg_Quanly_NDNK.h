#pragma once
#include "stdafx.h"

class Dlg_Quanly_NDNK : public CDialog
{
	DECLARE_EASYSIZE
public:
	Dlg_Quanly_NDNK(CWnd* pParent = NULL);	// standard constructor
	
	enum { IDD = DLG_QLYNDNKY};

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);  

protected:
	virtual BOOL OnInitDialog();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnSizing(UINT fwSide, LPRECT pRect);
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnContextMenu(CWnd* /*pWnd*/, CPoint /*point*/);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);	
	afx_msg LRESULT OnHeaderRightClick(WPARAM wParam, LPARAM lParam);

public:
	_bstr_t shNK;
	_WorksheetPtr psNK;
	RangePtr prNK;
	CString _ngbd,_ngkt;
	CString sday_bd,sday_kt;

	CComboBox m_cbbox;
	CButton m_btn1,m_btn2;
	CButton m_chk2;
	CStatic m_thongbao;
	CConnection ObjConn;
	CToolTipCtrl m_ToolTips;
	COLORREF clr1,clr2;

	int iRsaochep;
	int gtlit;

	int TongNgay;
	int dwR,dwG,dwB;
	long gtriCSV;

	int nItem, nSubItem;

	CString _cpyNThg;
	CString szCheckNgay;
	CString NgayAll[2000];
	CString schep[25];
	CString szGtri[15];

	CString sNgayle[1000];
	CString sGchu[1000];

	struct MyStrImage
	{
		CString maso;
		CString duongdan;
		CString vitri;
		CString kichthuoc;
		CString ghichu;
	};

	int slImage;
	MyStrImage MSI[200];

	struct MyStrNhanlucThietbi
	{
		CString maso;
		CString cot[5];
	};

	int slNhanluc, slThietbi;
	MyStrNhanlucThietbi MSNhanluc[200];
	MyStrNhanlucThietbi MSThietbi[200];

public:
	afx_msg void f_Get_size();
	afx_msg void f_Get_width();
	afx_msg void f_Style_noidung();
	afx_msg void f_Delete_list();
	afx_msg void f_Ktra_dkien();
	afx_msg int f_Set_batdau(CString sbd);
	afx_msg int f_Set_ketthuc(CString skt);	
	afx_msg long f_GetLineFile(CString szPath);
	afx_msg long f_Read_Nhatky();
	afx_msg long f_Read_file_CSV(CString pathCSV, long num);
	afx_msg void f_Load_NULL(CString sdaybd, CString sdaykt);
	afx_msg void f_Load_ALL(CString sdaybd, CString sdaykt, int itype);		// itype=0 -> ALL | =1 chỉ load có dữ liệu
	afx_msg int f_Check_nghile(int row);
	afx_msg void f_TkiemT1();
	afx_msg void f_TkiemT2();
	afx_msg void f_Xac_dinh_ngay();
	afx_msg void f_Load_ToolTip();
	afx_msg int CheckDulieu(CString szngay);
	afx_msg void f_Load_SQL_ALL();	
	afx_msg void f_Load_SQL_01(CString szngay, CString szTbl, int col, int itypeall);
	afx_msg void f_Load_SQL_02(CString szngay, CString szTbl, CString Fld, int col, int itypeall);
	afx_msg void f_Load_SQL_03(CString szngay);
	afx_msg void f_Load_SQL_04(CString szngay);

	afx_msg void f_Tack_colorRGB(CString sColor);
	afx_msg void _CheckUpdate(CString sGtri);
	afx_msg void _SelectAllListMain();
	afx_msg void OnBnClickedB3();
	afx_msg void OnBnClickedB1();
	afx_msg void OnBnClickedB2();
	afx_msg void OnBnClickedCk2();
	afx_msg void OnNMClickL1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMRClickL1(NMHDR *pNMHDR, LRESULT *pResult);

	afx_msg void f_Fill_color_row(int nRow, CString s1, CString s2);
	afx_msg void f_Change_color(int cot);
	afx_msg void OnCbnSelchangeCb1();

	afx_msg void PasteDulieu(int itype1, int itype2, int itype3, int itype4, int itype5);
	afx_msg void GetNhancong(CString szNgay);
	afx_msg void GetThietbi(CString szNgay);
	afx_msg void GetHinhAnh(CString szNgay);
	
	afx_msg void OnNh40004();	// bỏ qua (nghỉ)
	afx_msg void OnNh40005();	// color background
	afx_msg void OnNh40006();	// color text	
	afx_msg void OnNh40064();

	afx_msg void OnNh40007();	// copy		
	afx_msg void OnD40053();	// dán all
	afx_msg void OnD40054();	
	afx_msg void OnD40055();
	afx_msg void OnD40057();
	afx_msg void OnD40056();	
	afx_msg void OnD40059();	

	afx_msg void OnX40012();	// xóa tất cả
	afx_msg void OnX40014();
	afx_msg void OnX40016();
	afx_msg void OnX40017();
	afx_msg void OnX40018();
	afx_msg void OnX40058();

	afx_msg void OnX40062();	// Xóa dữ liệu ALL
	afx_msg void OnX40063();	// Xóa toàn bộ dữ liệu file nhật ký

	afx_msg void OnDwShowAll();
	afx_msg void OnDwShowGhichu();
	afx_msg void OnDwShowThoitiet();
	afx_msg void OnDwShowNhancong();
	afx_msg void OnDwShowMaythicong();
	afx_msg void OnDwShowChitieu();
	afx_msg void OnDwShowNhanxet();
	afx_msg void OnDwShowVesinhATLD();
	afx_msg void OnDwShowPathHinhanh();

	afx_msg int CheckDeleteDulieu(CString szNgay, CString szTbl);
	afx_msg void Themdulieu_01(CString szNgay, CString szTbl, CString szGtri);
	afx_msg void Themdulieu_02(CString szNgay, CString szTbl, CString szGtri);
	afx_msg void Themdulieu_03(CString szNgay);

	afx_msg void PasteThoitietNhietdo(CString szNgay);
	afx_msg void PasteNhancong(CString szNgay);
	afx_msg void PasteMaythicong(CString szNgay);
	afx_msg void PasteGhichuNhanxet(CString szNgay);
	afx_msg void PasteHinhAnh(CString szNgay);
	afx_msg void PasteBangNhanluc(CString szNgay);
	afx_msg void PasteBangThietbi(CString szNgay);

	afx_msg void OnEnKillfocusE1();		// Thay đổi dữ liệu

	afx_msg void OnEnSetfocusE1();
	afx_msg void OnEnSetfocusT1();
	afx_msg void OnEnSetfocusT2();
	
	afx_msg void OnHdnEndtrackL1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void ShowColumnWidth(int nColumn, int nWidth, int itype);
};
