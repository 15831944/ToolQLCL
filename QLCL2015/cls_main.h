#include "stdafx.h"

class cls_main
{

public:
	_WorksheetPtr pSheet,psTS,psBB,psDMVL,psDMCV,psDMGD,psVLNV,psDMHS,psNDNK;
	RangePtr pRange,prTS,prBB,prDMVL,prDMCV,prDMGD,prVLNV,prDMHS,prNDNK,PRS,PRS2;
	_bstr_t _code,shTS,shBB,shDMVL,shDMCV,shDMGD,shVLNV,shDMHS,shNDNK;
	CConnection ObjConn,ObjConnCV;

	CString msg;
	int _colnv[30],_colvl[30],_colcv[30],_colgd[30],_colhs[30];
	
	int ikey;
	int numNgay;
	CString _ngaythang;
	CString _xnk[5000];
	int _numNK,nDbo,_ivtr[8];	
	int slvchon;
	int iVtchon[5000];
	int slMaTC;
	CString sKTMaTC[100];
	CString sKTTenTC[100];

	int iCheckTenTT;
	CString szfDMNhancong, szfTNMay;
	CString szconDMNC[100], szconTNMay[100];

	CString sFullPath, sDDanfile;
	CString sFiledulieu[100];
	CString sFpathdulieu[100];
	int slfileDL;

public:
	void f_truyen_dulieu_VLNV();
	void f_xacdinh_cot_VLNV();
	void f_xacdinh_cot_DMVL();
	void f_xacdinh_cot_DMCV();
	void f_xacdinh_cot_DMGD();
	void f_xacdinh_cot_DMHS();
	void f_VLNV_Laymaungang(int irow, int icol, int gan, int num);
	void f_autoFilter_VLNV();
	void f_load_stt_bienban();
	void f_SetRowHight(int _nRow, int _ibd, int _ikt);
	void f_xuat_nhat_ky();
	void f_load_csdl_nky(CString sngay,int iRend);	
	void f_chekbox1(CString v1,int n1,CString v2,int n2,CString v3,int n3);
	void f_XNKY_5(int iRend);
	void f_XNKY_4(int iRend);
	void f_XNKY_3(int iRend);
	void f_XNKY_2(int iRend);
	void f_XNKY_1(int iRend);
	void f_capnhat_stt();
	void f_importspic();
	void f_dinh_dang_khung(RangePtr pRg,int itype,int itype2);
	void f_implinkpicsheet();
	void f_stt(_bstr_t fSheet,int ibegin);
	void f_xacdinh_DMHS(_bstr_t fSheet,int col1,int ibegin,
		int vt1,int vt2,int vt3,int vt4,int vt5,int vt6,int vt7,int vt8,int vt9,int ichen);
	void f_truyen_dulieu_DMHS();
	void auto_date_bienban();
	void f_FindBrowse();
	CString F3so(CString vl, int num);
	long f_Read_file_CSV(CString pathCSV);
	long f_GetLineFile(CString szPath);
	CString f_load_congviec(int _numcv,_bstr_t _bb);
	CString f_tachcv(CString noidung,int n);
	CString f_tachklg(CString kluong,int n);
	void f_Load_file_vatlieu();
	void f_Load_file_congviec();
	void f_xacdinh_danhmuc();
	void f_dong_bo_danhmuc(int _vtri);
	void f_even_enter();
	void f_keyboard();
	void f_luu_noidung();
	void f_AddImage_nky();
	void f_load_STT_NCM(CString sNgayNK,int num1,int num2,int cot,int cong,CString frTable);
	void f_SQL_Add_NC_MTC(CString sNgayNK,int num1,int num2,int cot,int cong,CString frTable);
	void f_load_ndung(CString sNgayNK,int num1,int num2,int cot,int cong,CString frTable,int type,int type2,int iRend);
	void f_SQL_AddNdung(CString sNgayNK,int num1,int num2,int cot,int cong,CString frTable,int cotcheck,int iUpd);	
	void f_AutoMerge(_WorksheetPtr pSh1,int iRw,int iCl1,int iCl2);	
	int f_Create_table(_WorksheetPtr pSh1, int iCopy, int iColE);
	void f_Link_manmonth(_WorksheetPtr pSh1,_WorksheetPtr pSh2,_WorksheetPtr pSh3);
	void f_Delete_table(_bstr_t num);
	void f_select_gtri();
	void f_move_bieugia(_WorksheetPtr psh1,int itype);
	void f_move_bangvattu(_WorksheetPtr psh1,int itype);
	void Enter_click_KLG();
	void autofill_TL();
	void auto_tiendo();
	void f_enter_giaidoan();
	void _getDinhmucNhancong(CString szMaCV, CString szLLk);
	void f_Danh_sachfile();
	int f_Check_tieuchuan(CString sKtraTC);
	CString ReplateChamphay(CString sNoidung);
};


