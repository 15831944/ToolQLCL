#pragma once
#include "stdafx.h"
#include "afxdialogex.h"
#include <afxmsg_.h>
#include <afxdb.h>
#include <stdlib.h> 
#include "msolib.h"
#include <iomanip>
#include <iostream>
#include <string>
#include <vector>
#include <fstream>
#include <algorithm>
#include <sys/stat.h>
#include "resource.h"
#include "easysize.h"
#include "ListCtrlEx.h"
#include "Color.h"
#include "ColorEdit.h"
#include "ColorStatic.h"
#include "ComboBoxExt.h"
#include "ComboBoxExtList.h"
#include "NumSpinCtrl.h"
#include "TextProgressCtrl.h"
#include "FileFinder.h"
#include "FileFilter.h"
#include "ScrollEdit.h"
#include "StrUtil.h"
#include "Connection.h"
#include <windows.h>

#include <wininet.h>
#pragma comment(lib, "wininet.lib")

using namespace std;

#include "funct.h"

#ifndef _SECURE_ATL
#define _SECURE_ATL
#endif

#ifndef VC_EXTRALEAN
#define VC_EXTRALEAN		
#endif


#ifndef WINVER				
#define WINVER 0x0501		
#endif

#ifndef _WIN32_WINNT		                   
#define _WIN32_WINNT 0x0501	
#endif						

#ifndef _WIN32_WINDOWS		
#define _WIN32_WINDOWS 0x0410 
#endif

#ifndef _WIN32_IE			
#define _WIN32_IE 0x0600	
#endif

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS	


#define _AFX_ALL_WARNINGS

#include <afxwin.h>         
#include <afxext.h>         
#include <afxdisp.h>        

#ifndef _AFX_NO_OLE_SUPPORT
#include <afxdtctl.h>		
#endif
#ifndef _AFX_NO_AFXCMN_SUPPORT
#include <afxcmn.h>			
#endif 
#include <afxcontrolbars.h>
#include <afxcontrolbars.h>
#include <afxcontrolbars.h>
#include <afxcontrolbars.h>
#include <afxcontrolbars.h>
#include <afxcontrolbars.h>
#include <afxcontrolbars.h>
#include <afxcontrolbars.h>
#include <afxcontrolbars.h>

#ifdef _UNICODE
#if defined _M_IX86
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='x86' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_IA64
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='ia64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_X64
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='amd64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#else
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#endif
#endif

// Số cột + ký tự dòng trong CSV
#define SIZE_LINE	2000
#define ROW_LINE	2000

// Biến cơ sở dữ liệu ACCESS
#define sDriver			L"Microsoft Access Driver (.mdb, .accdb)"
#define _pathFILE		L"Dulieu\\DulieuQLCL\\"
#define _pathDlieu		L"Dulieu\\"
#define _pathImage		L"_Anhsohoa\\"
#define _pathTVien		L"Thuvien\\"
#define _dFileTC		L"TCsave.ini"
#define _dFileMDB		L"Dulieugoc.mdb"
#define _dFileQCLM		L"Banglaymau.mdb"
#define _dFileNKY		L"Nhatky.mdb"
#define _dFileConfig	L"Config.mdb"
#define _dFileCPTV		L"CPTV.csv"
#define _dFileNle		L"ngayle.csv"
#define _prXPS			L"Microsoft XPS Document Writer"
#define _prPDF			L"Adobe PDF"
#define _duoiXPS		L".xps"
#define _duoiPDF		L".pdf"
#define _rowheight		L"&RH="
#define _dfTiento		L"Gxd.vn-"
#define _dfError		L"Gxd-error"
#define _urlSever		L"3jRaMzBZEGRbMlIxZEc5aGJtZDRaQzUyYmk5a2RXeHBaWFV2ZEdsbGRXTm9kPQ1=TdUi=2oO"
#define _WcheckConn		L"TjRaMzBZIGRTeTkzZDNjdVoyOXZaMnhsajE0j1T3=G9Y"
#define _Wsearch		L"TjRaMzBZIGRTeTkzZDNjdVoyOXZaMnhsTG1OdmJTNTJiaTl6WldGPTBl1keyFVnM"
#define _KeySearch		L"jaRWj3BVYmJUR1U2Wkc5aklFOVNJR1pwYkdWMGVYQmxPPQ1=jFba=U3W"
#define _KeyDoc			L"DRMdMUUhxFRMRgV=n1MQ=23T"
#define _KeyPdf			L"DSRdMlFhsFJMRgV=n1MQ=23T"

#define _dFolderQLCL	L"DulieuQLCL"
#define _dFolderQCLM	L"Banglaymau"
#define _dFolderNHKY	L"Nhatky"

extern HWND hWndViewBmau;
extern int dPosX, dPosY, dWidthF, dHeightF;

// Biến cơ sở dữ liệu ACCESS
extern CDatabase Database;
extern CString SqlString,sDsn;
extern CString pathMDB,pathNK,pathQCLM,pathCongviec;

extern _WorksheetPtr pShDMBMau;
extern int iRowBMau, iColBMau;
extern int iRowDLGoc,iColumnDLGoc;
extern int jCheckViewBM;

// Bắt lỗi/ ngắt preview
extern int ClickEsc;
extern int erp,iDlog,iSetup;

// In hình ảnh đính kèm
extern int _idkem,_iptrang;

// Biến hình ảnh
extern int _imulti,_ipoints,_iwidth,_iheight;

// Kích thước thực của ảnh
extern int iWThuc,iHThuc;

// Sắp xếp in
extern int nccBBin;
extern int isxin[3],isxbs,ixkoin;
extern CString SXI_KO_cn[100],SXI_BS_cn[100],SXI_BS_name[100];
extern CString SXI_VL_cn[30],SXI_CV_cn[30],SXI_GD_cn[30];
extern CString SXI_VL_name[30],SXI_CV_name[30],SXI_GD_name[30];

// Lấy mẫu bê tông
extern CString schuoiMTN;

// Biến in lẻ công việc
extern int jInle;

// Quản lý workbooks
extern CEdit qlwb_edit;
extern CListCtrlEx qlwb_list;

// Biến EXCEL
extern _ApplicationPtr xl;
extern _WorkbookPtr pWb;
extern PageSetupPtr _psp;

extern _WorksheetPtr psTS;
extern RangePtr prTS;
extern _bstr_t shTS;

extern CFileFinder _dsfile;

extern int iCheckTimerDM;

// Biến khôi phục name bị mất
extern int iKprow,iKpcol;
extern CString szKpname;
extern CString keyLMTN;
extern CString sIDmau;

// path ảnh
extern CString sFolder,sNKLuu;

// Vị trí spin
extern int _RSpin,_CSpin;

extern int iCbbThemtc;

// Biến đánh lại stt
extern int mnCphai;
extern int iAtoSTT,iKieuSTT,iKieuST2;
extern int iSapxepIN;

// Biến kiểm tra ngày tháng
extern int ichk7,ichkCN,ichkLe,ichkCmt;
extern COLORREF clrKT_Txt,clrKT_Bg;

// Biến thiết lập tùy chọn trang in
struct MyPrint
{
	int ileft;
	int itop;
	int iright;
	int ibottom;
	int ihtop;
	int ifbot;
	int ihor;
	int iver;
	int isizehd;
	int isizeft;
	int istylehd;
	int istyleft;
	int isechd;
	int isecft;
	int inumpage;
	int ibwhite;
	int ierror;
	int izoom;
	int ifirst;
	int ibreaks;
	int isecpage;
	int inextpage;
	int iprivate;

	CString cstylehd;
	CString cstyleft;
	CString ctxthd;
	CString ctxtft;
};
extern MyPrint _myPrint;

struct MyQSort
{
	CString ival;
	CString dqs[20];
};
extern MyQSort Hoso[5000];

extern CString _CtrTiento[5000],_CtrHauto[5000];

struct MyDongBo
{
	int iNumTC;
	CString sMa;
	CString sTen;
	CString tch[100];
	CString mch[100];
};
extern MyDongBo DBO[5000];

struct MyVLNV
{
	int nDG;
	_bstr_t sMa;
	_bstr_t sTen;
	_bstr_t sXuatxu[1000];
	_bstr_t sNgay[1000];
	_bstr_t sBB[1000];
	_bstr_t sDG[1000];
	_bstr_t sMAC[1000];
	_bstr_t sDVI[1000];
	_bstr_t sKLG[1000];
	_bstr_t sGchu[1000];
};
extern MyVLNV _MyVL[5000];

// Biến imppic
extern CString _bienpic;
extern int _ihidepic,_inextpic;

// Biến thay đổi kích thước
extern int _tctc0,_tctc1,_tctc2;
extern int _prt0,_prt1,_prt2;

// Biến selection
extern CString _arrSLS[1000];

// Biến lưu trạng thái cột tra cứu công việc
extern int _wL1[2],_wL2[2],_wL3[2];
extern int _wV1[3],_wV2[2],_wV3[4];
extern int _wLBMAU[4];

extern int _iTimeDlg;
extern CEdit m_timeDlg[3];

// Biến phân biệt gọi tiêu chuẩn từ hàm nào? (Vật liệu hay công việc)
extern int _pcaltc;

// Biến lưu trạng thái tiêu chuẩn
extern int jLuuDGKL,jTTTC,_wL4[3];
extern int jNhomTC,qThemTC;

// Biến sắp xếp ngày danh mục hồ sơ
extern int isktrHS,ichkpnhom,istang,iscbb;

// Biến lưu trạng thái mẫu vật liệu
extern int jTMCV,jTMVL,_wL5[4];

// Biến danh mục công việc
extern CListCtrlEx lit_dmcv,lit_tccv,lit_ndkt;
extern CColorEdit txtE_DMCV,txtE_TCCV,txtE_NDCV,txtE_NKTC;

// Biến danh mục vật liệu
extern CListCtrlEx lit_dmvl,lit_tcvl,lit_ppvl;
extern CColorEdit txtE_DMVL,txtE_TCVL,txtE_PPVL;

// Edit Name
extern CColorEdit m_editName;
extern CListCtrlEx m_listName;

// Nhật ký
extern CListCtrlEx m_listNDNK;

// Biến lưu ngày tháng biên bản
extern CString _bNThang;
extern int gTime[4]; 

// Biến edit text list control
extern int CLRow,CLCol;

// Kết quả tìm kiếm
extern int iTestnull;
extern CString xl_tukhoa;

extern int iKqua;
extern int iKey;
extern CString pKey[200];
extern CString pKLMBT[6];
extern CString sTukhoa;

extern CString scvtUTF[134];
extern CString scvtKOD[134];
extern CString sLuuDownload;
extern int iCheckDownTC;
extern int iSLTang;

// Decode website
extern CString scharutf8[152],scharweb[152];
extern CString scharutf8_2[18],scharweb_2[18];

// Các ký hiệu cột kiểu A1
extern CString sColumnA1[27];

// Các ký tự đặc biệt
extern CString cSabc[23],cSpr[20];

// Tìm kiếm Excel
struct TKEXL
{
	CString CDG[5];
};
extern TKEXL XLXD[5000];

// Biến mở rộng sheet
extern int slbsThuc,slInbs,slcpy[4],ctInbs[1000];
extern _bstr_t csInbs[1000],csInnm[1000];

struct MyCopySheet
{
	_bstr_t sCN[50];
	_bstr_t sName[50];
	int iSLCV[50];
};
extern MyCopySheet MCS[4];

struct MyWBk
{
	int iNum;
	int iShow;
	_bstr_t sName;
	_bstr_t sCode;
	COLORREF sColor;
	COLORREF cAfter;
};

struct MyPic
{
	CString _spt[5];		// Đường dẫn ảnh thứ (i)
	CString _sNgay;			// Ngày
	CString _sNdung[5];		// Nội dung
	int iWt[5];				// Độ rộng (i)
	int iHt[5];				// Chiều cao (i)
};

// Cho phép sử dụng phím tắt
extern int iHotk,iATdate;

// Biến di chuyển biểu giá
extern CString sbgHM,sbgGD;
extern int ibgTG,ibgTM,ibgTC,ibgDMNC;

// Biến quản lý in ảnh
extern int grPic;
extern MyPic Mpic[100];

// Lưu tích chọn
extern int iLuuDMCVCon,iLuuTCCV,iLuuTCVL;

// Biến quản lý sheet
extern MyWBk MyWb[500];
extern CString MySymbol[6];
extern CString MySymABC[36];
extern CString sColABC[52];

// Bổ sung lưu hồ sơ (23.06.2016)
extern int iLuuType,luuhsCB1,luuhsCB2;
extern int slcv_VL,slcv_CV,slcv_GD;
extern int iLuuChk[9];
extern int inLMVL[5000];
extern int inLMCV[5000];
extern int inKTBT[5000];
extern int iLuuKEY,iLuuMDP;
extern int iCountCB2;
extern int _iNgay;
extern CString iCViec[1000];

// Check bản quyền
extern int count_key;

// Trạng thái ẩn hiện của sheet
extern int slSheet;
extern int iHideS[200];

extern int chkClrCV,chkClrVL;
extern COLORREF fclrCV,fclrVL;

// Nhật ký
extern CString ngbd,ngkt,ngay1,ngay2;

// Quy cách lấy mẫu
extern CListCtrlEx listQclm,litLMN3,litCRETD;
extern CColorEdit edTxt,edLMN,edCreateTxt;
extern int irecH,irecW,irecH2,irecW2;
extern int irH2,irW2;
extern int irHCV,irWCV,irHVL,irWVL,irHBMAU,irWBMAU;
extern int jW1[5],jW2[8],jW3[20];
extern int iChekUp;

// Vị trí Man Month
extern _WorksheetPtr wsTHCPtv;
extern int iVMMn,iMMRow,iMMCol;
extern int _mmvt[6];

// Sheet bổ sung từ dự toán
extern _bstr_t bsActivate;
extern int ivtdchuyen;

void _d(int num, int itype=0);
void _s(CString val, int itype=0);
void _n(CString sval, int num, int itype=0);
int _q(CString sVal, int num, int itype=0);
int _yn(CString sval, int nStyle=0, int itype=0);

void xl_Create();
CString getPathCV();
CString getPathQLCL();
CString getPathQCLM();
CString getPathNHKY();
CString getDesktopDir();
void CheckNhomKy(_WorksheetPtr pSh1, int iRow, int cotKyBB);

inline char* ConvertToChar(const CString &s)
{
  int nSize = s.GetLength();
  char *pAnsiString = new char[nSize+1];
  memset(pAnsiString,0,nSize+1);
  wcstombs(pAnsiString, s, nSize+1);
  return pAnsiString;
}

inline BOOL IsFileExist(CString sPathName)
{
	HANDLE hFile;
	hFile = CreateFile(sPathName, GENERIC_READ, FILE_SHARE_READ, NULL, CREATE_NEW, NULL, NULL);
	if (hFile == INVALID_HANDLE_VALUE) return FALSE;
	CloseHandle(hFile);
	return TRUE;
}