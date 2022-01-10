#include "stdafx.h"

HWND hWndViewBmau;
int dPosX = 0, dPosY = 0, dWidthF = 100, dHeightF = 100;

// Biến cơ sở dữ liệu ACCESS
CDatabase Database;
CString SqlString,sDsn;
CString pathMDB=L"",pathNK=L"",pathQCLM=L"",pathCongviec=L"";

_WorksheetPtr pShDMBMau;
int iRowBMau=0, iColBMau=0;
int iRowDLGoc=1,iColumnDLGoc=5;
int jCheckViewBM = 1;

// Bắt lỗi/ ngắt preview
int ClickEsc=0;
int erp,iDlog,iSetup=0;

// In hình ảnh đính kèm
int _idkem=1,_iptrang=2 ;

// Biến hình ảnh
int _imulti=10,_ipoints=75,_iwidth=605,_iheight=160;

// Kích thước thực của ảnh
int iWThuc=400,iHThuc=160;

// Sắp xếp in
int nccBBin=-1;
int isxin[3]={0,0,0};
int isxbs=0,ixkoin=0;
CString SXI_KO_cn[100],SXI_BS_cn[100],SXI_BS_name[100];
CString SXI_VL_cn[30],SXI_CV_cn[30],SXI_GD_cn[30];
CString SXI_VL_name[30],SXI_CV_name[30],SXI_GD_name[30];

// Lấy mẫu bê tông
CString schuoiMTN=L"";

// Biến in lẻ công việc
int jInle=0;

// Quản lý workbooks
CEdit qlwb_edit;
CListCtrlEx qlwb_list;

// Biến EXCEL
_ApplicationPtr xl;
_WorkbookPtr pWb;
PageSetupPtr _psp;

_WorksheetPtr psTS;
RangePtr prTS;
_bstr_t shTS;

CFileFinder _dsfile;

int iCheckTimerDM=1;

// Biến khôi phục name bị mất
int iKprow=0,iKpcol=0;
CString szKpname=L"";
CString keyLMTN=L"LM";
CString sIDmau=L"N00";

// path ảnh
CString sFolder = L"";
CString sNKLuu=L"";

// Vị trí spin
int _RSpin=5,_CSpin=29;

int iCbbThemtc=0;

// Biến đánh lại stt
int mnCphai=0;
int iAtoSTT=1,iKieuSTT=-1,iKieuST2=-1;
int iSapxepIN=0;

// Biến kiểm tra ngày tháng
int ichk7=0,ichkCN=1,ichkLe=1,ichkCmt=1;
COLORREF clrKT_Txt = WHITE,clrKT_Bg=DRED;

// Biến thiết lập tùy chọn trang in
MyPrint _myPrint;

// Biến danh mục hồ sơ
MyQSort Hoso[5000];

CString _CtrTiento[5000],_CtrHauto[5000];

// Biến đồng bộ danh mục
MyDongBo DBO[5000];

// Vật liệu nhập về
MyVLNV _MyVL[5000];

// Biến imppic
CString _bienpic;
int _ihidepic=0;
int _inextpic=0;

// Biến thay đổi kích thước
int _tctc0 = 150;
int _tctc1 = 500;
int _tctc2 = 150;
int _prt0 = 40;
int _prt1 = 80;
int _prt2 = 220;

// Biến selection
CString _arrSLS[1000];

// Biến lưu trạng thái cột tra cứu công việc
int _wL1[2]={80,600};
int _wL2[2]={0,360};
int _wL3[2]={235,200};
int _wLBMAU[4]={120,300,80,80};

// Biến tra cứu vật liệu
int _wV1[3]={80,400,300};
int _wV2[2]={0,360};
int _wV3[4]={0,280,70,70};
int _iTimeDlg=0;
CEdit m_timeDlg[3];

// Biến phân biệt gọi tiêu chuẩn từ hàm nào? (Vật liệu hay công việc)
int _pcaltc=0;

// Biến sắp xếp ngày danh mục hồ sơ
int isktrHS=0,ichkpnhom=1,istang=1,iscbb=0;

// Biến lưu trạng thái tiêu chuẩn
int jLuuDGKL=1,jTTTC=0,_wL4[3]={150,500,150};
int jNhomTC=0,qThemTC=0;

// Biến lưu trạng thái mẫu vật liệu
int jTMCV=0,jTMVL=0,_wL5[4]={80,500,80,80};

// Biến danh mục công việc
CListCtrlEx lit_dmcv,lit_tccv,lit_ndkt;
CColorEdit txtE_DMCV,txtE_TCCV,txtE_NDCV,txtE_NKTC;

// Biến danh mục vật liệu
CListCtrlEx lit_dmvl,lit_tcvl,lit_ppvl;
CColorEdit txtE_DMVL,txtE_TCVL,txtE_PPVL;

// Edit Name
CColorEdit m_editName;
CListCtrlEx m_listName;

// Nhật ký
CListCtrlEx m_listNDNK;

// Biến lưu ngày tháng biên bản
CString _bNThang=L"";

// g0,g1 = giờ - g2,g3 = phút
int gTime[4]={-1,-1,-1,-1};

// Biến edit text list control
int CLRow,CLCol;

// Kết quả tìm kiếm
int iTestnull=0;
CString xl_tukhoa=L"";

int iKqua = 0;
int iKey = 0;
CString pKey[200];
CString pKLMBT[6];
CString sTukhoa=L"";
CString sLuuDownload=L"";
int iCheckDownTC=0;
int iSLTang=5;

CString sColumnA1[27] = {L"_",L"A",L"B",L"C",L"D",L"E",L"F",L"G",L"H",L"I",L"J",
	L"K",L"L",L"M",L"N",L"O",L"P",L"Q",L"R",L"S",L"T",
	L"U",L"V",L"W",L"X",L"Y",L"Z"};

// Các chữ cái (23 chữ, bỏ I,V,X là các số la mã)
CString cSabc[23]={L"A",L"B",L"C",L"D",L"E",L"F",L"G",L"H",L"J",L"K",L"L",L"M",
	L"N",L"O",L"P",L"Q",L"R",L"S",L"T",L"U",L"W",L"Y",L"Z"};

// Các số La mã
CString cSpr[20]={L"I",L"II",L"III",L"IV",L"V",L"VI",L"VII",L"VIII",L"IX",L"X",
	L"XI",L"XII",L"XIII",L"XIV",L"XV",L"XVI",L"XVII",L"XVIII",L"XIX",L"XX"};

CString scvtUTF[134] = {
	L"à",L"á",L"ả",L"ã",L"ạ",L"ă",L"ằ",L"ắ",L"ẳ",L"ẵ",L"ặ",L"â",L"ầ",L"ấ",L"ẩ",L"ẫ",L"ậ",
	L"À",L"Á",L"Ả",L"Ã",L"Ạ",L"Ă",L"Ằ",L"Ắ",L"Ẳ",L"Ẵ",L"Ặ",L"Â",L"Ầ",L"Ấ",L"Ẩ",L"Ẫ",L"Ậ",
	L"ò",L"ó",L"ỏ",L"õ",L"ọ",L"ô",L"ố",L"ồ",L"ổ",L"ỗ",L"ộ",L"ơ",L"ờ",L"ớ",L"ở",L"ỡ",L"ợ",
	L"Ò",L"Ó",L"Ỏ",L"Õ",L"Ọ",L"Ô",L"Ồ",L"Ố",L"Ổ",L"Ỗ",L"Ộ",L"Ơ",L"Ờ",L"Ớ",L"Ở",L"Ỡ",L"Ợ",
	L"è",L"é",L"ẻ",L"ẽ",L"ẹ",L"ê",L"ề",L"ế",L"ể",L"ễ",L"ệ",
	L"È",L"É",L"Ẻ",L"Ẽ",L"Ẹ",L"Ê",L"Ề",L"Ế",L"Ể",L"Ễ",L"Ệ",
	L"ù",L"ú",L"ủ",L"ũ",L"ụ",L"ư",L"ừ",L"ứ",L"ử",L"ữ",L"ự",
	L"Ù",L"Ú",L"Ủ",L"Ũ",L"Ụ",L"Ư",L"Ừ",L"Ứ",L"Ử",L"Ữ",L"Ự",
	L"ì",L"í",L"ỉ",L"ĩ",L"ị",L"ỳ",L"ý",L"ỷ",L"ỹ",L"ỵ",L"đ",
	L"Ì",L"Í",L"Ỉ",L"Ĩ",L"Ị",L"Ỳ",L"Ý",L"Ỷ",L"Ỹ",L"Ỵ",L"Đ"};

CString scvtKOD[134] = {
	L"a",L"a",L"a",L"a",L"a",L"a",L"a",L"a",L"a",L"a",L"a",L"a",L"a",L"a",L"a",L"a",L"a",
	L"A",L"A",L"A",L"A",L"A",L"A",L"A",L"A",L"A",L"A",L"A",L"A",L"A",L"A",L"A",L"A",L"A",
	L"o",L"o",L"o",L"o",L"o",L"o",L"o",L"o",L"o",L"o",L"o",L"o",L"o",L"o",L"o",L"o",L"o",
	L"O",L"O",L"O",L"O",L"O",L"O",L"O",L"O",L"O",L"O",L"O",L"O",L"O",L"O",L"O",L"O",L"O",
	L"e",L"e",L"e",L"e",L"e",L"e",L"e",L"e",L"e",L"e",L"e",
	L"E",L"E",L"E",L"E",L"E",L"E",L"E",L"E",L"E",L"E",L"E",
	L"u",L"u",L"u",L"u",L"u",L"u",L"u",L"u",L"u",L"u",L"u",
	L"U",L"U",L"U",L"U",L"U",L"U",L"U",L"U",L"U",L"U",L"U",
	L"i",L"i",L"i",L"i",L"i",L"y",L"y",L"y",L"y",L"y",L"d",
	L"I",L"I",L"I",L"I",L"I",L"Y",L"Y",L"Y",L"Y",L"Y",L"D" };

CString scharutf8[152] = {L"%",L" ",
	L"&",L"=",L"?",L"$",L"#",L"^",L"@",L";",
	L"+",L":",L"/",L"÷",L"[",L"]",L"|",L"\"",
	L"à",L"á",L"ả",L"ã",L"ạ",L"ă",L"ằ",L"ắ",L"ẳ",L"ẵ",L"ặ",L"â",L"ầ",L"ấ",L"ẩ",L"ẫ",L"ậ",
	L"À",L"Á",L"Ả",L"Ã",L"Ạ",L"Ă",L"Ằ",L"Ắ",L"Ẳ",L"Ẵ",L"Ặ",L"Â",L"Ầ",L"Ấ",L"Ẩ",L"Ẫ",L"Ậ",
	L"ò",L"ó",L"ỏ",L"õ",L"ọ",L"ô",L"ố",L"ồ",L"ổ",L"ỗ",L"ộ",L"ơ",L"ờ",L"ớ",L"ở",L"ỡ",L"ợ",
	L"Ò",L"Ó",L"Ỏ",L"Õ",L"Ọ",L"Ô",L"Ồ",L"Ố",L"Ổ",L"Ỗ",L"Ộ",L"Ơ",L"Ờ",L"Ớ",L"Ở",L"Ỡ",L"Ợ",
	L"è",L"é",L"ẻ",L"ẽ",L"ẹ",L"ê",L"ề",L"ế",L"ể",L"ễ",L"ệ",L"È",L"É",L"Ẻ",L"Ẽ",L"Ẹ",L"Ê",L"Ề",L"Ế",L"Ể",L"Ễ",L"Ệ",
	L"ù",L"ú",L"ủ",L"ũ",L"ụ",L"ư",L"ừ",L"ứ",L"ử",L"ữ",L"ự",L"Ù",L"Ú",L"Ủ",L"Ũ",L"Ụ",L"Ư",L"Ừ",L"Ứ",L"Ử",L"Ữ",L"Ự",
	L"ì",L"í",L"ỉ",L"ĩ",L"ị",L"Ì",L"Í",L"Ỉ",L"Ĩ",L"Ị",
	L"ỳ",L"ý",L"ỷ",L"ỹ",L"ỵ",L"Ỳ",L"Ý",L"Ỷ",L"Ỹ",L"Ỵ",L"đ",L"Đ"};

CString scharweb[152] = {L"%25",L"%20",
	L"%26",L"%3D",L"%3F",L"%24",L"%23",L"%5E",L"%40",L"%3B",
	L"%2B",L"%3A",L"%2F",L"%C3%B7",L"%5B",L"%5D",L"%7C",L"%22",
	L"%C3%A0",L"%C3%A1",L"%E1%BA%A3",L"%C3%A3",L"%E1%BA%A1",L"%C4%83",L"%E1%BA%B1",L"%E1%BA%AF",
	L"%E1%BA%B3",L"%E1%BA%B5",L"%E1%BA%B7",L"%C3%A2",L"%E1%BA%A7",L"%E1%BA%A5",L"%E1%BA%A9",L"%E1%BA%AB",L"%E1%BA%AD",
	L"%C3%A0",L"%C3%A1",L"%E1%BA%A3",L"%C3%A3",L"%E1%BA%A1",L"%C4%83",L"%E1%BA%B1",L"%E1%BA%AF",L"%E1%BA%B3",L"%E1%BA%B5",
	L"%E1%BA%B7",L"%C3%A2",L"%E1%BA%A7",L"%E1%BA%A5",L"%E1%BA%A9",L"%E1%BA%AB",L"%E1%BA%AD",L"%C3%B2",L"%C3%B3",L"%E1%BB%8F",
	L"%C3%B5",L"%E1%BB%8D",L"%C3%B4",L"%E1%BB%91",L"%E1%BB%93",L"%E1%BB%95",L"%E1%BB%97",L"%E1%BB%99",L"%C6%A1",L"%E1%BB%9D",
	L"%E1%BB%9B",L"%E1%BB%9F",L"%E1%BB%A1",L"%E1%BB%A3",L"%C3%B2",L"%C3%B3",L"%E1%BB%8F",L"%C3%B5",L"%E1%BB%8D",L"%C3%B4",
	L"%E1%BB%91",L"%E1%BB%93",L"%E1%BB%95",L"%E1%BB%97",L"%E1%BB%99",L"%C6%A1",L"%E1%BB%9D",L"%E1%BB%9B",L"%E1%BB%9F",
	L"%E1%BB%A1",L"%E1%BB%A3",L"%C3%A8",L"%C3%A9",L"%E1%BA%BB",L"%E1%BA%BD",L"%E1%BA%B9",L"%C3%AA",L"%E1%BB%81",L"%E1%BA%BF",
	L"%E1%BB%83",L"%E1%BB%85",L"%E1%BB%87",L"%C3%A8",L"%C3%A9",L"%E1%BA%BB",L"%E1%BA%BD",L"%E1%BA%B9",L"%C3%AA",L"%E1%BB%81",
	L"%E1%BA%BF",L"%E1%BB%83",L"%E1%BB%85",L"%E1%BB%87",L"%C3%B9",L"%C3%BA",L"%E1%BB%A7",L"%C5%A9",L"%E1%BB%A5",L"%C6%B0",
	L"%E1%BB%AB",L"%E1%BB%A9",L"%E1%BB%AD",L"%E1%BB%AF",L"%E1%BB%B1",L"%C3%B9",L"%C3%BA",L"%E1%BB%A7",L"%C5%A9",L"%E1%BB%A5",
	L"%C6%B0",L"%E1%BB%AB",L"%E1%BB%A9",L"%E1%BB%AD",L"%E1%BB%AF",L"%E1%BB%B1",L"%C3%AC",L"%C3%AD",L"%E1%BB%89",L"%C4%A9",
	L"%E1%BB%8B",L"%C3%AC",L"%C3%AD",L"%E1%BB%89",L"%C4%A9",L"%E1%BB%8B",L"%E1%BB%B3",L"%C3%BD",L"%E1%BB%B7",L"%E1%BB%B9",
	L"%E1%BB%B5",L"%E1%BB%B3",L"%C3%BD",L"%E1%BB%B7",L"%E1%BB%B9",L"%E1%BB%B5",L"%C4%91",L"%C4%91"};

CString scharutf8_2[18] = {L"%",L" ",
	L"&",L"=",L"?",L"$",L"#",L"^",L"@",L";",
	L"+",L":",L"/",L"÷",L"[",L"]",L"|",L"\""};

CString scharweb_2[18] = {L"%25",L"%20",
	L"%26",L"%3D",L"%3F",L"%24",L"%23",L"%5E",L"%40",L"%3B",
	L"%2B",L"%3A",L"%2F",L"%C3%B7",L"%5B",L"%5D",L"%7C",L"%22"};

// Tìm kiếm Excel
TKEXL XLXD[5000];

// Biến mở rộng sheet
int slbsThuc=0,slInbs=0,slcpy[4]={0,0,0,0};
int ctInbs[1000];
_bstr_t csInbs[1000],csInnm[1000];
MyCopySheet MCS[4];

// Cho phép sử dụng phím tắt
int iHotk=-1, iATdate=-1;

// Biến di chuyển biểu giá
CString sbgHM=L"",sbgGD=L"";
int ibgTG=1,ibgTM=1,ibgTC=1,ibgDMNC=1;

// Biến quản lý in ảnh
int grPic=0;
MyPic Mpic[100];

// Biến lưu tích chọn
int iLuuDMCVCon=1,iLuuTCCV=1,iLuuTCVL=1;

// Biến quản lý sheet
MyWBk MyWb[500];
CString MySymbol[6]={L"*",L"/",L"\\",L":",L"?",L"'"};
CString MySymABC[36]={L"a",L"b",L"c",L"d",L"e",L"f",L"g",L"h",L"i",L"j",
	L"k",L"l",L"m",L"n",L"o",L"p",L"q",L"r",L"s",L"t",L"u",L"v",L"w",L"x",L"y",L"z",
	L"0",L"1",L"2",L"3",L"4",L"5",L"6",L"7",L"8",L"9"};

CString sColABC[52]={L"a",L"b",L"c",L"d",L"e",L"f",L"g",L"h",L"i",L"j",
	L"k",L"l",L"m",L"n",L"o",L"p",L"q",L"r",L"s",L"t",L"u",L"v",L"w",L"x",L"y",L"z",
	L"aa",L"ab",L"ac",L"ad",L"ae",L"af",L"ag",L"ah",L"ai",L"aj",
		L"ak",L"al",L"am",L"an",L"ao",L"ap",L"aq",L"ar",L"as",L"at",L"au",L"av",L"aw",L"ax",L"ay",L"az"};

// Bổ sung lưu hồ sơ (23.06.2016)
int iLuuType=0,luuhsCB1=0,luuhsCB2=0;
int slcv_VL=0,slcv_CV=0,slcv_GD=0;
int iLuuChk[9];
int inLMVL[5000];
int inLMCV[5000];
int inKTBT[5000];
int iLuuKEY,iLuuMDP;
int iCountCB2;
CString iCViec[1000];

// Check bản quyền
int count_key=0;

// Trạng thái ẩn hiện của sheet
int slSheet=0;
int iHideS[200];

int chkClrCV=1,chkClrVL=1;
COLORREF fclrCV = RGB(204,255,204),fclrVL = RGB(204,255,204);

// Nhật ký
int _iNgay=10;
CString ngbd,ngkt,ngay1,ngay2;

// Quy cách lấy mẫu
CListCtrlEx listQclm,litLMN3,litCRETD;
CColorEdit edTxt,edLMN,edCreateTxt;
int irecH=465,irecW=804,irecH2=0,irecW2=0;
int irH2=400,irW2=800;
int irHCV=0,irWCV=0,irHVL=0,irWVL=0,irHBMAU=0,irWBMAU=0;
int jW1[5]={40,400,100,200,80};
int jW2[8]={0,250,150,200,100,100,100,50};
int jW3[20]={0,40,100,100,150,150,150,150,150,150,150,150,150,150,150,150,150,150,0,0};
int iChekUp=0;

// Vị trí Man Month
_WorksheetPtr wsTHCPtv;
int iVMMn=1,iMMRow=1,iMMCol=1;
int _mmvt[6]={0,0,0,0,0,0};

// Sheet bổ sung từ dự toán
_bstr_t bsActivate = (_bstr_t)"";
int ivtdchuyen=0;

// Hàm thông báo kiểu số
void _d(int num, int itype)
{
	CString val = L"";
	val.Format(L"Num=%d", num);
	if (itype == 1) xl->PutScreenUpdating(VARIANT_TRUE);
	MessageBox(NULL, val, L"www.giaxaydung.vn", MB_OK | MB_ICONINFORMATION);
	if (itype == 1) xl->PutScreenUpdating(VARIANT_FALSE);
}

// Hàm thông báo kiểu chuỗi
void _s(CString val, int itype)
{
	if (itype == 1) xl->PutScreenUpdating(VARIANT_TRUE);
	MessageBox(NULL, val, L"www.giaxaydung.vn", MB_OK | MB_ICONINFORMATION);
	if (itype == 1) xl->PutScreenUpdating(VARIANT_FALSE);
}

// Hàm thông báo kết hợp kiểu chuỗi & số
void _n(CString sval, int num, int itype)
{
	CString val = L"";
	val.Format(L"%s=%d", sval, num);
	val.Replace(L"==", L"=");
	if (itype == 1) xl->PutScreenUpdating(VARIANT_TRUE);
	MessageBox(NULL, val, L"www.giaxaydung.vn", MB_OK | MB_ICONINFORMATION);
	if (itype == 1) xl->PutScreenUpdating(VARIANT_FALSE);
}


int _q(CString sVal, int num, int itype)
{
	CString val = sVal;
	if (num > 0) val.Format(L"%s = %d", sVal, num);
	if (itype == 1) xl->PutScreenUpdating(VARIANT_TRUE);
	int n = MessageBox(NULL, val, L"Thông báo", MB_YESNO | MB_ICONQUESTION);
	if (itype == 1) xl->PutScreenUpdating(VARIANT_FALSE);
	if (n == 7) return 1;
	return 0;
}

int _yn(CString sval, int nStyle, int itype)
{
	int n = 6;
	if (itype == 1) xl->PutScreenUpdating(VARIANT_TRUE);
	if (nStyle == 0) n = MessageBox(NULL, sval, L"Thông báo", MB_YESNO | MB_ICONQUESTION);
	else n = MessageBox(NULL, sval, L"Thông báo", MB_YESNOCANCEL | MB_ICONQUESTION);
	if (itype == 1) xl->PutScreenUpdating(VARIANT_FALSE);
	return n;
}


void xl_Create()
{
	CoInitialize(NULL);
	xl.GetActiveObject(L"Excel.Application");
	xl->Visible = true;
	xl->EnableCancelKey = XlEnableCancelKey(FALSE);
	pWb = xl->GetActiveWorkbook();
}

CString getPathCV()
{
	try
	{
		shTS = name_sheet("shTS");
		psTS = xl->Sheets->GetItem(shTS);
		prTS = psTS->GetCells();

		CString szval=L"";
		for (int i = 1; i < 100; i++)
		{
			szval = GIOR(iRowDLGoc+i,iColumnDLGoc+3,prTS,L"");
			if(_wtoi(szval)==1)
			{
				szval = GIOR(iRowDLGoc+i,iColumnDLGoc+2,prTS,L"");
				if(szval!=L"")
				{
					if(szval.Find(L"\\")>0) pathCongviec = szval;
					else
					{
						CString spath = GIOR(iRowDLGoc,iColumnDLGoc,prTS,L"");
						CString szfolder = GIOR(iRowDLGoc+1,iColumnDLGoc,prTS,L"");
						
						if(spath.Right(1)!=L"\\") spath+=L"\\";
						pathCongviec.Format(L"%s%s\\%s",spath,szfolder,szval);
					}

					break;
				}
			}
		}

		return pathCongviec;
	}
	catch(...)
	{
		_s(L"Lỗi đọc đường dẫn công việc, bạn vui lòng chọn lại đường dẫn cơ sở dữ liệu");
		return L"";
	}

	return L"";
}

CString getPathQLCL()
{
	try
	{
		shTS = name_sheet("shTS");
		psTS = xl->Sheets->GetItem(shTS);
		prTS = psTS->GetCells();

		CString spath = GIOR(iRowDLGoc,iColumnDLGoc,prTS,L"");
		CString szfolder = GIOR(iRowDLGoc+2,iColumnDLGoc,prTS,L"");
		CString sfile = GIOR(iRowDLGoc+2,iColumnDLGoc+1,prTS,L"");

		if(spath.Right(1)!=L"\\") spath+=L"\\";
		pathMDB.Format(L"%s%s\\%s",spath,szfolder,sfile);

		return pathMDB;
	}
	catch(...)
	{
		_s(L"Lỗi đọc đường dẫn gốc, bạn vui lòng chọn lại đường dẫn cơ sở dữ liệu");
		return L"";
	}

	return L"";
}

CString getPathQCLM()
{
	try
	{
		shTS = name_sheet("shTS");
		psTS = xl->Sheets->GetItem(shTS);
		prTS = psTS->GetCells();

		CString spath = GIOR(iRowDLGoc,iColumnDLGoc,prTS,L"");
		CString szfolder = GIOR(iRowDLGoc+3,iColumnDLGoc,prTS,L"");
		CString sfile = GIOR(iRowDLGoc+3,iColumnDLGoc+1,prTS,L"");

		if(spath.Right(1)!=L"\\") spath+=L"\\";
		pathQCLM.Format(L"%s%s\\%s",spath,szfolder,sfile);

		return pathQCLM;
	}
	catch(...)
	{
		_s(L"Lỗi đọc đường dẫn lấy mẫu, bạn vui lòng chọn lại đường dẫn cơ sở dữ liệu");
		return L"";
	}

	return L"";
}

CString getPathNHKY()
{
	try
	{
		shTS = name_sheet("shTS");
		psTS = xl->Sheets->GetItem(shTS);
		prTS = psTS->GetCells();

		CString spath = GIOR(iRowDLGoc,iColumnDLGoc,prTS,L"");
		CString szfolder = GIOR(iRowDLGoc+4,iColumnDLGoc,prTS,L"");
		CString sfile = GIOR(iRowDLGoc+4,iColumnDLGoc+1,prTS,L"");

		if(spath.Right(1)!=L"\\") spath+=L"\\";
		pathNK.Format(L"%s%s\\%s",spath,szfolder,sfile);

		return pathNK;
	}
	catch(...)
	{
		_s(L"Lỗi đọc đường dẫn nhật ký, bạn vui lòng chọn lại đường dẫn cơ sở dữ liệu");
		return L"";
	}

	return L"";
}

CString getDesktopDir()
{
	CString szDesktop=L"";
	char* szpath = new char[MAX_PATH+1];
	if(SHGetSpecialFolderPathA(HWND_DESKTOP,szpath,CSIDL_DESKTOP,FALSE))
	{
		szDesktop = _ConvertCharsToCstring(szpath);
		if(szDesktop.Right(1)!=L"\\") szDesktop+=L"\\";
	}
	else szDesktop = L"";

	return szDesktop;
}

void CheckNhomKy(_WorksheetPtr pSh1, int iRow, int cotKyBB)
{
	try
	{
		_bstr_t sh = pSh1->CodeName;
		RangePtr pRg1 = pSh1->Cells;

		int ngay[5]={0,0,0,0,0};
		if(sh==(_bstr_t)L"shHSNTVL")
		{
			ngay[0] = getColumn(pSh1,"DMVL_NK_NG");
			ngay[1] = getColumn(pSh1,"DMVL_LM_NG");
			ngay[2] = getColumn(pSh1,"DMVL_NTNB_NG");
			ngay[3] = getColumn(pSh1,"DMVL_YC");
			ngay[4] = getColumn(pSh1,"DMVL_AB_NG");
		}
		else if(sh==(_bstr_t)L"shHSNTCV")
		{
			ngay[0] = getColumn(pSh1,"DMBB_TN_NGAY");
			ngay[1] = getColumn(pSh1,"DMBB_NB_NGAY");
			ngay[2] = getColumn(pSh1,"DMBB_PHIEUYC");
			ngay[3] = getColumn(pSh1,"DMBB_AB_NGAY");
			ngay[4] = ngay[3];
		}
		else if(sh==(_bstr_t)L"shHSNTGD")
		{
			ngay[0] = getColumn(pSh1,"DMGD_NTNB_NG");
			ngay[1] = getColumn(pSh1,"DMGD_YC");
			ngay[2] = getColumn(pSh1,"DMGD_AB_NG");
			ngay[3] = ngay[2];
			ngay[4] = ngay[2];
		}

		CString sNgay=L"";
		sNgay.Format(L"=MIN(%s,%s,%s,%s,%s)",
			GAR(iRow,ngay[0],pRg1,0),GAR(iRow,ngay[1],pRg1,0),
			GAR(iRow,ngay[2],pRg1,0),GAR(iRow,ngay[3],pRg1,0),GAR(iRow,ngay[4],pRg1,0));

		int clLast = (int)pSh1->Cells->SpecialCells(xlCellTypeLastCell)->GetColumn();
		CString szval = GIOR(iRow, clLast, pRg1, L"");
		if (szval != L"") clLast++;

		pRg1->PutItem(iRow, clLast, (_bstr_t)sNgay);
		RangePtr PRS = pRg1->GetItem(iRow, clLast);
		PRS->PutNumberFormat(L"dd/mm/yyyy");
		sNgay = GIOR(iRow, clLast, pRg1, L"");

		PRS->PutNumberFormat(L"General");
		PRS->ClearContents();
		PRS->ClearFormats();

		if(sNgay.Right(2)==L"00" || sNgay.Right(1)==L"M")
		{
			pRg1->PutItem(iRow,cotKyBB,(_bstr_t)L"");
			return;
		}

		// Sheet ký biên bản
		_bstr_t shKyBB = name_sheet((_bstr_t)"Sheet1");
		_WorksheetPtr psKyBB = xl->Sheets->GetItem(shKyBB);
		RangePtr prKyBB = psKyBB->Cells;

		// Kiểm tra nhóm ký biên bản nào có ngày tháng
		int dem=0;
		int iNhomKy=0;
		int iLen = sNgay.GetLength();
		
		for (int i = 3; i <= 100; i++)
		{
			szval = GIOR(1,i,prKyBB,L"");
			szval.Trim();

			if(szval!=L"" && szval.GetLength()==iLen)
			{
				// So sánh ngày ở sheet danh mục (N1) với ngày ở sheet ký biên bản (N2)
				if(compare_date(iLen,szval,sNgay)==1) iNhomKy = i-3;
				else break;
				dem=0;
			}
			else dem++;
			if(dem==10) break;
		}

		if(iNhomKy>0) pRg1->PutItem(iRow,cotKyBB,iNhomKy);
		else pRg1->PutItem(iRow,cotKyBB,(_bstr_t)L"");
	}
	catch(...){}
}
