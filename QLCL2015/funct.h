﻿#include "stdafx.h"

CString _fGPathFolder(CString szFile);
bool FileExists(char* filename);
int connectDsn(CString pathFile);
char* _ConvertCstringToChars(CString szvalue);
int _FileExists(CString pathFile, int itype);
CString _TackTenFile(CString spathURL, int itype);
_bstr_t name_sheet(_bstr_t code_name);
wchar_t *tachkytu_(wchar_t s[],int n);
wchar_t *tachchuoi_(wchar_t s[],int n,int vt);
CString getVCell(_WorksheetPtr pSheet, CString name);
CString getVRange(_WorksheetPtr pSheet, CString name);
CString GAR(int nRow,int nCol,RangePtr pRng,int st);
CString GIOR(int nRow,int nCol,RangePtr pRng,CString fkey);
int FindComment(_WorksheetPtr pSheet, int col, _bstr_t comment);
int FindCEnd(_WorksheetPtr pSheet, int col, _bstr_t comment, int iStart);
int FindValue(_WorksheetPtr pSheet, int col, int rowbd, int irowkt, _bstr_t bvalue);
int getColumn(_WorksheetPtr pSheet, CString name);
int getRow(_WorksheetPtr pSheet, CString name);
void putVal(_WorksheetPtr pSheet,_bstr_t name,_variant_t sVal);
int compare_AZ (const void * a, const void * b);
int compare_date(int numDay,CString date1,CString date2);
void f_Setup_Print(_WorksheetPtr pSh1,int iType);
CString fGetRowsSLS(CString _cRows);
int _fGetSelection(CString _cVal);
void _TackTukhoa(CString sKey, CString sChar1, CString sChar2);
void f_proper_string(RangePtr pRg1,int nR,int nC);
double fRoundDouble(double doValue, int nPrecision);
CString FormatTien(CString ftien,CString fCach,CString fSTP);
void f_ktra_date(CString sDate, RangePtr _pRg, int nRow, int nCol);
int f_Check_printer(_bstr_t _prt);
void f_call_moduleXLA(CString sfModule, CString sfSub);
void f_SavePDF(CString sPth,CString sPRg);
void f_Auto_stt_dmuc(RangePtr pRg1,int iDanhstt,int iType,int numbt,int numkt,int icot1,int icot2,int icot4);
void f_Sort_danhmuc(CString sNSheet,CString sRsort,CString sPRg);
int _formatPixcel(RangePtr pRg1, int cot);
bool GetImageSize(const char *fn, int *x,int *y);
void _fGetImage(CString sPath);
int f_Check_ban_quyen();
int f_Mod_check();
void f_Sort_stt_qc();
void xl_calculator_dgkl(RangePtr pRg1,int dongAC,int cotAC,int cotLG);
int FindMulti(CString sChuoi, CString sFind);
void f_Khoiphuc_name(CString sname);
CString fThuMay(CString sNgay);
COLORREF fRandomRGB();
CString _NameColumnHeading(int itype, int col, CString sname);
void _TackTukhoaChuan(CString sKeyValue, CString symbol1, CString symbol2, int itype1, int itype2);
void CopyDulieu(CString sNoidung);
CString f_read_txt(CString sFiledl);
void f_write_txt(CString sChuoi,CString sFiledl);
void f_Get_All_File(CString m_sBaseFolder, CString m_sFileMask);
int f_CheckFile(CString pFile);
CString _ReplateUTF8toWeb(CString szLink);
CString _ReplateWebtoUTF8(CString szLink);
CString _ConvertKhongDau(CString sKeyValue);
void f_SearchGoogle(CString sWebSearch, CString sTenfile, int itype);
void _xlSetHyperlink(_WorksheetPtr pSheet, RangePtr PRS,
	CString pathFile, CString szName, COLORREF clrBkg, COLORREF clrTxt,
	CString szFontName, int iSize, bool bCenter, bool bFormula);
CString f_Get_DownloadTieuchuan(CString szTenTC);
CString f_DocFileTC(CString sTenfile, int iOpen);
char *_ConvertToChar(const CString &s);
CString _ReadingWeb(CString sLinkW);
SOCKET _ConnectToServer(char *szServerName, WORD portNum);
char* _ReadURL2(CString szUrl, long &bytesReturnedOut);
CString _ConvertCharsToCstring(char* chr);
CString _xlConvertA1toRC(CString sValueA1, int itype);
int _DayValue(int year, int month, int day);
int _NumDay2(CString date1, CString date2, int cong);
int _FileExistsUnicode(CString pathFile);
int _xlGetAllHyperLink(RangePtr pRgSelect, vector<CString> &vecHyper);

// Gọi hộp thoại quy cách lấy mẫu ngang
void OQCLMAU();	