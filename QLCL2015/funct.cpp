#include "funct.h"
#include "LaymauNgang.h"
#include "Dlg_ktrDownload.h"
#include "base64.h"

// Hàm lấy đường dẫn từ tệp tin có sẵn
CString _fGPathFolder(CString szFile)
{
	HMODULE hModule;
	hModule= GetModuleHandle(szFile);
	TCHAR szFileName[500];
	DWORD dSize = 500;
	GetModuleFileName(hModule,szFileName,dSize);

	CString szSource=szFileName; 
	for(int i=szSource.GetLength()-1;i>0;i--)
		if(szSource.GetAt(i)=='\\')
		{
			szSource=szSource.Left(
				szSource.GetLength()-(szSource.GetLength()-i)+1); 
			break;
		}

	return szSource;
}


// Hàm bổ trợ kiểm tra kết nối
bool FileExists(char* filename) 
{
	struct stat fileInfo;
	return stat(filename, &fileInfo) == 0;
}

// Hàm kiểm tra đường dẫn có tồn tại không?
int connectDsn(CString pathFile)
{
	try
	{
		char ptr[200];
		sprintf_s(ptr, "%S", pathFile);
		if(!FileExists(ptr))
		{
			CString msg = L"";
			msg.Format(L"Không load được dữ liệu xin hãy kiểm tra lại đường dẫn."
				L"\npath = %s",pathFile);
			MessageBox(NULL, msg, L"Thông báo", MB_OK | MB_ICONWARNING); 
			return -1;
		}
		else return 1;
	}
	catch(...)
	{
		MessageBox(NULL, 
			L"Không load được dữ liệu xin hãy kiểm tra lại đường dẫn", 
			L"Thông báo", MB_OK | MB_ICONWARNING);
		return -1;
	}
}

// Leo add 23.11.2018
char* _ConvertCstringToChars(CString szvalue)
{
	wchar_t* wszString = new wchar_t[szvalue.GetLength()+1];
	wcscpy(wszString,szvalue);
	int num = WideCharToMultiByte(CP_UTF8, NULL, wszString, wcslen(wszString), NULL, 0, NULL, NULL);
	char* chr = new char[num+1];
	WideCharToMultiByte(CP_UTF8, NULL,wszString,wcslen(wszString),chr,num,NULL,NULL);
	chr[num] = '\0';	
	return chr;
}

// Leo add 23.11.2018
int _FileExists(CString pathFile, int itype) 
{
	CString sThongbao = L"Đường dẫn tệp tin không hợp lệ. Vui lòng kiểm tra lại.\nPath = " + pathFile;

	try
	{
		struct stat fileInfo;
		if(!stat(_ConvertCstringToChars(pathFile), &fileInfo)) return 1;
		if(itype==1) _s(sThongbao);
		return -1;
	}
	catch(...)
	{
		if(itype==1) _s(sThongbao);
		return -1;
	}
	return 0;
}

// Hàm tách tên file
CString _TackTenFile(CString spathURL, int itype)
{
	try
	{
		CString szval=L"";
		int ilen = spathURL.GetLength();
		for (int i = ilen-1; i > 0; i--)
		{
			if(spathURL.Mid(i,1)==L"\\" || spathURL.Mid(i,1)==L"/")
			{
				if(itype==1) return spathURL.Right(ilen-i-1);
				else
				{
					spathURL = spathURL.Right(ilen-i-1);
					ilen = spathURL.GetLength();
					for (int j = ilen-1; j > 0; j--)
					{
						if(spathURL.Mid(j,1)==L".")
						{
							if(itype==0) return spathURL.Left(i);
							else spathURL.Right(ilen-i-1);
						}
					}
				}
			}
		}

		return L"";
	}
	catch(...){}
	return L"";
}

// Hàm lấy sheet name từ code name
_bstr_t name_sheet(_bstr_t code_name)
{
	try
	{
		_bstr_t cd,name_;
		_WorksheetPtr psheet_ = pWb->ActiveSheet;
	
		int p_count = xl->ActiveWorkbook->Sheets->Count;
		for (int i=1;i<=p_count;i++)
		{	
			psheet_ = pWb->Worksheets->Item[i];
			cd = psheet_->CodeName;
			if(cd==code_name)
			{
				name_=psheet_->Name;
				return name_;
				break;
			}
		}

		return name_;
	}
	catch(...)
	{
		CString msg=L"";
		msg.Format(L"Không tìm thấy sheet '%s'. Vui lòng kiểm tra lại.",(LPCTSTR)code_name);
	}

	return "";
}


// hàm tách lấy 1 ký tự
wchar_t *tachkytu_(wchar_t s[],int n)
{
	
	if (n>wcslen(s))
		return s;
	else
	{
		wchar_t *s1;
		s1=new wchar_t[n+1];
		int i;
		for (i=0;i<n;i++)
		{
			s1[i]=s[i];
		}
		s1[n]='\0';
		s1=&s1[n-1];
		return s1;
	}
	
}

// hàm tách lấy 1 chuỗi ký tự
wchar_t *tachchuoi_(wchar_t s[],int n,int vt)  // s:chuỗi - n:số ký tự lấy ra từ trái->phải - vt:số lượng chuỗi bị cắt đi từ trái->phải
{
	
	if (n>wcslen(s)||vt>=n)
		return s;
	else
	{
		wchar_t *s1;
		s1=new wchar_t[n+1];
		int i;
		for (i=0;i<n;i++)
		{
			s1[i]=s[i];
		}
		s1[n]='\0';
		s1=&s1[vt];
		return s1;
	}
	
}


// Hàm tìm kiếm comment
int FindComment(_WorksheetPtr pSheet, int col, _bstr_t comment)
{
	int row_end=0;
	CString szVal =L"";
	try
	{
		RangePtr fRange,pRng, pRange = pSheet->Cells;
		
		szVal.Format(L"%s:%s", GAR(1, col, pRange,0), GAR(65000, col, pRange, 0));
		fRange = pRange -> GetRange(COleVariant(szVal));
		if(pRng=fRange->Find(
			comment,vtMissing,XlFindLookIn::xlComments,XlLookAt::xlWhole,
			XlSearchOrder::xlByColumns,XlSearchDirection::xlNext,
			false,false,false))
		row_end = pRng->Cells->Row;

		// Reset Find (23.04.2016)
		szVal.Format(L"%s:%s", GAR(1, 1, pRange,0), GAR(1, 1, pRange, 0));
		fRange = pRange -> GetRange(COleVariant(szVal));
		fRange->Find((_bstr_t)L"",vtMissing,XlFindLookIn::xlFormulas,XlLookAt::xlPart,
			XlSearchOrder::xlByColumns,XlSearchDirection::xlNext,
			false,false,false);
		
		if(row_end == 0) 
		{
			if (comment != (_bstr_t)L"END")
			{
				szVal.Format(L"Không tìm thấy comment \"%s\" trong sheet \"%s\"."
					, (wchar_t*)comment.GetBSTR(), (wchar_t*)pSheet->Name.GetBSTR());
				MessageBox(NULL, szVal, L"Thông báo", MB_OK | MB_ICONWARNING);
				return 0;
			}			
		}
		else
		{
			return row_end;
		}
	}
	catch(...)
	{
		if (comment != (_bstr_t)L"END")
		{
			szVal.Format(L"Không tìm thấy comment \"%s\" trong sheet \"%s\"."
				, (wchar_t*)comment.GetBSTR(), (wchar_t*)pSheet->Name.GetBSTR());
			MessageBox(NULL, szVal, L"Thông báo", MB_OK | MB_ICONWARNING);
			return 0;
		}		
	}

	return (int)pSheet->Cells->SpecialCells(xlCellTypeLastCell)->GetRow() + 1;
}


// Hàm tìm kiếm comment cuối cùng
int FindCEnd(_WorksheetPtr pSheet, int col, _bstr_t comment, int iStart)
{
	int row_end=0;
	CString szVal =L"";
	try
	{
		RangePtr fRange,pRng, pRange = pSheet->Cells;
		
		szVal.Format(L"%s:%s", GAR(iStart, col, pRange,0), GAR(65000, col, pRange, 0));
		fRange = pRange -> GetRange(COleVariant(szVal));
		if(pRng=fRange->Find(
			comment,vtMissing,XlFindLookIn::xlComments,XlLookAt::xlWhole,
			XlSearchOrder::xlByColumns,XlSearchDirection::xlNext,
			false,false,false))
		row_end = pRng->Cells->Row;

		// Reset Find (27.04.2017)
		szVal.Format(L"%s:%s", GAR(1, 1, pRange,0), GAR(1, 1, pRange, 0));
		fRange = pRange -> GetRange(COleVariant(szVal));
		fRange->Find((_bstr_t)L"",vtMissing,XlFindLookIn::xlFormulas,XlLookAt::xlPart,
			XlSearchOrder::xlByColumns,XlSearchDirection::xlNext,
			false,false,false);

		return row_end;
	}
	catch(...){return 0;}
}


// Hàm tìm kiếm giá trị
// col=0 --> lấy toàn bộ bảng (cột cần tìm)
// row=0 --> row=1 (dòng bắt đầu tìm)
int FindValue(_WorksheetPtr pSheet, int col, int rowbd, int rowkt, _bstr_t bvalue)
{
	int row_end=0;
	CString szVal =L"";
	try
	{
		RangePtr PRS;
		RangePtr pRange = pSheet->Cells;
		RangePtr fRange = pRange;		
		
		if(rowbd<=0) rowbd=1;
		if(rowkt<=rowbd) rowkt=65000;

		if(col>0)
		{
			szVal.Format(L"%s:%s", GAR(rowbd, col, pRange,0), GAR(rowkt, col, pRange, 0));
			fRange = pRange->GetRange(COleVariant(szVal));
		}

		if(PRS = fRange->Find(
			bvalue,vtMissing,XlFindLookIn::xlValues,XlLookAt::xlPart,
			XlSearchOrder::xlByRows,XlSearchDirection::xlNext,
			false,false,false))
		row_end = PRS->Cells->Row;

		// Reset Find (27.04.2017)
		szVal.Format(L"%s:%s", GAR(1, 1, pRange,0), GAR(1, 1, pRange, 0));
		fRange = pRange -> GetRange(COleVariant(szVal));
		fRange->Find((_bstr_t)L"",vtMissing,XlFindLookIn::xlFormulas,XlLookAt::xlPart,
			XlSearchOrder::xlByColumns,XlSearchDirection::xlNext,
			false,false,false);

		return row_end;
	}
	catch(...){return 0;}
	return 0;
}


// Hàm lấy giá trị cột chứ name
int getColumn(_WorksheetPtr pSheet, CString name)
{
   try
	{
		RangePtr pRange = pSheet->Cells;
		RangePtr dRangeptr = pRange->GetRange((_variant_t)name, (_variant_t)name);
		int val = dRangeptr->Cells->Column;
		return val;
	}
	catch(...)
	{
		f_Khoiphuc_name(name);
		if(iKpcol>0) return iKpcol;

		CString msg=L"";
			msg.Format(L"Không tìm thấy name \"%s\" trong sheet \"%s\".",name,(wchar_t*)pSheet->Name.GetBSTR());
			MessageBox(NULL, msg , L"Thông báo",MB_OK | MB_ICONWARNING);
		return 100; 
	}
}

// Hàm lấy giá trị dòng chứa name
int getRow(_WorksheetPtr pSheet, CString name)
{
	try
	{
		RangePtr pRange = pSheet->Cells;
		RangePtr dRangeptr=pRange->GetRange((_variant_t)name, (_variant_t)name);
		int val = dRangeptr->Cells->Row;
		return val;
	}
	catch(...)
	{
		f_Khoiphuc_name(name);
		if(iKprow >0) return iKprow;

		CString msg=L"";
		msg.Format(L"Không tìm thấy name \"%s\" trong sheet \"%s\".",name,(wchar_t*)pSheet->Name.GetBSTR());
		MessageBox(NULL, msg , L"Thông báo",MB_OK | MB_ICONWARNING);
		return 100;
	}
}


void putVal(_WorksheetPtr pSheet,_bstr_t name,_variant_t sVal)
{
	try
	{
		RangePtr _pRg1 = pSheet->Cells;
		RangePtr PRS = _pRg1->GetRange(name,name);
		PRS->PutValue2(sVal);
	}
	catch(...)
	{
		f_Khoiphuc_name((LPCTSTR)name);		
		if(szKpname!=L"")
		{
			putVal(pSheet,name,sVal);
			return;
		}

		CString msg=L"";
			msg.Format(L"Không tìm thấy name \"%s\" trong sheet \"%s\".",name,(wchar_t*)pSheet->Name.GetBSTR());
			MessageBox(NULL, msg , L"Thông báo",MB_OK | MB_ICONWARNING);
	}
}


// hàm trả về giá trị hiển thị trên sheet
CString getVCell(_WorksheetPtr pSheet, CString name)
{
	CString val=L"";
	try
	{
		int _x0 = getRow(pSheet,name);
		int _y0 = getColumn(pSheet,name);
		RangePtr pRange = pSheet->Cells;
		val = pRange->GetItem(_x0,_y0);
		return val;
	}
	catch(...)
	{
		f_Khoiphuc_name(name);
		if(szKpname!=L"") return szKpname;

		CString msg=L"";
			msg.Format(L"Không tìm thấy name \"%s\" trong sheet \"%s\".",name,(wchar_t*)pSheet->Name.GetBSTR());
			MessageBox(NULL, msg , L"Thông báo",MB_OK | MB_ICONWARNING);
		return val;
	}

	return val;
}


// hàm trả về giá trị ô
CString getVRange(_WorksheetPtr pSheet, CString name)
{
	CString val=L"";
	try
	{
		RangePtr pRange = pSheet->Cells;
		RangePtr dRangeptr=pRange->GetRange((_variant_t)name, (_variant_t)name);
		val = dRangeptr->GetValue2();
	}
	catch(_com_error & error)
	{
		CString msg=L"";
		msg.Format(L"Không tìm thấy name \"%s\" trong sheet \"%s\".",name,(wchar_t*)pSheet->Name.GetBSTR());
		MessageBox(NULL, msg , L"Thông báo",MB_OK | MB_ICONWARNING);
		return val;
	}

	return val;
}
 
// Hàm lấy ngẫu nhiên màu
COLORREF fRandomRGB()
{
	try
	{
		COLORREF clrand[14]={RGB(217,217,217),RGB(191,191,191),RGB(255,153,102),RGB(106,208,40),
			RGB(204,153,255),RGB(252,213,180),RGB(255,255,0),RGB(146,205,220),RGB(255,192,0),
			RGB(150,255,153),RGB(0,204,255),RGB(255,153,153),RGB(51,153,255),RGB(244,94,94)};

		CString szval=L"";
		time_t now = time(0);
		tm *ltm = localtime(&now);
		szval.Format(L"%d%d",ltm->tm_min,ltm->tm_sec);
		int num = _wtoi(szval);
		return clrand[(num+rand())%14];
	}
	catch(...){}

	return RGB(217,217,217);
}

// hàm xác định vị trí ô cell
//st=0 kiểu A1 , st=1 kiểu $A1, st=2 kiểu A$1, st=3 kiểu $A$1
CString GAR(int nRow,int nCol,RangePtr pRng,int st)
{
	RangePtr dRangeptr = pRng->GetItem(nRow,nCol);
	_bstr_t szAdress = dRangeptr->GetAddress(false,false,XlReferenceStyle::xlA1);
	CString kq = (LPCTSTR)szAdress;

	if(st==0) return kq;

	int k,L = kq.GetLength();
	for (int i = 1; i < L; i++)
		if(_wtoi(kq.Mid(i,1))>0)
		{
			k=i;
			break;
		}

	if(st==1) kq.Format(L"$%s",kq);
	else if(st==2) kq.Format(L"%s$%s",kq.Left(k),kq.Right(L-k));
	else if(st==3) kq.Format(L"$%s$%s",kq.Left(k),kq.Right(L-k));

	return kq;
}

// Hàm try catch #REF
CString GIOR(int nRow,int nCol,RangePtr pRng,CString fkey)
{
	try 
	{
		_variant_t vItem= pRng->GetItem(nRow,nCol);
		_bstr_t bstr(vItem);
		LPCTSTR lpstr= bstr;
		return (CString)lpstr;
	}
	catch(_com_error &e) {return fkey;}
}

// hàm định nghĩa phép toán định mức
int operator >> ( const MyQSort &t, const MyQSort &u )
{
	if (t.ival < u.ival) return -1;
	if (t.ival == u.ival) return 0;
	if (t.ival > u.ival) return 1;
}


// hàm sắp xếp theo thứ tự phép toán định nghĩa
int compare_AZ (const void * a, const void * b) 
{
	return (*(MyQSort*)a >> *(MyQSort*)b);
}


// hàm so sánh ngày tháng
// numDay = số ký tự ngày (8 hoặc 10)
// return = 0 tương ứng với date1 > date2
// return = 1 tương ứng với date1 <= date2
int compare_date(int numDay,CString date1,CString date2)
{
	int iLen1 = date1.GetLength();
	int iLen2 = date1.GetLength();

	if(iLen1==numDay && iLen2==numDay)
	{
		int d1 = _wtoi(date1.Left(2));
		int d2 = _wtoi(date2.Left(2));

		int m1 = _wtoi(date1.Mid(3,2));
		int m2 = _wtoi(date2.Mid(3,2));

		int y1 = _wtoi(date1.Right(4));
		int y2 = _wtoi(date2.Right(4));

		if(y1<y2 || (y1==y2 && m1<m2) || (y1==y2 && m1==m2 && d1<=d2)) return 1;
		else return 0;
	}

	return 0;
}


void f_Setup_Print(_WorksheetPtr pSh1,int iType)
{
	try
	{
		PageSetupPtr _psetPrt = pSh1->PageSetup;

		// Thiết lập mặc định
		//_psetPrt->PaperSize = xlPaperA4;
		//_psetPrt->Orientation = xlPortrait;

		_psetPrt->PrintGridlines = 0;
		_psetPrt->PrintComments = xlPrintNoComments;
		if(_myPrint.ihtop>0) _psetPrt->HeaderMargin = xl->Application->CentimetersToPoints((double) _myPrint.ihtop/10);
		if(_myPrint.ifbot>0) _psetPrt->FooterMargin = xl->Application->CentimetersToPoints((double) _myPrint.ifbot/10);

		// Thông tin tiêu đề đầu trang
		if(_myPrint.ctxthd!=L"")
		{
			CString fHeader = _myPrint.ctxthd;
			if(_myPrint.istylehd==0) fHeader.Format(L"&%d%s",_myPrint.isizehd,_myPrint.ctxthd);
			else if(_myPrint.istylehd==4) fHeader.Format(L"&U%s",_myPrint.ctxthd);
			else fHeader.Format(L"&\"Times New Roman,%s\"&%d%s", _myPrint.cstylehd, _myPrint.isizehd, _myPrint.ctxthd);

			if(_myPrint.isechd==1)
			{
				// Căn giữa
				_psetPrt->LeftHeader = (_bstr_t)L"";
				_psetPrt->CenterHeader = (_bstr_t)fHeader;
				_psetPrt->RightHeader = (_bstr_t)L"";
			}
			else if(_myPrint.isechd==2)
			{
				// Căn phải
				_psetPrt->LeftHeader = (_bstr_t)L"";
				_psetPrt->CenterHeader = (_bstr_t)L"";
				_psetPrt->RightHeader = (_bstr_t)fHeader;
			}
			else
			{
				// Căn trái
				_psetPrt->LeftHeader = (_bstr_t)fHeader;
				_psetPrt->CenterHeader = (_bstr_t)L"";
				_psetPrt->RightHeader = (_bstr_t)L"";
			}
		}

		// Thông tin tiêu đề chân trang
		if(_myPrint.ctxtft!=L"")
		{
			CString fFooter = _myPrint.ctxtft;
			if(_myPrint.istyleft==0) fFooter.Format(L"&%d%s",_myPrint.isizeft,_myPrint.ctxtft);
			else if(_myPrint.istyleft==4) fFooter.Format(L"&U%s",_myPrint.ctxtft);
			else fFooter.Format(L"&\"Times New Roman,%s\"&%d%s", _myPrint.cstyleft, _myPrint.isizeft, _myPrint.ctxtft);

			if(_myPrint.isecft==1)
			{
				_psetPrt->LeftFooter = (_bstr_t)L"";
				_psetPrt->CenterFooter = (_bstr_t)fFooter;
				_psetPrt->RightFooter = (_bstr_t)L"";
			}
			else if(_myPrint.isecft==2)
			{
				_psetPrt->LeftFooter = (_bstr_t)L"";
				_psetPrt->CenterFooter = (_bstr_t)L"";
				_psetPrt->RightFooter = (_bstr_t)fFooter;
			}
			else
			{
				_psetPrt->LeftFooter = (_bstr_t)fFooter;
				_psetPrt->CenterFooter = (_bstr_t)L"";
				_psetPrt->RightFooter = (_bstr_t)L"";
			}
		}

		// Căn chỉnh lề trang in
		if(_myPrint.ileft>0) _psetPrt->LeftMargin = xl->Application->CentimetersToPoints((double) _myPrint.ileft/10);
		if(_myPrint.itop>0) _psetPrt->TopMargin = xl->Application->CentimetersToPoints((double) _myPrint.itop/10);
		if(_myPrint.iright>0) _psetPrt->RightMargin = xl->Application->CentimetersToPoints((double) _myPrint.iright/10);
		if(_myPrint.ibottom>0) _psetPrt->BottomMargin = xl->Application->CentimetersToPoints((double) _myPrint.ibottom/10);
	
		// Căn chỉnh giữa
		_psetPrt->CenterHorizontally = _myPrint.ihor;
		_psetPrt->CenterVertically = _myPrint.iver;

		// Tùy chọn khác (đánh số trang in thiết lập ngoài)
		_psetPrt->BlackAndWhite = _myPrint.ibwhite;
		if(_myPrint.ierror==1) _psetPrt->PrintErrors = xlPrintErrorsBlank;

		// Thiết lập đầy trang
		if(_myPrint.izoom>0) _psetPrt->PutZoom(_myPrint.izoom);
		else
		{
			if(iType==1)
			{
				xl->ActiveWindow->PutView(xlPageBreakPreview);
				int _idoc = pSh1->VPageBreaks->GetCount()+1;

				// Kéo trục dọc thành 1 vùng in duy nhất
				if(_idoc>=2) pSh1->VPageBreaks->GetItem(1)->DragOff(xlToRight,1);
			}			
		}
	}
	catch(...){_s(L"#QL06.04: Lỗi thiết lập bảng tính.");}
}


CString fGetRowsSLS(CString _cRows)
{
	int ncount = _cRows.GetLength();
	int _pos = _cRows.Find(L":");
	CString g0=L"",g1,g2;
	if(_pos==-1)
	{
		// Tồn tại duy nhất 1 ô
		_pos = _cRows.Find(L"C");
		g0 = _cRows.Mid(1,_pos-1);
	}
	else
	{
		// Tồn tại nhiều hơn 1 ô
		g1 = _cRows.Left(_pos).Mid(1,_cRows.Left(_pos).Find(L"C")-1);
		g2 = _cRows.Right(ncount-_pos-1).Mid(1,_cRows.Right(ncount-_pos-1).Find(L"C")-1);
		
		// Kiểm tra vị trí dòng có trùng nhau không?
		if(g1==g2) g0=g1;
		else g0.Format(L"%s-%s",g1,g2);
	}

	return g0;
}

int _fGetSelection(CString _cVal)
{
	// Khai báo mảng struct MySelection để gán địa chỉ cột, dòng
	if(_cVal==L"") return 0;
	if(_cVal.Right(1)!=L"," && _cVal.Right(1)!=L";") _cVal+=L",";	

	int nLen = _cVal.GetLength();
	int i=0;
	int j=0;
	int dem=0;
	CString val=L"";
	while (j<nLen)
	{
		val=_cVal.Mid(j,1);
		if(val==L"," || val==L";")
		{
			dem++;
			_arrSLS[dem] = fGetRowsSLS(_cVal.Mid(i,j-i));
			i=j+1;
		}
		j++;
	}

	return dem;
}

// Tách từ khóa
void _TackTukhoa(CString sKey, CString sChar1, CString sChar2)
{
	try
	{
		iKey=0;
		sKey.Replace(L"\n",L"");
		if(sKey==L"") return;
		if(sKey.Right(1)!=sChar1) sKey+=sChar1;		
		
		int vtri=0,iktr=0;
		CString szval=L"",sKtr=L"";
		for (int i = 0; i < sKey.GetLength(); i++)
		{
			szval = sKey.Mid(i,1);
			if((szval==sChar1 || szval==sChar2) && i>=vtri)
			{
				iktr=0;
				if(i>vtri) sKtr = sKey.Mid(vtri,i-vtri);
				else sKtr=L"";
				
				if(iKey>=1)
				for (int j = 1; j <= iKey; j++)
				{
					if(sKtr==pKey[j])
					{
						iktr=1;
						break;
					}
				}			

				if(iktr==0)
				{
					iKey++;
					pKey[iKey] = sKtr;
					pKey[iKey].Trim();
				}
				
				vtri=i+1;
			}
		}
	}
	catch(...){}
}


// Hàm chuyển chữ thường thành chữ hoa
void f_proper_string(Excel::RangePtr pRg1,int nR,int nC)
{
	try
	{
		// Chỉnh lại tên viết hoa
		CString val = GIOR(nR,nC,pRg1,L"");
		if(val==L"") return;

		val.Format(L"=TRIM(PROPER(%s))",GAR(nR,nC,pRg1,0));
		pRg1->PutItem(nR,89,(_bstr_t)val);
		val = pRg1->GetItem(nR,89);

		RangePtr PRS = pRg1->GetItem(nR,89);
		PRS->ClearContents();
		pRg1->PutItem(nR,nC,(_bstr_t)val);
	}
	catch(...){}
}


// Hàm làm tròn số
double fRoundDouble(double doValue, int nPrecision)
{
	static const double doBase = 10.0;
	double doComplete5, doComplete5i;
	
	doComplete5 = doValue * pow(doBase, (double) (nPrecision + 1));
	
	if(doValue < 0.0)
		doComplete5 -= 5.0;
	else
		doComplete5 += 5.0;
	
	doComplete5 /= doBase;
	modf(doComplete5, &doComplete5i);
	
	return doComplete5i / pow(doBase, (double) nPrecision);
}


// Hàm thêm phẩy (,) vào giá tiền
CString FormatTien(CString ftien,CString fCach,CString fSTP)
{
	// fCach = dấu ngăn cách (dấu ,) hoặc ngược lại
	// fSTP = dấu thập phân (dấu .)

	ftien.Replace(fCach,fSTP);
	int adg = 1;
	if(ftien.Left(1)==L"-")
	{
		adg=0;
		ftien.Replace(L"-",L"");
	}

	CString _ftien = ftien;

	// Kiểm tra xem có phần lẻ không
	int iLen = ftien.GetLength();
	CString ktra = L"";
	int tang=0;
	ktra = ftien.Mid(tang,1);
	while (ktra!=fSTP && tang<iLen)
	{
		tang++;
		ktra = ftien.Mid(tang,1);
	}

	CString m1=ftien,m2=L"";
	if(tang<iLen)
	{
		// Tồn tại phần thập phân
		m1 = ftien.Left(tang);
		m2 = ftien.Right(iLen-tang);

		// Kiểm tra xem m2 có chẵn > 0 không?
		if(_wtof(m2.Right(m2.GetLength()-1))==0) m2=L"";
	}

	CString tak[10];
	// Đếm số ký tự
	iLen = m1.GetLength();
	int du = iLen%3;
	int tg = (iLen-du)/3;

	if(iLen>3)
	{
		// Tách 3 ký tự
		if(du==1)
		{
			tak[0] = m1.Left(1);
			for(int i=1;i<=tg;i++) tak[i] = m1.Mid(3*i-2,3);
		}
		else if(du==2)
		{
			tak[0] = m1.Left(2);
			for(int i=1;i<=tg;i++) tak[i] = m1.Mid(3*i-1,3);
		}
		else for(int i=0;i<tg;i++) tak[i] = m1.Mid(3*i,3);

	// Ghép ký tự
	m1 = tak[0];
	for (int i=1;i<tg;i++) m1.Format(L"%s%s%s",m1,fCach,tak[i]);
	if(du>0) m1.Format(L"%s%s%s",m1,fCach,tak[tg]);
	
	}

	_ftien = m1+m2;
	if(adg==0) _ftien = L"-" + _ftien;

	return _ftien;
}


// Hàm kiểm tra ngày có bị đổ ngược không!
void f_ktra_date(CString sDate, RangePtr _pRg, int nRow, int nCol)
{
	CString szval = L"";
	RangePtr PRS = _pRg->GetItem(nRow,nCol);
	PRS->PutNumberFormat(L"dd/mm/yyyy");

	if(_wtoi(sDate.Left(2))>12)
	{
		szval.Format(L"%s/%s/%s",sDate.Mid(3,2),sDate.Left(2),sDate.Right(sDate.GetLength()-6));
		_pRg->PutItem(nRow,nCol,(_bstr_t)szval);
		return;
	}

	_pRg->PutItem(nRow,nCol,(_bstr_t)sDate);
	szval = _pRg->GetItem(nRow,nCol);
	if(szval!=sDate) _pRg->PutItem(nRow,nCol,(_bstr_t)szval);
}


int f_Check_printer(_bstr_t _prt)
{
	CString msg = (LPCTSTR)_prt;
	int _pos1 = msg.Find(L"XPS");
	int _pos2 = msg.Find(L"PDF");
	if(_pos1>0 || _pos2>0)
	{
		_s(L"Lựa chọn mặc định đang là máy in ảo. "
			L"\nVui lòng nhấn chọn lại máy in hoặc sử dụng tính năng lưu hồ sơ.");
		return 1;
	}

	return 0;
}


void f_call_moduleXLA(CString sfModule, CString sfSub)
{
	try
	{
		_variant_t  varUnused( (long) DISP_E_PARAMNOTFOUND, VT_ERROR );
		CString PATH_MODULEXLA=L"";
		PATH_MODULEXLA.Format(L"QLCL.xla!%s.%s",sfModule,sfSub);
		_variant_t szMacro(PATH_MODULEXLA);

		xl->Run(szMacro,
			varUnused, varUnused, varUnused, varUnused,	varUnused, varUnused, varUnused, varUnused, 
			varUnused, varUnused, varUnused, varUnused,	varUnused, varUnused, varUnused, varUnused, 
			varUnused, varUnused, varUnused, varUnused, varUnused, varUnused, varUnused, varUnused,
			varUnused, varUnused, varUnused, varUnused,	varUnused, varUnused);
	}
	catch(...){}
}


void f_SavePDF(CString sPth,CString sPRg)
{
	prTS->PutItem(1,100,(_bstr_t)sPth);
	prTS->PutItem(2,100,(_bstr_t)sPRg);

	// Call sub xla
	f_call_moduleXLA(L"Module1",L"Call_EAFF");

	RangePtr PRS = psTS->GetRange(prTS->GetItem(1,100),prTS->GetItem(2,100));
	PRS->ClearContents();
}


void f_Auto_stt_dmuc(RangePtr pRg1,int iDanhstt,int iType,int numbt,int numkt,int icot1,int icot2,int icot4)
{
	try
	{
		int stt=1,ttHMGD=1;
		_bstr_t bscmt;
		CString sadd=L"",strtp=L"",flst=L"";
		RangePtr PRS;

		if(iType==1 && iDanhstt>=0)
		{
			for(int j=numbt;j<numkt;j++)
			{
				strtp= pRg1->GetItem(j,icot2);
				strtp.Trim();
				if(strtp == L"HM" || strtp == L"GD")
				{
					// Kiểm tra có comment không?
					PRS = pRg1->GetItem(j,icot4);
					if(PRS->GetComment()!=NULL)
					{
						bscmt = PRS->GetComment()->Text();
						sadd = (LPCTSTR)(bscmt);
						sadd.Trim();
					}
					else
					{
						if(strtp == L"HM") sadd=L"";
					}

					// Ngắt đánh stt khi chuyển sang HM hoặc GD mới
					if((strtp == L"HM" && iDanhstt==1) || (strtp == L"GD" && iDanhstt==2)) ttHMGD=1;
				}
				else
				{
					if(strtp!=L"" && strtp.Find(L"*")==-1)
					{
						if(ttHMGD<10) flst.Format(L"0%d",ttHMGD);				
						else flst.Format(L"%d",ttHMGD);
						if(sadd!=L"") flst+=sadd;

						pRg1->PutItem(j,icot1,stt);
						pRg1->PutItem(j,icot4,(_bstr_t)flst);
						stt++;
						ttHMGD++;
					}
				}
			}
		}
		else
		{
			for(int j=numbt;j<numkt;j++)
			{
				strtp= pRg1->GetItem(j,icot2);
				strtp.Trim();
				if(strtp != L"HM" && strtp != L"GD" && strtp != L"")
				{
					pRg1->PutItem(j,icot1,stt);
					stt++;
				}
			}
		}		
	}
	catch(...){_s(L"Lỗi đánh lại số thứ tự");}
}


void f_Sort_danhmuc(CString sNSheet,CString sRsort,CString sPRg)
{
	prTS->PutItem(1,100,(_bstr_t)sNSheet);
	prTS->PutItem(2,100,(_bstr_t)sRsort);
	prTS->PutItem(3,100,(_bstr_t)sPRg);

	// Call sub xla
	f_call_moduleXLA(L"Module1",L"Call_sortDM");

	RangePtr PRS = psTS->GetRange(prTS->GetItem(1,100),prTS->GetItem(3,100));
	PRS->ClearContents();
}


int _formatPixcel(RangePtr pRg1, int cot)
{
	RangePtr PRS = pRg1->GetItem(1,cot);
	int _iPoints  = PRS->GetWidth();
	int _iChars  = PRS->GetColumnWidth();
	int _iPixcels = (int)(4*_iPoints/3);

	return _iPixcels;

	// COLUMNWIDTH
	// screenres = 96;
	//_iPixcels = (_iPoints/72)*screenres = (4/3)*_iPoints;

	// ROWHIGHT
	// 1 pixel = 72 / 96 points= 0.75 points.
}


bool GetImageSize(const char *fn, int *x,int *y)
{
	FILE *f=fopen(fn,"rb");
	if (f==0) return false;
	fseek(f,0,SEEK_END);
	long len=ftell(f);
	fseek(f,0,SEEK_SET);
	if (len<24) 
	{
		fclose(f);
		return false;
	}

  // Strategy:
  // reading GIF dimensions requires the first 10 bytes of the file
  // reading PNG dimensions requires the first 24 bytes of the file
  // reading JPEG dimensions requires scanning through jpeg chunks
  // In all formats, the file is at least 24 bytes big, so we'll read that always
  unsigned char buf[24]; fread(buf,1,24,f);

  // For JPEGs, we need to read the first 12 bytes of each chunk.
  // We'll read those 12 bytes at buf+2...buf+14, i.e. overwriting the existing buf.
  
  // Đoạn code cũ
  /*if (buf[0]==0xFF && buf[1]==0xD8 && buf[2]==0xFF && buf[3]==0xE0 && buf[6]=='J' && buf[7]=='F' && buf[8]=='I' && buf[9]=='F')
  { 
	long pos=2;
	while (buf[2]==0xFF)
	{ if (buf[3]==0xC0 || buf[3]==0xC1 || buf[3]==0xC2 
			|| buf[3]==0xC3 || buf[3]==0xC9 || buf[3]==0xCA || buf[3]==0xCB) break;
	  
	  pos += 2+(buf[4]<<8)+buf[5];
	  if (pos+12>len) break;
	  fseek(f,pos,SEEK_SET);
	  fread(buf+2,1,12,f);
	}
  }*/

  // 24.12.2016: Code mới (Đã fix lỗi xác định hình ảnh đã qua chỉnh sửa photoshop: x= 26112 & y=30825)
  // 05/01/2017: Thêm 1 trường hợp ảnh down từ facebook: x=17219 & y=24400
  /*if ((buf[0]==0xFF && buf[1]==0xD8 && buf[2]==0xFF && buf[3]==0xE0 && buf[6]=='J' && buf[7]=='F' && buf[8]=='I' && buf[9]=='F') ||
		(buf[0]==0xFF && buf[1]==0xD8 && buf[2]==0xFF && buf[3]==0xE1 && buf[6]=='E' && buf[7]=='x' && buf[8]=='i' && buf[9]=='f'))*/
  if (buf[0]==0xFF && buf[1]==0xD8 && buf[2]==0xFF)
	{
		long pos=2;
		while (buf[2]==0xFF)
		{
			if (buf[3]==0xC0 || buf[3]==0xC1 || buf[3]==0xC2 
				|| buf[3]==0xC3 || buf[3]==0xC9 || buf[3]==0xCA || buf[3]==0xCB) break;

			pos += 2+(buf[4]<<8)+buf[5];
			if (pos+12>len) break;
			fseek(f,pos,SEEK_SET);
			fread(buf+2,1,12,f);
		}
	}

  fclose(f);

  // JPEG: (first two bytes of buf are first two bytes of the jpeg file; rest of buf is the DCT frame
  if (buf[0]==0xFF && buf[1]==0xD8 && buf[2]==0xFF)
  {

	*y = (buf[7]<<8) + buf[8];
	*x = (buf[9]<<8) + buf[10];

	return true;
  }

 // // GIF: first three bytes say "GIF", next three give version number. Then dimensions
 // if (buf[0]=='G' && buf[1]=='I' && buf[2]=='F')
 // {
	//*x = buf[6] + (buf[7]<<8);
 //   *y = buf[8] + (buf[9]<<8);
 //   return true;
 // }

 // // PNG: the first frame is by definition an IHDR frame, which gives dimensions
 // if ( buf[0]==0x89 && buf[1]=='P' && buf[2]=='N' && buf[3]=='G' && buf[4]==0x0D && buf[5]==0x0A && buf[6]==0x1A && buf[7]==0x0A
 //   && buf[12]=='I' && buf[13]=='H' && buf[14]=='D' && buf[15]=='R')
 // {
	//*x = (buf[16]<<24) + (buf[17]<<16) + (buf[18]<<8) + (buf[19]<<0);
 //   *y = (buf[20]<<24) + (buf[21]<<16) + (buf[22]<<8) + (buf[23]<<0);
 //   return true;
 // }

  return false;
}


void _fGetImage(CString sPath)
{
	try
	{
		// Convert Cstring to Char
		const char *theFile = ConvertToChar(sPath);
		int the_x = 0;
		int the_y = 0;

		// Xác định kích thước ảnh
		bool didRun = false;
		didRun = GetImageSize(theFile, &the_x, &the_y);

		// Gán giá trị kích thước thực của ảnh tìm được
		iWThuc = the_x;
		iHThuc = the_y;
	}
	catch(...){}
}


int f_Check_ban_quyen()
{
	try
	{
		// LEO 18.04.2017
		bool kt = false;
		typedef bool (__stdcall *p)();
		HINSTANCE loadDL = LoadLibrary(L"DllQLCL.dll");
		p _pCall ;

		if(loadDL != NULL)
		{
			_pCall  = (p)GetProcAddress(loadDL,"CheckLicense");
			if(_pCall != NULL) kt = _pCall ();
		}

		FreeLibrary(loadDL);
		return kt;
	}
	catch(...){}
}


int f_Mod_check()
{
	int num=10;
	count_key++;
	if (count_key % num == (num-1))
	{
		if(f_Check_ban_quyen()!=1)
		{
			count_key=num-2;
			return 0;
		}

		count_key=0;
	}
	
	return 1;
}


void f_Sort_stt_qc()
{
	int dem = listQclm.GetItemCount();
	if(dem==0) return;

	CString szval=L"";
	for (int i = 0; i < dem; i++)
	{
		if(i<9) szval.Format(L"0%d",i+1);
		else szval.Format(L"%d",i+1);
		listQclm.SetItemText(i,0,szval);	
	}
}


// Hàm tìm kiếm nhiều ký tự
int FindMulti(CString sChuoi, CString sFind)
{
	try
	{
		// Tách chuỗi
		int dem=0, vtri=0;
		sFind.Trim();
		if(sFind==L"") return 0;

		sFind += L"|";
		CString szval=L"", sTak[100];
		for (int i = 0; i < sFind.GetLength(); i++)
		{
			szval = sFind.Mid(i,1);
			if(szval==L"|" && i>vtri)
			{
				dem++;
				sTak[dem] = sFind.Mid(vtri,i-vtri);
				vtri=i+1;
			}
		}

		vtri = -1;
		for (int i = 1; i <= dem; i++)
		{
			vtri = sChuoi.Find(sTak[i]);
			if(vtri>=0) break;
		}

		return vtri;
	}
	catch(...){ return 0;}
}


// Hàm diễn giải khổi lượng
void xl_calculator_dgkl(RangePtr pRg1,int dongAC,int cotAC,int cotLG)
{
	try
	{
		CString szval=L"";
		szval.Format(L"=kl(%s)",GAR(dongAC,cotAC,pRg1,0));
		pRg1->PutItem(dongAC,cotLG,(_bstr_t)szval);
		Excel::RangePtr PRS = pRg1->GetItem(dongAC,cotLG);
		
		szval = GIOR(dongAC,cotLG,pRg1,L"");
		if(szval!=L"")
		{
			PRS->Font->PutColor(BLUE);
			PRS->Font->PutItalic(1);
		}
		else
		{
			PRS->Font->PutColor(BLACK);
			PRS->Font->PutItalic(0);
		}

		//RangePtr PRS = pRg1->GetItem(dongAC,cotAC);
		//CString szval = PRS->GetFormula();
		//szval.Trim();
		//int pos = szval.Find(L":");
		//if(pos>0)
		//{
		//	int iktr=-1;
		//	CString s0=L"";
		//	CString sten = szval.Left(pos);
		//	for (int i = 0; i < sten.GetLength(); i++)
		//	{
		//		s0 = sten.Mid(i,1);
		//		if(_wtoi(s0)==0 && s0!=L"0")
		//		{
		//			iktr=i;
		//			break;
		//		}
		//	}

		//	if(iktr>=0) szval = szval.Right(szval.GetLength()-pos-1);
		//}		

		//szval.Replace(L",",L".");
		//szval.Replace(L"x",L"*");
		//szval.Replace(L"X",L"*");
		//szval.Replace(L":",L"/");
		//szval.Replace(L"[",L"(");
		//szval.Replace(L"]",L")");

		//if(szval==L"") return;
		//if(szval.Left(1)!=L"=") szval = L"="+szval;
		//pRg1->PutItem(dongAC,cotLG,(_bstr_t)szval);	// đổ công thức
		//szval = GIOR(dongAC,cotLG,pRg1,L"");
		//szval.Replace(L",",L".");
		//pRg1->PutItem(dongAC,cotLG,(_bstr_t)szval);	// copy và đổ số chết
		//
		//if(szval!=L"")
		//{
		//	PRS = pRg1->GetItem(dongAC,cotLG);
		//	PRS->Font->PutColor(BLUE);
		//	PRS->Font->PutItalic(1);
		//}
		//else
		//{
		//	PRS = pRg1->GetItem(dongAC,cotLG);
		//	PRS->Font->PutColor(BLACK);
		//	PRS->Font->PutItalic(0);
		//}
	}
	catch(...)
	{
		pRg1->PutItem(dongAC,cotLG,(_bstr_t)L"");
		RangePtr PRS = pRg1->GetItem(dongAC,cotLG);
		PRS->Font->PutColor(BLACK);
		PRS->Font->PutItalic(0);		
	}
}


// Hàm khôi phục các name không tồn tại
void f_Khoiphuc_name(CString sname)
{
	try
	{
		szKpname = L"";
		iKprow = 0, iKpcol = 0;

		if (sname == "DMBB_HK_TONGNGHI")
		{
			_bstr_t shDMCV = name_sheet("shHSNTCV");
			_WorksheetPtr psDMCV = xl->Sheets->GetItem(shDMCV);
			RangePtr prDMCV = psDMCV->GetCells();

			int idong = getRow(psDMCV, L"DMBB_HK_TONGNGAY");
			int icot = getColumn(psDMCV, L"DMBB_HK_TONGNGAY");
			
			RangePtr PRS = prDMCV->GetItem(1, icot + 1);
			PRS->EntireColumn->Insert(xlToRight);

			iKpcol = icot + 1;
			PRS = prDMCV->GetItem(idong, iKpcol);
			PRS->PutName((_bstr_t)sname);
			putVal(psDMCV, (_bstr_t)sname, (_bstr_t)L"TỔNG SỐ NGÀY NGHỈ");
		}		
		else if (sname == "DMBB_HK_CTNGHI")
		{
			_bstr_t shDMCV = name_sheet("shHSNTCV");
			_WorksheetPtr psDMCV = xl->Sheets->GetItem(shDMCV);
			RangePtr prDMCV = psDMCV->GetCells();

			int idong = getRow(psDMCV, L"DMBB_HK_TONGNGHI");
			int icot = getColumn(psDMCV, L"DMBB_HK_TONGNGHI");

			RangePtr PRS = prDMCV->GetItem(1, icot + 1);
			PRS->EntireColumn->Insert(xlToRight);

			iKpcol = icot + 1;
			PRS = prDMCV->GetItem(idong, iKpcol);
			PRS->PutName((_bstr_t)sname);
			putVal(psDMCV, (_bstr_t)sname, (_bstr_t)L"CHI TIẾT NGÀY NGHỈ");
		}
		else if (sname == "NHTC_OPEN")
		{
			_bstr_t shTChuan = name_sheet("shNhomTC");
			_WorksheetPtr psTChuan = xl->Sheets->GetItem(shTChuan);
			RangePtr prTChuan = psTChuan->GetCells();

			int idong = getRow(psTChuan, L"NHTC_TC");
			int icot = getColumn(psTChuan, L"NHTC_TC");
			int icot2 = getColumn(psTChuan, L"NHTC_GC");
			
			RangePtr PRS;
			if (icot + 1 == icot2)
			{
				PRS = prTChuan->GetItem(1, icot + 1);
				PRS->EntireColumn->Insert(xlToRight);
			}

			iKpcol = icot + 1;
			PRS = prTChuan->GetItem(idong, iKpcol);
			PRS->PutName((_bstr_t)sname);
			putVal(psTChuan, (_bstr_t)sname, (_bstr_t)L"MỞ FILE");
		}
	}
	catch(...)
	{
		_s(L"Không thể khôi phục được name: " + sname);
	}	
}

// Hàm trả về thứ mấy
CString fThuMay(CString sNgay)
{
	if (sNgay == L"") return L"";

	int dd = _wtoi(sNgay.Left(2));
	int mm = _wtoi(sNgay.Mid(3, 2));
	int yy = _wtoi(sNgay.Right(2)) + 2000;

	int nIndex = (dd + ((153 * (mm + 12 * ((14 - mm) / 12) - 3) + 2) / 5) +
		(365 * (yy + 4800 - ((14 - mm) / 12))) +
		((yy + 4800 - ((14 - mm) / 12)) / 4) -
		((yy + 4800 - ((14 - mm) / 12)) / 100) +
		((yy + 4800 - ((14 - mm) / 12)) / 400) - 32045) % 7;

	CString cs[7] = { L"T2",L"T3", L"T4", L"T5", L"T6", L"T7", L"CN" };
	if (nIndex >= 0 && nIndex < 7) return cs[nIndex];

	return L"";
}


CString _NameColumnHeading(int itype, int col, CString sname)
{
	try
	{
		CString strNome=L"";
		if(itype==1) strNome = sname;
		
		CHeaderCtrl *pHeader = litLMN3.GetHeaderCtrl();
		HDITEM hdi;
		hdi.mask = HDI_TEXT;
		hdi.pszText = strNome.GetBuffer(256);
		hdi.cchTextMax = 256;
		
		if(itype==0) pHeader->GetItem(col, &hdi );
		else pHeader->SetItem(col, &hdi );
		strNome.ReleaseBuffer();
		
		return strNome;
	}
	catch(...){}

	return L"";
}


void _TackTukhoaChuan(CString sKeyValue, CString symbol1, CString symbol2, int itype1, int itype2)
{
	try
	{
		iKey=0;
		sKeyValue.Replace(L"\n",L"");
		if(sKeyValue==L"") return;		
		if (sKeyValue.Right(1) != symbol1 && sKeyValue.Right(1) != symbol2)
		{
			sKeyValue += symbol1;
		}
		
		int vtri=0,iktr=0;
		CString szval=L"",sKtr=L"";
		for (int i = 0; i < sKeyValue.GetLength(); i++)
		{
			szval = sKeyValue.Mid(i,1);
			if((szval==symbol1 || szval==symbol2) && i>=vtri)
			{
				iktr=0;
				if(i>vtri) sKtr = sKeyValue.Mid(vtri,i-vtri);
				else sKtr=L"";

				if(itype1==1)
				{
					if(iKey>=1)
					for (int j = 1; j <= iKey; j++)
					{
						if(sKtr==pKey[j])
						{
							iktr=1;
							break;
						}
					}
				}

				if(iktr==0)
				{
					iKey++;
					pKey[iKey] = sKtr;
					if(itype2==0) pKey[iKey].Trim();
				}
				
				vtri=i+1;
			}
		}
	}
	catch(...){}
}


void OQCLMAU()
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
	LaymauNgang open_dlg;
	open_dlg.DoModal();
}

void CopyDulieu(CString sNoidung)
{
	try
	{
		if ( !OpenClipboard(NULL) )
		{
			_s(L"Không thể mở Clipboard");
			return;
		}
		// Remove the current Clipboard contents
		if( !EmptyClipboard() )
		{
			_s(L"Không thể để trống Clipboard");
			return;
		}
		size_t size = sizeof(TCHAR)*(1+sNoidung.GetLength());
		HGLOBAL hResult = GlobalAlloc(GMEM_MOVEABLE, size); 
		LPTSTR lptstrCopy = (LPTSTR)GlobalLock(hResult); 
		memcpy(lptstrCopy, sNoidung.GetBuffer(), size); 
		GlobalUnlock(hResult); 
		#ifndef _UNICODE
		if ( ::SetClipboardData( CF_TEXT, hResult ) == NULL )
		#else
		if ( ::SetClipboardData( CF_UNICODETEXT, hResult ) == NULL )
		#endif
		{
			TRACE(L"Không thể đặt dữ liệu Clipboard");
			GlobalFree(hResult); 
			CloseClipboard();
			return;
		}
		CloseClipboard();
	}
	catch(...){}
}


CString f_read_txt(CString sFiledl)
{
	CString strTemp=L"";
	wchar_t txtopen[255];

	FILE * fp = fopen(ConvertToChar(sFiledl), "r");
	fgetws(txtopen, 255, fp);
	strTemp.Format(L"%s", txtopen);
	fclose(fp);

	return strTemp;
}


void f_write_txt(CString sChuoi,CString sFiledl)
{
	std::ofstream myfile;
	myfile.open (ConvertToChar(sFiledl));
	myfile << ConvertToChar(sChuoi);
	myfile.close();
}


// Hàm tìm kiếm tệp tin trong thư mục
void f_Get_All_File(CString m_sBaseFolder, CString m_sFileMask)
{
	try
	{
		CFileFinder::CFindOpts	opts;
		opts.sBaseFolder = m_sBaseFolder;
		opts.sFileMask.Format(L"*%s*", m_sFileMask);
		opts.FindNormalFiles();

		_dsfile.RemoveAll();
		_dsfile.Find(opts);
	}
	catch(...){_s(L"Lỗi xác định số lượng tệp tin");}
}


// Hàm kiểm tra sự tồn tại của file (bao gồm cả file có dấu tiếng việt)
int f_CheckFile(CString pFile)
{
	try
	{
		// Tách đường dẫn & tên tệp tin
		CString sPth=L"",sFl=L"",sDuoi=L"",sval=L"";
		int iktr=0;
		int N = pFile.GetLength();
		for (int i = N-1; i > 0; i--)
		{
			sval = pFile.Mid(i,1);
			if(sval==L"\\")
			{
				sPth = pFile.Left(i+1);
				sFl = pFile.Right(N-i-1);
				break;
			}
			else if(sval==L"." && iktr==0)
			{
				iktr=1;
				sDuoi = pFile.Right(N-i-1);
			}
		}

		// Kiểm tra tệp tin có tồn tại không?
		if(sPth!=L"")
		{
			if(sDuoi!=L"") f_Get_All_File(sPth,sDuoi);
			else f_Get_All_File(sPth,L"*");

			int nCount = _dsfile.GetFileCount();
			if(nCount>0)
				for(int i = 0; i<nCount; i++)
					if(sFl==_dsfile.GetFilePath(i).GetFileName()) return 1;
		}

		return 0;
	}
	catch(...){}
}


CString _ReplateUTF8toWeb(CString szLink)
{
	for (int i = 0; i < 152; i++) szLink.Replace(scharutf8_2[i],scharweb_2[i]);
	return szLink;
}


// Chú ý khi decode web thì duyệt for ngược từ dưới lên %25 --> %
CString _ReplateWebtoUTF8(CString szLink)
{
	for (int i = 17; i >= 0; i--) szLink.Replace(scharweb_2[i],scharutf8_2[i]);
	return szLink;
}


CString _ConvertKhongDau(CString sKeyValue)
{
	for (int i = 0; i < 134; i++) sKeyValue.Replace(scvtUTF[i],scvtKOD[i]);
	return sKeyValue;
}

// sTenfile = mã tiêu chuẩn
void f_SearchGoogle(CString sWebSearch, CString sTenfile, int itype)
{
	//if(itype==1)
	//	if(_yn(L"Không tìm thấy dữ liệu. Bạn có muốn tìm kiếm trên Internet không?",0)!=6) return;

	CString sLink=L"";
	if(sWebSearch==L"") sWebSearch = f_B64DecodeEX(_Wsearch);
	sLink.Format(L"%s%s",sWebSearch,sTenfile);
	ShellExecute(NULL, L"open",sLink, NULL, NULL, SW_SHOWMAXIMIZED);
}

void _xlSetHyperlink(_WorksheetPtr pSheet, RangePtr PRS, CString pathFile, CString szName,
	COLORREF clrBkg, COLORREF clrTxt, CString szFontName, int iSize, bool bCenter, bool bFormula)
{
	try
	{
		pSheet->Hyperlinks->Add(PRS, (_bstr_t)pathFile);
		PRS->PutValue2((_bstr_t)szName);
		PRS->Interior->PutColor(clrBkg);
		PRS->Font->PutColor(clrTxt);
		PRS->Font->PutName((_bstr_t)szFontName);
		PRS->Font->PutSize(iSize);
		if (bCenter) PRS->PutHorizontalAlignment(xlCenter);
	}
	catch (...) {}
}

CString f_Get_DownloadTieuchuan(CString szTenTC)
{
	try
	{
		if (pathMDB == L"")
		{
			pathMDB = getPathQLCL();
		}

		CString sFullPath = pathMDB;
		if (sFullPath == L"") return L"";
		if (connectDsn(sFullPath) == -1) return L"";

		CConnection ObjConn;
		ObjConn.OpenAccessDB(sFullPath, L"", L"");

		CString szNameLink = L"";
		SqlString.Format(L"SELECT * FROM tbTentieuchuan WHERE MaTC LIKE '%s';", szTenTC);
		CRecordset* Recset = ObjConn.Execute(SqlString);
		while (!Recset->IsEOF())
		{
			Recset->GetFieldValue(L"Download", szNameLink);
			break;
		}
		Recset->Close();
		ObjConn.CloseDatabase();

		return szNameLink;
	}
	catch (...) {}
	return L"";
}

// sTenfile = mã tiêu chuẩn
CString f_DocFileTC(CString sTenfile, int iOpen)
{
	try
	{
		// Cập nhật ngày 06/08/2020
		CString szLink = f_Get_DownloadTieuchuan(sTenfile);

		if (szLink == L"")
		{
			if (iOpen == 1) MessageBox(NULL, L"Tiêu chuẩn chưa được cập nhật trên sever.", L"Thông báo", MB_OK | MB_ICONWARNING);
			return L"";
		}

		// Decode
		CString decoSever = f_B64DecodeEX(_urlSever);
		CString decoCheckConn = f_B64DecodeEX(_WcheckConn);
		CString decoWSearch = f_B64DecodeEX(_Wsearch);
		CString decoKeySearch = f_B64DecodeEX(_KeySearch);
		CString decoDoc = f_B64DecodeEX(_KeyDoc);
		CString decoPdf = f_B64DecodeEX(_KeyPdf);

		int pos = sTenfile.Find(L"(");
		if (pos >= 0) sTenfile = sTenfile.Left(pos);
		sTenfile.Trim();
		if (sTenfile == L"") return L"";

		CString sTenGoogle = sTenfile;
		CString sGGSearch = sTenfile;

		sTenfile.Replace(L":", L"-");
		sTenfile.Replace(L"/", L"-");
		sTenfile.Replace(L"|", L"-");
		sTenfile.Replace(L" ", L"-");
		sTenfile.Replace(L"---", L"-");
		sTenfile.Replace(L"--", L"-");

		CString szval = sTenfile;
		sTenfile.Format(L"%s%s", _dfTiento, szval);

		CString pathDir = L"";
		CString pathEXE = _fGPathFolder(L"ToolQLCL.dll");

		pathDir.Format(L"%s\\%s", pathEXE, _dFileTC);
		pathDir.Replace(L"\\\\", L"\\");
		IsFileExist(pathDir);

		CString sBrowse = f_read_txt(pathDir);
		sBrowse.Trim();

		szval = L"";
		if (sBrowse.GetLength() > 0)
		{
			szval = sBrowse.Left(1);
			szval.MakeLower();
		}

		if (sBrowse == L"" || (szval != L"c" && szval != L"d"
			&& szval != L"e" && szval != L"f" && szval != L"h")) sBrowse.Format(L"%s\\%s", pathEXE, _pathTVien);

		sBrowse.Replace(L"\\\\", L"\\");
		if (sBrowse.Right(1) != L"\\") sBrowse += L"\\";

		// Tìm kiếm trong máy tính
		pos = 1;
		CString pathSave = L"";

		pathSave.Format(L"%s%s", sBrowse, szLink);
		if (!FileExists(_ConvertToChar(pathSave))) pos = 0;

		/*pathSave.Format(L"%s%s.doc",sBrowse,sTenfile);
		if(!FileExists(_ConvertToChar(pathSave)))
		{
			pathSave.Format(L"%s%s.pdf",sBrowse,sTenfile);
			if(!FileExists(_ConvertToChar(pathSave)))
			{
				pathSave.Format(L"%s%s.docx",sBrowse,sTenfile);
				if(!FileExists(_ConvertToChar(pathSave))) pos=0;
			}
		}*/

		// Tìm thấy thì mở, không thấy thì tìm trên mạng
		if (pos == 0)
		{
			// Kiểm tra kết nối mạng
			if (InternetCheckConnection(decoCheckConn, FLAG_ICC_FORCE_CONNECTION, 0))
			{
				// Số lượng kiểm tra lượt download tiêu chuẩn
				int iSLDown = 10;
				if (iCheckDownTC >= iSLDown)
				{
					AFX_MANAGE_STATE(AfxGetStaticModuleState());
					Dlg_ktrDownload _dlg;
					_dlg.DoModal();
				}

				if (iCheckDownTC >= iSLDown) return L"";
				iCheckDownTC++;

				// Tìm kiếm trên sever GXD
				CString sURL = L"";

				sURL.Format(L"%s%s", decoSever, szLink);
				pathSave.Format(L"%s%s", sBrowse, szLink);
				URLDownloadToFileW(0, sURL, pathSave, 0, 0);

				pos = 1;
				if (!FileExists(_ConvertToChar(pathSave))) pos = 0;

				/*sURL.Format(L"%s%s.doc",decoSever,sTenfile);
				pathSave.Format(L"%s%s.doc",sBrowse,sTenfile);
				URLDownloadToFileW(0,sURL,pathSave,0,0);

				pos=1;
				if(!FileExists(_ConvertToChar(pathSave)))
				{
					sURL.Format(L"%s%s.pdf",decoSever,sTenfile);
					pathSave.Format(L"%s%s.pdf",sBrowse,sTenfile);
					URLDownloadToFileW(0,sURL,pathSave,0,0);
					if(!FileExists(_ConvertToChar(pathSave)))
					{
						sURL.Format(L"%s%s.docx",decoSever,sTenfile);
						pathSave.Format(L"%s%s.docx",sBrowse,sTenfile);
						URLDownloadToFileW(0,sURL,pathSave,0,0);
						if(!FileExists(_ConvertToChar(pathSave))) pos=0;
					}
				}*/

				// Nếu không tìm thấy (hoặc download về không thành công) --> Tìm kiếm tiếp trên Google
				if (pos == 1)
				{
					if(iOpen == 1) ShellExecute(NULL, L"open", pathSave, NULL, NULL, SW_SHOWMAXIMIZED);
					return pathSave;
				}
				else
				{
					if (iOpen != 1) return L"";

					//if(_yn(L"Tiêu chuẩn không có sẵn trong thư viện. "
					//	L"Bạn có thể nhấn 'Yes' và chờ trong giây lát để tải tiêu chuẩn xuống.",0)!=6) return;

/////////////////////////////////////////////////////////////////////////////////////////////////////////

					int icheckNew = -1;
					int checkDoc = -1, checkPDF = -1;

					CString sLink = L"", sDuoif = L"";
					szval.Format(L" %s", decoKeySearch);
					sTenGoogle += szval;
					sTenGoogle.Replace(L" ", L"%20");
					sLink.Format(L"%s%s", decoWSearch, sTenGoogle);

					// Trường hợp tìm trang đầu tiên không thấy có thể tìm sang trang tiếp theo bằng cách:
					// Chạy vòng for (i=0 --> N)
					// szval.Format(L"%s&start=%d",sLink,10*i);
					// sURL = _ReadingWeb(szval);				

					sURL = _ReadingWeb(sLink);
					if (sURL != _dfError)
					{
						while (true)
						{
							icheckNew = sURL.Find(L"a href=\"/url?q=");
							if (icheckNew == -1) break;

							sURL = sURL.Right(sURL.GetLength() - icheckNew - 15);
							szval = sURL;

							/*checkDoc = sURL.Find(decoDoc);
							checkPDF = sURL.Find(decoPdf);

							if(checkDoc==-1 && checkPDF==-1) break;
							else if(checkDoc==-1 && checkPDF>=0) pos = checkPDF;
							else if(checkDoc>=0 && checkPDF==-1) pos = checkDoc;
							else
							{
								pos = checkDoc;
								if(checkPDF < pos) pos = checkPDF;
							}

							sURL = sURL.Right(sURL.GetLength()-pos-12);
							szval = sURL;*/

							pos = szval.Find(L"http");
							szval = szval.Right(szval.GetLength() - pos);
							pos = szval.Find(L"&amp");
							szval = szval.Left(pos);
							szval.Trim();

							szval = _ReplateWebtoUTF8(szval);

							for (int j = szval.GetLength() - 1; j >= 0; j--)
							{
								if (szval.Mid(j, 1) == L".")
								{
									sDuoif = szval.Right(szval.GetLength() - j - 1);
									if (sDuoif == L"doc" || sDuoif == L"pdf" || sDuoif == L"docx")
									{
										pathSave.Format(L"%s%s.%s", sBrowse, sTenfile, sDuoif);
										URLDownloadToFileW(0, szval, pathSave, 0, 0);
										if (FileExists(_ConvertToChar(pathSave)))
										{
											if (iOpen == 1) ShellExecute(NULL, L"open", pathSave, NULL, NULL, SW_SHOWMAXIMIZED);
											return pathSave;
										}
									}

									break;
								}
							}
						}
					}

					// Trường hợp không tìm thấy hoặc không thể tải xuống --> Bật Web Google
					f_SearchGoogle(decoKeySearch, sGGSearch, 1);

					/////////////////////////////////////////////////////////////////////////////////////////////////////////

				}
			}
			else
			{
				if (iOpen == 1) _s(L"Bạn chưa kết nối Internet. Vui lòng kiểm tra lại!");
			}
		}
		else
		{
			if (iOpen == 1) ShellExecute(NULL, L"open", pathSave, NULL, NULL, SW_SHOWMAXIMIZED);
			return pathSave;
		}

		return pathSave;
	}
	catch (...) {}
	return L"";
}

char *_ConvertToChar(const CString &s)
{
	int nSize = s.GetLength();
	char *pAnsiString = new char[nSize+1];
	memset(pAnsiString,0,nSize+1);
	wcstombs(pAnsiString,s,nSize+1);
	return pAnsiString;
}


CString _ReadingWeb(CString sLinkW)
{
	try
	{
		WSADATA wsaData;
		if ( WSAStartup(0x101, &wsaData) != 0) return _dfError;

		long fileSize=0;
		char* memBuffer = _ReadURL2(sLinkW,fileSize);
		if(strcmp(memBuffer,"")==0) return _dfError;
		
		CString szval=L"";
		int wchars_num = MultiByteToWideChar(CP_UTF8,0,memBuffer,-1,NULL,0);
		wchar_t* wstr = new wchar_t[wchars_num];
		MultiByteToWideChar(CP_UTF8,0,memBuffer,-1,wstr,wchars_num);
		szval.Format(L"%s",wstr);

		if (fileSize != 0) delete(memBuffer);
		WSACleanup();		
		
		return szval;
	}
	catch(...){return _dfError;}
	return _dfError;
}

SOCKET _ConnectToServer(char *szServerName, WORD portNum)
{
	try
	{
		struct hostent *hp;
		unsigned int addr;
		struct sockaddr_in server;
		SOCKET conn;

		conn = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP);
		if (conn == INVALID_SOCKET) return NULL;

		if(inet_addr(szServerName)==INADDR_NONE) hp=gethostbyname(szServerName);
		else
		{
			addr=inet_addr(szServerName);
			hp=gethostbyaddr((char*)&addr,sizeof(addr),AF_INET);
		}

		if(hp==NULL)
		{
			closesocket(conn);
			return NULL;
		}

		server.sin_addr.s_addr=*((unsigned long*)hp->h_addr);
		server.sin_family=AF_INET;
		server.sin_port=htons(portNum);
		if(connect(conn,(struct sockaddr*)&server,sizeof(server)))
		{
			closesocket(conn);
			return NULL;
		}

		return conn;
	}
	catch(...){}
}

char* _ReadURL2(CString szUrl, long &bytesReturnedOut)
{
	const int bufSize = 512;
	char readBuffer[bufSize];
	char *tmpResult=NULL, *result=NULL;
	long totalBytesRead, thisReadSize;
	
	CString szval = szUrl;
	szval.Replace(L"https://",L"");
	szval.Replace(L"http://",L"");

	CString sSever=szval, sGET=L"GET /", sHost = L"Host: ", sConnect=L"Connection: close";
	int pos = szval.Find(L"/");
	if(pos>0)
	{
		sSever = szval.Left(pos);
		sGET+=szval.Right(szval.GetLength()-pos-1);
		sHost+= sSever;
	}
	else
	{
		sGET+=L" ";
		sHost+=sSever;
	}

	sGET+= L" HTTP/1.1";
	szval.Format(L"%s\r\n%s\r\n%s\r\n\r\n",sGET,sHost,sConnect);
	char *sendBuffer = _ConvertToChar(szval);

	SOCKET conn = _ConnectToServer(_ConvertToChar(sSever), 80);
	send(conn, sendBuffer, strlen(sendBuffer), 0);

	totalBytesRead = 0;	
	while(1)
	{
		memset(readBuffer, 0, bufSize);
		thisReadSize = recv (conn, readBuffer, bufSize, 0);		

		if( thisReadSize <= 0 ) break;

		tmpResult = (char*)realloc(tmpResult, thisReadSize+totalBytesRead);

		memcpy(tmpResult+totalBytesRead, readBuffer, thisReadSize);
		totalBytesRead += thisReadSize;
	}

	result = new char[totalBytesRead+1];
	memcpy(result, tmpResult, totalBytesRead);
	result[totalBytesRead] = 0x0;
	
	bytesReturnedOut = totalBytesRead;
	closesocket(conn);
	return(result);
}

CString _ConvertCharsToCstring(char* chr)
{
	int num = MultiByteToWideChar(CP_UTF8,0,chr,-1,NULL,0);
	wchar_t* wstr = new wchar_t[num+1];
	MultiByteToWideChar(CP_UTF8,0,chr,-1,wstr,num);
	wstr[num] = '\0';
	CString szval=L"";
	szval.Format(L"%s",wstr);
	return szval;
}

CString _xlConvertA1toRC(CString sValueA1, int itype)
{
	try
	{
		sValueA1.Trim();
		if(sValueA1==L"") return L"";
		if(_wtoi(sValueA1.Right(1))==0 && sValueA1.Right(1)!=L"0") sValueA1+=L"1";

		CString szval = L"";
		CString sr[3]={L"_",L"_",L"A"};
		int num[3]={0,0,1};
		int iRow=1,iColumn=1;

		for (int i = 0; i < sValueA1.GetLength(); i++)
		{
			if(_wtoi(sValueA1.Mid(i,1))>0)
			{
				iRow = _wtoi(sValueA1.Right(sValueA1.GetLength()-i));
				szval = sValueA1.Left(i);				
				
				sr[2] = szval.Right(1);

				int len = szval.GetLength();
				if(len==2) sr[1] = szval.Left(1);
				else if(len==3)
				{
					sr[0] = szval.Left(1);
					sr[1] = szval.Mid(1,1);
				}

				break;
			}
		}

		if(szval!=L"")
		{
			for (int i = 0; i < 3; i++)
			{
				for (int j = 0; j < 27; j++)
				{
					if(sr[i]==sColumnA1[j])
					{
						num[i]=j;
						break;
					}
				}
			}

			int iColumn = 676*num[0]+26*num[1]+num[2];			
			if(itype==0) szval.Format(L"R%dC%d",iRow,iColumn);
			else if(itype==1) szval.Format(L"%d",iRow);
			else if(itype==2) szval.Format(L"%d",iColumn);
			return szval;
		}
	}
	catch(...){}
	return L"";
}


int _DayValue(int year, int month, int day)
{
	if (month < 3)
	{
		year--;
		month += 12;
	}
	return 365 * year + year / 4 - year / 100 + year / 400 + (153 * month - 457) / 5 + day - 306;
}

int _NumDay2(CString date1, CString date2, int cong)
{
	int d1 = _wtoi(date1.Left(2));
	int m1 = _wtoi(date1.Mid(3, 2));
	int y1 = 2000 + _wtoi(date1.Right(2));

	int d2 = _wtoi(date2.Left(2));
	int m2 = _wtoi(date2.Mid(3, 2));
	int y2 = 2000 + _wtoi(date2.Right(2));

	int numberOfDays = _DayValue(y1, m1, d1) - _DayValue(y2, m2, d2);
	return numberOfDays + cong;
}

int _FileExistsUnicode(CString pathFile)
{
	try
	{
		LPTSTR lpPath = pathFile.GetBuffer(MAX_PATH);
		if (PathFileExists(lpPath)) return 1;
		return -1;
	}
	catch (...){}
	return 0;
}

int _xlGetAllHyperLink(RangePtr pRgSelect, vector<CString> &vecHyper)
{
	try
	{
		vecHyper.clear();
		int ncount = (int)pRgSelect->Hyperlinks->GetCount();
		if (ncount > 0)
		{
			CString szlink = L"", szsub = L"";
			for (int i = 1; i <= ncount; i++)
			{
				szlink = (LPCTSTR)pRgSelect->Hyperlinks->GetItem(i)->GetAddress();
				szsub = (LPCTSTR)pRgSelect->Hyperlinks->GetItem(1)->GetSubAddress();
				if (szsub != L"") szlink += (L"#" + szsub);

				szlink.Replace(L"/", L"\\");

				vecHyper.push_back(szlink);
			}
		}
		return (int)vecHyper.size();
	}
	catch (...) {}
	return 0;
}