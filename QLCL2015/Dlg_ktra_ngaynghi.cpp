// Dlg_ktra_ngaynghi.cpp : implementation file
//

#include "stdafx.h"
#include "Dlg_ktra_ngaynghi.h"
#include "afxdialogex.h"


// Dlg_ktra_ngaynghi dialog

IMPLEMENT_DYNAMIC(Dlg_ktra_ngaynghi, CDialog)

Dlg_ktra_ngaynghi::Dlg_ktra_ngaynghi(CWnd* pParent /*=NULL*/)
	: CDialog(Dlg_ktra_ngaynghi::IDD, pParent)
{

}

Dlg_ktra_ngaynghi::~Dlg_ktra_ngaynghi()
{
}

void Dlg_ktra_ngaynghi::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, KTNTG_CK1, m_chk1);
	DDX_Control(pDX, KTNTG_CK2, m_chk2);
	DDX_Control(pDX, KTNTG_CK3, m_chk3);
	DDX_Control(pDX, KTNTG_CK4, m_chk4);
	DDX_Control(pDX, KTNTG_RD1, m_rad1);
	DDX_Control(pDX, KTNTG_RD2, m_rad2);
	DDX_Control(pDX, KTNTG_S1, m_stt1);
}

BEGIN_MESSAGE_MAP(Dlg_ktra_ngaynghi, CDialog)
	ON_BN_CLICKED(IDC_BUTTON1, &Dlg_ktra_ngaynghi::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON3, &Dlg_ktra_ngaynghi::OnBnClickedButton3)
	ON_BN_CLICKED(IDC_BUTTON2, &Dlg_ktra_ngaynghi::OnBnClickedButton2)
	ON_BN_CLICKED(IDC_BUTTON6, &Dlg_ktra_ngaynghi::OnBnClickedButton6)
	ON_BN_CLICKED(KTNTG_RD1, &Dlg_ktra_ngaynghi::OnBnClickedRd1)
	ON_BN_CLICKED(KTNTG_RD2, &Dlg_ktra_ngaynghi::OnBnClickedRd2)
	ON_BN_CLICKED(KTNTG_CK1, &Dlg_ktra_ngaynghi::OnBnClickedCk1)
	ON_BN_CLICKED(KTNTG_CK2, &Dlg_ktra_ngaynghi::OnBnClickedCk2)
	ON_BN_CLICKED(KTNTG_CK3, &Dlg_ktra_ngaynghi::OnBnClickedCk3)
	ON_BN_CLICKED(KTNTG_CK4, &Dlg_ktra_ngaynghi::OnBnClickedCk4)
END_MESSAGE_MAP()


// Dlg_ktra_ngaynghi message handlers
BOOL Dlg_ktra_ngaynghi::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	CFont m_Font;
	m_Font.CreateFont( 20, 0, 0, 0, FW_BOLD, false, false,
		0, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY,
		FIXED_PITCH|FF_MODERN, _T("MS Shell Dlg") );
	m_stt1.SubclassDlgItem(KTNTG_S1,this);
	m_stt1.SetTextColor(clrKT_Txt);
	m_stt1.SetBkColor(clrKT_Bg);
	m_stt1.SetFont(&m_Font,TRUE);

	m_chk1.SetCheck(ichk7);
	m_chk2.SetCheck(ichkCN);
	m_chk3.SetCheck(ichkLe);	
	m_chk4.SetCheck(ichkCmt);	

	// Xác định vùng được lựa chọn
	pSheet = pWb->ActiveSheet;	
	_irow = pSheet->Application->ActiveCell->Row;
	jbd = _irow, jkt= _irow;

	RangePtr PRc = pSheet->Application->Selection;
	_bstr_t _bsArr = PRc->GetAddress(1,1,XlReferenceStyle::xlR1C1);
	int _nSelection = _fGetSelection((LPCTSTR)_bsArr);
	if(_nSelection>0)
	{
		int _pos = _arrSLS[1].Find(L"-");
		if(_pos>0)				
		{
			// Có nhiều lựa chọn
			jrdAll=0;
			m_rad1.SetCheck(1);
			jbd = _wtoi(_arrSLS[1].Left(_pos));
			jkt = _wtoi(_arrSLS[1].Right(_arrSLS[1].GetLength()-_pos-1));
		}
		else
		{
			jrdAll=1;
			m_rad2.SetCheck(1);
		}
	}

	return TRUE; 
}

BOOL Dlg_ktra_ngaynghi::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		CDialog::OnCancel();
		return TRUE; 
	}

	return FALSE;
}


void Dlg_ktra_ngaynghi::OnBnClickedButton1()
{
	ichk7 = m_chk1.GetCheck();
	ichkCN = m_chk2.GetCheck();
	ichkLe = m_chk3.GetCheck();

	if(ichk7==0 && ichkCN==0 && ichkLe==0)
	{
		_s(L"Bạn chưa lựa chọn ngày cần kiểm tra.");
		return;
	}

	CDialog::OnOK();

	// Run -->
	f_Kiemtra_ngay();
}


void Dlg_ktra_ngaynghi::f_Kiemtra_ngay()
{
	try
	{
		_code = pSheet->CodeName;
		if(_code==(_bstr_t)"shHSNTVL" || _code==(_bstr_t)"shHSNTCV" || _code==(_bstr_t)"shHSNTGD")
		{
			int slcot=0;
			int _jcot[10] = {1,0,0,0,0,0,0,0,0,0};

			if(_code==(_bstr_t)"shHSNTVL")
			{
				slcot=6;
				_jcot[1] = getColumn(pSheet,"DMVL_NK_NG");
				_jcot[2] = getColumn(pSheet,"DMVL_LM_NG");
				_jcot[3] = getColumn(pSheet,"DMVL_LM_KQ");
				_jcot[4] = getColumn(pSheet,"DMVL_NTNB_NG");
				_jcot[5] = getColumn(pSheet,"DMVL_YC");
				_jcot[6] = getColumn(pSheet,"DMVL_AB_NG");
			}
			else if(_code==(_bstr_t)"shHSNTCV")
			{
				slcot=8;
				_jcot[1] = getColumn(pSheet,"DMBB_TN_NGAY");
				_jcot[2] = getColumn(pSheet,"DMBB_NB_NGAY");
				_jcot[3] = getColumn(pSheet,"DMBB_PHIEUYC");
				_jcot[4] = getColumn(pSheet,"DMBB_AB_NGAY");
				_jcot[5] = getColumn(pSheet,"DMBB_HK_NGAYBD");
				_jcot[6] = getColumn(pSheet,"DMBB_HK_NGAYKT");
				_jcot[7] = getColumn(pSheet,"DMBB_DKDBT_NGAY");
				_jcot[8] = getColumn(pSheet,"DMBB_TDDBT_NGAY");
			}
			else if(_code==(_bstr_t)"shHSNTGD")
			{
				slcot=3;
				_jcot[1] = getColumn(pSheet,"DMGD_NTNB_NG");
				_jcot[2] = getColumn(pSheet,"DMGD_YC");
				_jcot[3] = getColumn(pSheet,"DMGD_AB_NG");
			}

			// Xác định vùng kiểm tra 
			if(jrdAll==1)
			{
				jbd=8;
				jkt = FindComment(pSheet,_jcot[0],"END");
			}

			long gtri=0;
			if(ichkLe==1)
			{
				// Đọc từ file MDB
				gtri = f_Read_Nhatky();

				// Đọc file dữ liệu csv
				CString spath=L"";
				spath.Format(L"%s\\%s",_fGPathFolder(L"ToolQLCL.dll"),_dFileNle);
				spath.Replace(L"\\\\",L"\\");
				gtri += f_Read_file_CSV(spath,gtri);
			}

			ichkCmt = m_chk4.GetCheck();

			// Kiểm tra ngày
			int _iktra=0,_islerror=0;
			int _ptram=1,dem=0;
			int _tong= jkt-jbd+1;
			CString sngay=L"",szval=L"";

			RangePtr PRS;
			pRange = pSheet->Cells;
			for (int i = jbd; i <= jkt; i++)
			{
				dem++;
				_ptram = (int)(100*dem/_tong);
				szval.Format(L"Đang tiến hành kiểm tra ngày tháng. Vui lòng đợi trong giây lát... (%d%s)",_ptram,L"%");		
				xl->PutStatusBar((_bstr_t)szval);

				for (int j = 1; j <= slcot; j++)
				{
					PRS = pRange->GetItem(i,_jcot[j]);
					PRS->ClearComments();

					_iktra = 0;
					sngay = GIOR(i,_jcot[j],pRange,L"");
					if(sngay!=L"")
					{						
						if(ichk7==1 || ichkCN==1)
						{
							/*// Nghỉ cuối tuần
							if(ichk7==0 && ichkCN==1)
								szval.Format(L"=IFERROR(CHOOSE(WEEKDAY(%s),\"CN\"),\"\")",GAR(i,_jcot[j],pRange,0));
							else if(ichk7==1 && ichkCN==0)
								szval.Format(
									L"=CHOOSE(WEEKDAY(%s),\"\",\"\",\"\",\"\",\"\",\"\",\"T7\")",
									GAR(i,_jcot[j],pRange,0));
							else if(ichk7==1 && ichkCN==1)
								szval.Format(
									L"=CHOOSE(WEEKDAY(%s),\"CN\",\"\",\"\",\"\",\"\",\"\",\"T7\")",
									GAR(i,_jcot[j],pRange,0));	

							pRange->PutItem(1,99,(_bstr_t)szval);
							szval = GIOR(1,99,pRange,L"");*/

							szval = fThuMay(sngay);
							if ((ichk7 == 1 && szval == L"T7") || (ichkCN == 1 && szval == L"CN"))
							{
								_iktra=1;
								_islerror++;
								PRS->Interior->PutColor(clrKT_Bg);
								PRS->Font->PutColor(clrKT_Txt);
							}							
						}

						if(_iktra==0)
						{
							// Kiểm tra tiếp xem có phải ngày nghỉ lễ tết không?
							if(gtri>0)
							{
								for (int k = 1; k <= gtri; k++)
								{
									if(sngay==sNgayle[k])
									{
										_iktra=1;
										_islerror++;
										PRS->Interior->PutColor(clrKT_Bg);
										PRS->Font->PutColor(clrKT_Txt);
										if(ichkCmt==1 && sGchu[k]!=L"") PRS->AddComment((_bstr_t)sGchu[k]);
										break;
									}
								}
							}
						}

						if (_iktra == 0)
						{
							PRS->Interior->PutColor(xlNone);
							PRS->Font->PutColor(RGB(0, 0, 0));
						}
					}
				}
			}

			if(_islerror>0)
			{
				szval.Format(L"Tồn tại (%d) ngày nhập rơi vào cuối tuần hoặc lễ tết.",_islerror);
				MessageBox(szval,L"Cảnh báo",MB_OK | MB_ICONWARNING);
			}
			else _s(L"Không ngày nào rơi vào cuối tuần hoặc nghỉ lễ tết!");

			//pRange->PutItem(1,99,(_bstr_t)L"");
			xl->PutStatusBar((_bstr_t)L"Ready");
		}
		else _s(L"Tính năng chỉ áp dụng cho các bảng tính danh mục nghiệm thu.");
	}
	catch(...){_s(L"#ER56: Lỗi kiểm tra ngày tháng.");}
}


long Dlg_ktra_ngaynghi::f_Read_Nhatky()
{
	try
	{
		szCheckNgay=L"";
		if(getPathNHKY()==L"") return 0;
		if(connectDsn(pathNK)==-1) return 0;
		
		CConnection ObjConn;
		ObjConn.OpenAccessDB(pathNK, L"", L"");
		CRecordset* Recset = ObjConn.Execute(L"SELECT * FROM tbBoqua;");

		int dem=0;
		int icheck=0;
		CString szval=L"";
		while( !Recset->IsEOF() )
		{
			icheck=0;
			Recset->GetFieldValue(L"ngayghi",szval);

			if(szCheckNgay!=L"")
			{
				if(szCheckNgay.Find(szval)>=0) icheck=1;
			}

			if(icheck==0)
			{
				dem++;
				sNgayle[dem] = szval;
				Recset->GetFieldValue(L"ghichu",sGchu[dem]);

				szCheckNgay+=sNgayle[dem];
				szCheckNgay+=L"|";
			}

			Recset->MoveNext();
		}
		Recset->Close();
		ObjConn.CloseDatabase();
		
		return dem;
	}
	catch(...){}
	return 0;
}


long Dlg_ktra_ngaynghi::f_Read_file_CSV(CString pathCSV, long num)
{
	int icheck=0;
	long cong=num;

	try
	{
		wchar_t **dArray;
		int numberOfLine = 0;
		int i;
	
		long _size;
		_size= f_GetLineFile(pathCSV);
		_size++;
		dArray = (wchar_t **)calloc(_size,sizeof(wchar_t *));
		wchar_t **p;
		p = (wchar_t **)calloc(_size,sizeof(wchar_t *));
		for (i = 0; i < _size; i++)
		{
			*(dArray+i) = (wchar_t *)malloc(SIZE_LINE*sizeof(wchar_t));
			*(p+i) = (wchar_t *)malloc(SIZE_LINE*sizeof(wchar_t));
		}

		FILE *pReadFile;
		if(_wfopen_s(&pReadFile, pathCSV, L"r, ccs=UTF-16LE") != 0)
		{
			_s(L"Không thể mở được file dữ liệu. Vui lòng chọn lại file dữ liệu.");
			for (i = 0; i < _size; i++)
			{
				free(*(dArray+i));
			}
			free(dArray);
			return 0;
		}	
		else
		{
			i=0;
			CString szVarString=L"";
			while(fgetws(*(p+i), ROW_LINE, pReadFile)) i++;
			
			numberOfLine= i;
			fclose(pReadFile);

			int fdem=0;
			CString ft[2]={L"",L""};

			for(int i= 1;i<numberOfLine;i++)
			{
				wcscpy_s(*(dArray+i), wcslen(*(p+i))+1, *(p+i));
				wchar_t *textString = new wchar_t[SIZE_LINE];
				wcscpy_s(textString,wcslen(*(dArray+i))+1, *(dArray+i));
				szVarString= textString;

				long length;
				CString szLeft;
				int nToken= szVarString.Find(L'\t');

				fdem=0;
				while(nToken>=0)
				{
					length= szVarString.GetLength();
					szLeft= szVarString.Left(nToken);
					szLeft.TrimLeft();szLeft.TrimRight();
					szVarString=szVarString.Right(length-nToken-1);
					nToken= szVarString.Find(L'\t');

					fdem++;
					if(fdem==1) ft[0]=szLeft;

					// Thêm vào để lấy cột cuối cùng (by Leo 23.06.15)
					ft[1]=L"";
					if(nToken<0) ft[1]=szVarString;
				}

				ft[0].Trim();
				ft[0].Replace(L" ",L"");
				ft[0].Replace(L"N",L"");
				if(ft[0]!=L"")
				{
					icheck=0;
					if(szCheckNgay!=L"")
					{
						if(szCheckNgay.Find(ft[0])>=0) icheck=1;
					}

					if(icheck==0)
					{
						cong++;
						sNgayle[cong] = ft[0];
						sGchu[cong] = ft[1];
						sGchu[cong].Trim();
					}
				}
				free(textString);
			}
			fclose(pReadFile);
		}
	
		//free memory
		for (i = 0; i < _size; i++)
		{
			free(*(dArray+i));
			free(*(p+i));
		}
		free(dArray);
		free(p);
	}
	catch(_com_error & error){_s(L"#QL4.126: Lỗi đọc dữ liệu ngày tháng nghỉ lễ.");}

	return cong;
}


long Dlg_ktra_ngaynghi::f_GetLineFile(CString szPath)
{
	FILE * pFile;
	wchar_t *c = (wchar_t *)malloc(2000*sizeof(wchar_t));
	long line = 0;
	pFile=_wfopen(szPath,L"r, ccs=UTF-16LE");
	if (pFile==NULL) perror ("Error opening file");
	else
	{
		while (fgetws(c, 2000, pFile))
		{
			line++;
		}
		fclose (pFile);
		printf ("File contains %d.\n",line);
	}
	free(c);
	return line;
}


void Dlg_ktra_ngaynghi::OnBnClickedButton3()
{
	CDialog::OnCancel();
}


void Dlg_ktra_ngaynghi::OnBnClickedButton2()
{
	CColorDialog dlg; 
	if (dlg.DoModal() == IDOK)
	{
		clrKT_Bg = dlg.GetColor();
		m_stt1.SetBkColor(clrKT_Bg);
	}
}


void Dlg_ktra_ngaynghi::OnBnClickedButton6()
{
	CColorDialog dlg; 
	if (dlg.DoModal() == IDOK)
	{
		clrKT_Txt = dlg.GetColor();
		m_stt1.SetTextColor(clrKT_Txt);
	}
}


void Dlg_ktra_ngaynghi::OnBnClickedRd1()
{
	jrdAll=0;
}


void Dlg_ktra_ngaynghi::OnBnClickedRd2()
{
	jrdAll=1;
}


void Dlg_ktra_ngaynghi::OnBnClickedCk1()
{
	ichk7 = m_chk1.GetCheck();
}


void Dlg_ktra_ngaynghi::OnBnClickedCk2()
{
	ichkCN = m_chk2.GetCheck();
}


void Dlg_ktra_ngaynghi::OnBnClickedCk3()
{
	ichkLe = m_chk3.GetCheck();
}


void Dlg_ktra_ngaynghi::OnBnClickedCk4()
{
	ichkCmt = m_chk4.GetCheck();
}
