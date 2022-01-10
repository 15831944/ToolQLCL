#include "stdafx.h"
#include "Dlg_Chonngay_paste.h"
#include "Dlg_Timer.h"


// Dlg_Chonngay_paste dialog
IMPLEMENT_DYNAMIC(Dlg_Chonngay_paste, CDialog)

Dlg_Chonngay_paste::Dlg_Chonngay_paste(CWnd* pParent /*=NULL*/)
	: CDialog(Dlg_Chonngay_paste::IDD, pParent)
{
	sSaochepngay = L"";
}

Dlg_Chonngay_paste::~Dlg_Chonngay_paste()
{
}

void Dlg_Chonngay_paste::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, CHNP_S1, stt1);
	DDX_Control(pDX, CHNP_S2, stt2);
	DDX_Control(pDX, CHNP_T1, m_timeDlg[2]);	
}


BEGIN_MESSAGE_MAP(Dlg_Chonngay_paste, CDialog)
	ON_BN_CLICKED(IDOK, &Dlg_Chonngay_paste::OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, &Dlg_Chonngay_paste::OnBnClickedCancel)
	ON_BN_CLICKED(IDCANCEL2, &Dlg_Chonngay_paste::OnBnClickedCancel2)
END_MESSAGE_MAP()


// Dlg_Chonngay_paste message handlers
BOOL Dlg_Chonngay_paste::OnInitDialog()
{
	CDialog::OnInitDialog();
	// Load Icon dự toán
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);
	
	m_timeDlg[2].SetCueBanner(L"dd/mm/yyyy");

	_bstr_t shNK = name_sheet((_bstr_t)"shMauNKY4");
	_WorksheetPtr psNK = xl->Sheets->GetItem(shNK);

	_ngbd = getVCell(psNK,L"HK_NBD");
	_ngkt = getVCell(psNK,L"HK_NKT");

	stt1.SetWindowText(L"Ngày bắt đầu: "+ _ngbd);
	stt2.SetWindowText(L"Ngày kết thúc: "+ _ngkt);

	return TRUE; 
}


void Dlg_Chonngay_paste::OnBnClickedOk()
{
	int iktr=1;
	CString szval=L"";
	m_timeDlg[2].GetWindowTextW(szval);
	szval.Trim();
	
	int numd = szval.GetLength();
	if(numd!=8 && numd!=10) iktr=0;
	if(szval.Mid(2,1)!=L"/" || szval.Mid(5,1)!=L"/") iktr=0;
	if(iktr==1)
	{
		// Kiểm tra ngày nhập có nằm trong khoảng giới hạn không?
		if(compare_date(numd,_ngbd,szval)==0)
		{
			_s(L"Ngày nhập phải sau ngày "+ _ngbd);
			return;
		}

		if(compare_date(numd,szval,_ngkt)==0)
		{
			_s(L"Ngày nhập phải trước ngày "+ _ngkt);
			return;
		}

		// Kiểm tra dữ liệu đã tồn tại chưa?
		if(iKqua>0)
		{
			for (int i = 1; i <= iKqua; i++)
			{
				if(szval==XLXD[i].CDG[0])
				{
					_s(L"Ngày nhập đã tồn tại dữ liệu. Vui lòng kiểm tra lại.");
					return;
				}
			}
		}

		// Cảnh báo ngày lễ
		long gtri=0;
		if(ichkLe==1)
		{
			// Đọc file dữ liệu csv
			CString spath=L"";
			spath.Format(L"%s\\%s",_fGPathFolder(L"ToolQLCL.dll"),_dFileNle);
			spath.Replace(L"\\\\",L"\\");
			gtri = f_Read_file_CSV(spath);

			// Kiểm tra tiếp xem có phải ngày nghỉ lễ tết không?
			if(gtri>0)
			{
				for (int k = 1; k <= gtri; k++)
				{
					if(szval==sNgayle[k])
					{
						spath.Format(
							L"Ngày bạn nhập là ngày nghỉ lễ. Bạn vẫn muốn tiếp tục? \n[%s] %s",
							sNgayle[k],sGchu[k]);

						if(_yn(spath,0)!=6) return;

						break;
					}
				}
			}
		}

		CDialog::OnOK();
		sSaochepngay = szval;
	}
	else _s(L"Bạn cần nhập đúng định dạng ngày tháng. (dd/mm/yyyy)");
}


long Dlg_Chonngay_paste::f_Read_file_CSV(CString pathCSV)
{
	long cong=0;

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
					cong++;
					sNgayle[cong] = ft[0];
					sGchu[cong] = ft[1];
					sGchu[cong].Trim();
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


long Dlg_Chonngay_paste::f_GetLineFile(CString szPath)
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


void Dlg_Chonngay_paste::OnBnClickedCancel()
{
	// TODO: Add your control notification handler code here
	CDialog::OnCancel();
}


void Dlg_Chonngay_paste::OnBnClickedCancel2()
{
	_iTimeDlg=3;
	AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
	Dlg_Timer _dlg;
	_dlg.DoModal();
}
