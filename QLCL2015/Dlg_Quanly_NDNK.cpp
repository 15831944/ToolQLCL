#include "Dlg_Quanly_NDNK.h"
#include "Dlg_Chonngay_paste.h"
#include "Dlg_Nky_Noidung.h"
#include "Dlg_Nky_Thoitiet.h"
#include "Dlg_Nky_BangNLTB.h"
#include "Dlg_Nky_Hinhanh.h"
#include "Dlg_Timer.h"

#define IDC_COLUMN_POPUP			3000
#define IDC_COLUMN_GHICHU			3001
#define IDC_COLUMN_THOITIET			3002
#define IDC_COLUMN_NHANCONG			3003
#define IDC_COLUMN_MAYTHICONG		3004
#define IDC_COLUMN_CHITIEU			3005
#define IDC_COLUMN_VSATLD			3006
#define IDC_COLUMN_NHANXET			3007
#define IDC_COLUMN_HINHANH			3008

// Class tra cứu phụ lục vữa
Dlg_Quanly_NDNK::Dlg_Quanly_NDNK(CWnd* pParent /*=NULL*/)
	: CDialog(Dlg_Quanly_NDNK::IDD, pParent)
{
	
}

void Dlg_Quanly_NDNK::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, QLND_CB1, m_cbbox);
	DDX_Control(pDX, QLND_T1, m_timeDlg[0]);
	DDX_Control(pDX, QLND_T2, m_timeDlg[1]);
	DDX_Control(pDX, QLND_L1, m_listNDNK);
	DDX_Control(pDX, QLND_E1, txtE_NKTC);
	DDX_Control(pDX, QLND_CK2, m_chk2);
	DDX_Control(pDX, QLND_B1, m_btn1);
	DDX_Control(pDX, QLND_B2, m_btn2);
	DDX_Control(pDX, QLND_S3, m_thongbao);	
}

BEGIN_MESSAGE_MAP(Dlg_Quanly_NDNK, CDialog)
	ON_WM_SIZE()
	ON_WM_SIZING()
	ON_WM_SYSCOMMAND()
	ON_WM_CONTEXTMENU()
	ON_BN_CLICKED(QLND_B3, &Dlg_Quanly_NDNK::OnBnClickedB3)
	ON_BN_CLICKED(QLND_B1, &Dlg_Quanly_NDNK::OnBnClickedB1)
	ON_BN_CLICKED(QLND_B2, &Dlg_Quanly_NDNK::OnBnClickedB2)
	ON_BN_CLICKED(QLND_CK2, &Dlg_Quanly_NDNK::OnBnClickedCk2)
	ON_NOTIFY(NM_CLICK, QLND_L1, &Dlg_Quanly_NDNK::OnNMClickL1)
	ON_CBN_SELCHANGE(QLND_CB1, &Dlg_Quanly_NDNK::OnCbnSelchangeCb1)	
	ON_COMMAND(ID_NH40007, &Dlg_Quanly_NDNK::OnNh40007)	
	ON_COMMAND(ID_NH40005, &Dlg_Quanly_NDNK::OnNh40005)
	ON_COMMAND(ID_NH40006, &Dlg_Quanly_NDNK::OnNh40006)
	ON_COMMAND(ID_NH40004, &Dlg_Quanly_NDNK::OnNh40004)
	ON_COMMAND(ID_NH40064, &Dlg_Quanly_NDNK::OnNh40064)
	ON_EN_KILLFOCUS(QLND_E1, &Dlg_Quanly_NDNK::OnEnKillfocusE1)
	ON_COMMAND(ID_X40012, &Dlg_Quanly_NDNK::OnX40012)
	ON_NOTIFY(NM_RCLICK, QLND_L1, &Dlg_Quanly_NDNK::OnNMRClickL1)
	ON_COMMAND(ID_D40053, &Dlg_Quanly_NDNK::OnD40053)
	ON_COMMAND(ID_D40054, &Dlg_Quanly_NDNK::OnD40054)
	ON_COMMAND(ID_D40055, &Dlg_Quanly_NDNK::OnD40055)
	ON_COMMAND(ID_D40057, &Dlg_Quanly_NDNK::OnD40057)
	ON_COMMAND(ID_D40056, &Dlg_Quanly_NDNK::OnD40056)
	ON_COMMAND(ID_D40059, &Dlg_Quanly_NDNK::OnD40059)
	ON_EN_SETFOCUS(QLND_E1, &Dlg_Quanly_NDNK::OnEnSetfocusE1)
	ON_EN_SETFOCUS(QLND_T1, &Dlg_Quanly_NDNK::OnEnSetfocusT1)
	ON_EN_SETFOCUS(QLND_T2, &Dlg_Quanly_NDNK::OnEnSetfocusT2)
	ON_COMMAND(ID_X40014, &Dlg_Quanly_NDNK::OnX40014)
	ON_COMMAND(ID_X40016, &Dlg_Quanly_NDNK::OnX40016)
	ON_COMMAND(ID_X40017, &Dlg_Quanly_NDNK::OnX40017)
	ON_COMMAND(ID_X40018, &Dlg_Quanly_NDNK::OnX40018)
	ON_COMMAND(ID_X40062, &Dlg_Quanly_NDNK::OnX40062)
	ON_COMMAND(ID_X40063, &Dlg_Quanly_NDNK::OnX40063)
	ON_COMMAND(ID_X40058, &Dlg_Quanly_NDNK::OnX40058)	
	ON_NOTIFY(HDN_ENDTRACK, 0, &Dlg_Quanly_NDNK::OnHdnEndtrackL1)

	ON_COMMAND(IDC_COLUMN_POPUP, &Dlg_Quanly_NDNK::OnDwShowAll)
	ON_COMMAND(IDC_COLUMN_GHICHU, &Dlg_Quanly_NDNK::OnDwShowGhichu)
	ON_COMMAND(IDC_COLUMN_THOITIET, &Dlg_Quanly_NDNK::OnDwShowThoitiet)
	ON_COMMAND(IDC_COLUMN_NHANCONG, &Dlg_Quanly_NDNK::OnDwShowNhancong)
	ON_COMMAND(IDC_COLUMN_MAYTHICONG, &Dlg_Quanly_NDNK::OnDwShowMaythicong)
	ON_COMMAND(IDC_COLUMN_CHITIEU, &Dlg_Quanly_NDNK::OnDwShowChitieu)
	ON_COMMAND(IDC_COLUMN_NHANXET, &Dlg_Quanly_NDNK::OnDwShowNhanxet)
	ON_COMMAND(IDC_COLUMN_VSATLD, &Dlg_Quanly_NDNK::OnDwShowVesinhATLD)
	ON_COMMAND(IDC_COLUMN_HINHANH, &Dlg_Quanly_NDNK::OnDwShowPathHinhanh)

	ON_MESSAGE(WM_LISTCTRLEX_HEADERRIGHTCLICK, OnHeaderRightClick)
END_MESSAGE_MAP()

BEGIN_EASYSIZE_MAP(Dlg_Quanly_NDNK)
	// EASYSIZE(control,left,top,right,bottom,options)
	// options = 1 -> không thay đổi, =0 có thay đổi
	// ES_BORDER = ghim (thay đổi theo)
	// ES_KEEPSIZE = giữ nguyên ban đầu

	EASYSIZE(QLND_CK2,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)	
	EASYSIZE(QLND_S2,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)	
	EASYSIZE(QLND_T1,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)	
	EASYSIZE(QLND_T2,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)	
	EASYSIZE(QLND_B1,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)	
	EASYSIZE(QLND_B2,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(QLND_S1,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(QLND_CB1,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(QLND_L1,ES_BORDER,ES_BORDER,ES_BORDER,ES_BORDER,0)
	EASYSIZE(QLND_S3,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)
	EASYSIZE(QLND_B3,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
END_EASYSIZE_MAP

BOOL Dlg_Quanly_NDNK::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);
	//f_Load_ToolTip();
	clr1 = RGB(220,220,220);
	clr2 = BLACK;
	gtlit=0,dwR=0,dwG=0,dwB=0;

	// Đọc từ file MDB
	gtriCSV = f_Read_Nhatky();

	// Đọc file dữ liệu csv
	CString spath=L"";
	spath.Format(L"%s\\%s",_fGPathFolder(L"ToolQLCL.dll"),_dFileNle);
	spath.Replace(L"\\\\",L"\\");
	
	gtriCSV += f_Read_file_CSV(spath,gtriCSV);

	// Xác định ngày tháng (bắt đầu, kết thúc, tổng số ngày, ...)
	f_Xac_dinh_ngay();	

	// Từ khóa hiển thị hướng dẫn
	m_timeDlg[0].SetCueBanner(_T("dd/mm/yyyy"), TRUE);
	m_timeDlg[1].SetCueBanner(_T("dd/mm/yyyy"), TRUE);

	// Load combobox
	m_cbbox.AddString(L"Các ngày đã ghi nhật ký");
	m_cbbox.AddString(L"Các ngày chưa ghi nhật ký");
	m_cbbox.AddString(L"Tất cả các ngày");
	m_cbbox.SetCurSel(2);

	m_timeDlg[0].EnableWindow(0);
	m_timeDlg[1].EnableWindow(0);
	m_btn1.EnableWindow(0);
	m_btn2.EnableWindow(0);

	CString szval=L"";
	szval.Format(L"Quản lý nhật ký từ ngày %s đến ngày %s",sday_bd,sday_kt);
	this->SetWindowTextW(szval);
	
	// Load list & dữ liệu
	f_Style_noidung();

	// Load dữ liệu đại diện cho tất cả các ngày
	f_Load_SQL_ALL();
	f_Ktra_dkien();

	_cpyNThg=L"";

	INIT_EASYSIZE;
	ShowWindow(SW_MAXIMIZE);

	return TRUE; 
}

// Bắt sự kiện click Enter
BOOL Dlg_Quanly_NDNK::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if (pMsg->message == WM_KEYDOWN 
		&& pMsg->wParam == VK_RETURN &&	pWndWithFocus == GetDlgItem(QLND_T1))
	{
		GotoDlgCtrl(GetDlgItem(QLND_T2));
		f_TkiemT1();
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN 
		&& pMsg->wParam == VK_RETURN && pWndWithFocus == GetDlgItem(QLND_T2))
	{
		GotoDlgCtrl(GetDlgItem(QLND_T1));
		f_TkiemT2();
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN 
		&& pMsg->wParam == VK_TAB && pWndWithFocus == GetDlgItem(QLND_T1))
	{
		GotoDlgCtrl(GetDlgItem(QLND_T2));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN 
		&& pMsg->wParam == VK_TAB && pWndWithFocus == GetDlgItem(QLND_T2))
	{
		GotoDlgCtrl(GetDlgItem(QLND_T1));
		return TRUE;
	}
	else if( pMsg->message == WM_LBUTTONDOWN ||
		pMsg->message == WM_LBUTTONUP ||
		pMsg->message == WM_MOUSEMOVE )
	{
		m_ToolTips.RelayEvent(pMsg);
		return CDialog::PreTranslateMessage(pMsg);
	}
	else if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		if(ClickEsc==0)
		{
			f_Get_width();
			CDialog:OnCancel();
		}
		else GotoDlgCtrl(GetDlgItem(QLND_L1));
		ClickEsc=0;
		
		return TRUE; 
	}
	else if (pMsg->message == WM_CHAR)
	{
		TCHAR chr = static_cast<TCHAR>(pMsg->wParam);
		if(pWndWithFocus == GetDlgItem(QLND_L1))
		{
			switch (chr)
			{
				case 0x01:
				{
					// Ctrl+A
					_SelectAllListMain();
					break;
				}

				case 0x03:
				{
					// Ctrl+C
					OnNh40007();
					break;
				}

				case 0x16:
				{
					// Ctrl+V
					OnD40053();
					break;
				}
				case 0x7F:
				{
					// Delete
					OnX40012();
					break;
				}

				default:
					break;
			}
		}		
	}

	return FALSE;
}


void Dlg_Quanly_NDNK::_SelectAllListMain()
{
	int count = m_listNDNK.GetItemCount();
	if(count==0) return;
	for (int i = 0; i < count; i++) m_listNDNK.SetItemState(i, LVIS_SELECTED, LVIS_SELECTED);
}

void Dlg_Quanly_NDNK::OnSize(UINT nType, int cx, int cy) 
{
	CDialog::OnSize(nType, cx, cy);
	UPDATE_EASYSIZE;
}

void Dlg_Quanly_NDNK::OnSizing(UINT fwSide, LPRECT pRect) 
{
	CDialog::OnSizing(fwSide, pRect);
	EASYSIZE_MINSIZE(560,250,fwSide,pRect);	// chiều rộng + chiều cao
}

long Dlg_Quanly_NDNK::f_GetLineFile(CString szPath)
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

long Dlg_Quanly_NDNK::f_Read_Nhatky()
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

long Dlg_Quanly_NDNK::f_Read_file_CSV(CString pathCSV, long num)
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


void Dlg_Quanly_NDNK::f_Xac_dinh_ngay()
{
	try
	{
		shNK = name_sheet((_bstr_t)"shMauNKY4");
		psNK = xl->Sheets->GetItem(shNK);
		prNK = psNK->Cells;

		_ngbd = getVCell(psNK,L"HK_NBD");
		_ngkt = getVCell(psNK,L"HK_NKT");

		int iColumnEnd = psNK->Cells->SpecialCells(xlCellTypeLastCell)->GetColumn()+1;
		RangePtr pRg1 = prNK->GetItem(5,iColumnEnd);
		pRg1->PutNumberFormat("General");

		// Xác định tổng số ngày
		CString szval=L"=HK_NKT-HK_NBD+1";
		prNK->PutItem(5,iColumnEnd,(_bstr_t)szval);
		szval = GIOR(5,iColumnEnd,prNK,L"");
		TongNgay = _wtoi(szval);
		
		pRg1->PutNumberFormat("dd/mm/yyyy");

		NgayAll[1] = _ngbd;
		NgayAll[TongNgay] = _ngkt;

		for (int i = 1; i < TongNgay-1; i++)
		{
			szval.Format(L"=HK_NBD+%d",i);
			prNK->PutItem(5,iColumnEnd,(_bstr_t)szval);
			szval = GIOR(5,iColumnEnd,prNK,L"");
			NgayAll[i+1] = szval;
		}

		m_timeDlg[0].SetWindowText(_ngbd);
		m_timeDlg[1].SetWindowText(_ngkt);

		pRg1->PutNumberFormat("General");
		sday_bd = _ngbd;
		sday_kt = _ngkt;
	}
	catch(...){}
}


void Dlg_Quanly_NDNK::f_Load_ToolTip()
{
	m_ToolTips.Create(this);
	m_ToolTips.SetMaxTipWidth(600);
	m_ToolTips.SetDelayTime(900);
}


void Dlg_Quanly_NDNK::f_Style_noidung()
{
	// Định dạng kiểu List Control
	m_listNDNK.InsertColumn(0,L"",LVCFMT_CENTER,0);
	m_listNDNK.InsertColumn(1,L"Thứ",LVCFMT_CENTER,jW3[1]);
	m_listNDNK.InsertColumn(2,L"Ngày",LVCFMT_CENTER,jW3[2]);
	m_listNDNK.InsertColumn(3,L"Ghi chú",LVCFMT_LEFT,jW3[3]);
	m_listNDNK.InsertColumn(4,L"Thời tiết",LVCFMT_LEFT,jW3[4]);	
	m_listNDNK.InsertColumn(5,L"Nhiệt độ",LVCFMT_LEFT,jW3[5]);
	m_listNDNK.InsertColumn(6,L"Nhân lực",LVCFMT_LEFT,jW3[6]);	
	m_listNDNK.InsertColumn(7,L"Thiết bị thi công",LVCFMT_LEFT,jW3[7]);
	m_listNDNK.InsertColumn(8,L"Biện pháp thi công",LVCFMT_LEFT,jW3[8]);
	m_listNDNK.InsertColumn(9,L"Tình hình vật liệu, cấu kiện",LVCFMT_LEFT,jW3[9]);
	m_listNDNK.InsertColumn(10,L"Chất lượng thi công",LVCFMT_LEFT,jW3[10]);
	m_listNDNK.InsertColumn(11,L"Tiến độ thực hiện",LVCFMT_LEFT,jW3[11]);
	m_listNDNK.InsertColumn(12,L"Công tác vệ sinh",LVCFMT_CENTER,jW3[12]);
	m_listNDNK.InsertColumn(13,L"An toàn lao động",LVCFMT_CENTER,jW3[13]);
	m_listNDNK.InsertColumn(14,L"Nhận xét của giám sát",LVCFMT_LEFT,jW3[14]);
	m_listNDNK.InsertColumn(15,L"Nhận xét của chủ đầu tư",LVCFMT_LEFT,jW3[15]);
	m_listNDNK.InsertColumn(16,L"Ý kiến của nhà thầu",LVCFMT_LEFT,jW3[16]);
	m_listNDNK.InsertColumn(17,L"Hình ảnh đính kèm",LVCFMT_LEFT,0);
	m_listNDNK.InsertColumn(18,L"Màu nền",LVCFMT_CENTER,0);
	m_listNDNK.InsertColumn(19,L"Màu chữ",LVCFMT_CENTER,0);
	m_listNDNK.InsertColumn(20,L"Bỏ qua",LVCFMT_CENTER,0);

	for (int i = 3; i <= 16; i++) m_listNDNK.SetColumnEditor(i,&txtE_NKTC);
	m_listNDNK.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);
}


void Dlg_Quanly_NDNK::f_Get_size()
{
	/*CRect rectCtrl;
	GetWindowRect(&rectCtrl);
	irecH = rectCtrl.Height();
	irecW = rectCtrl.Width();*/
}


void Dlg_Quanly_NDNK::f_Get_width()
{
	for (int i = 0; i < 20; i++) jW3[i] = m_listNDNK.GetColumnWidth(i);		
}


void Dlg_Quanly_NDNK::f_Delete_list()
{
	m_listNDNK.DeleteAllItems();
	ASSERT(m_listNDNK.GetItemCount() == 0);
}


void Dlg_Quanly_NDNK::OnBnClickedB3()
{
	f_Get_width();
	CDialog::OnCancel();
}


void Dlg_Quanly_NDNK::OnContextMenu(CWnd* /*pWnd*/, CPoint point)
{
	try
	{
		if(m_cbbox.GetCurSel()!=0 && nItem==-1) return;

		// Load the desired menu
		CMenu mnu2;
		mnu2.LoadMenu(IDR_MENU2);
		
		// Get a pointer to the button
		CListCtrlEx *pClist;
		pClist = reinterpret_cast<CListCtrlEx *>(GetDlgItem(QLND_L1));

		// Find the rectangle around the button
		CRect rectSubmitButton;
		pClist->GetWindowRect(&rectSubmitButton);
		CString szval=L"";

		// Get a pointer to the first item of the menu
		CMenu *mnuP2 = mnu2.GetSubMenu(0);

		if(m_cbbox.GetCurSel()==0 && nItem==-1)
		{
			mnuP2->EnableMenuItem(ID_NH40004, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			mnuP2->EnableMenuItem(ID_NH40005, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			mnuP2->EnableMenuItem(ID_NH40006, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			mnuP2->EnableMenuItem(ID_NH40007, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			
			mnuP2->EnableMenuItem(ID_X40012, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			mnuP2->EnableMenuItem(ID_X40014, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			mnuP2->EnableMenuItem(ID_X40016, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			mnuP2->EnableMenuItem(ID_X40017, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			mnuP2->EnableMenuItem(ID_X40018, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
		}

		// _cpyNThg=NULL biến chưa được gán sao chép
		if(_cpyNThg==L"")
		{
			mnuP2->EnableMenuItem(ID_D40053, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			mnuP2->EnableMenuItem(ID_D40054, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			mnuP2->EnableMenuItem(ID_D40055, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			mnuP2->EnableMenuItem(ID_D40057, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			mnuP2->EnableMenuItem(ID_D40056, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			mnuP2->EnableMenuItem(ID_D40059, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
		}

		szval = m_listNDNK.GetItemText(nItem,20);
		if(szval==L"1")
		{
			mnuP2->CheckMenuItem(ID_NH40004,MF_CHECKED);
			mnuP2->EnableMenuItem(ID_NH40005, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			mnuP2->EnableMenuItem(ID_NH40006, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
		}

		ASSERT(mnuP2);
		if( rectSubmitButton.PtInRect(point) )
			mnuP2->TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON, point.x, point.y, this);
	}
	catch(...){}
}


int Dlg_Quanly_NDNK::f_Check_nghile(int row)
{
	if(gtriCSV==0) return 0;
	
	CString sngay = m_listNDNK.GetItemText(row,2);
	for (int i = 1; i <= gtriCSV; i++)
	{
		if(sngay==sNgayle[i])
		{
			m_listNDNK.SetRowColors(row,RED,WHITE);
			return 1;
		}
	}

	return 0;
}


int Dlg_Quanly_NDNK::f_Set_batdau(CString sbd)
{
	int ibd=1;
	if(sbd!=L"")
		for (int i = 1; i <= TongNgay; i++)
			if(sbd==NgayAll[i]) return i;
	return ibd;
}


int Dlg_Quanly_NDNK::f_Set_ketthuc(CString skt)
{
	int ikt=TongNgay;
	if(skt!=L"")
		for (int i = TongNgay; i >= 1; i--)
			if(skt==NgayAll[i]) return i;
	return ikt;
}


void Dlg_Quanly_NDNK::f_Load_NULL(CString sdaybd, CString sdaykt)
{
	TRY
	{
		// Xác định vị trí bắt đầu và kết thúc
		int ibd = f_Set_batdau(sdaybd);
		int ikt = f_Set_ketthuc(sdaykt);

		// Xóa dữ liệu cũ (nếu có)
		f_Delete_list();

		// Load dữ liệu và truy vấn
		ObjConn.OpenAccessDB(pathNK, L"", L"");
		CRecordset* Recset;

		gtlit=0;
		CString sDem=L"",sThu=L"";
		for (int i = ibd; i <= ikt; i++)
		{
			// Kiểm tra dữ liệu có tồn tại không?
			SqlString.Format(L"SELECT COUNT(*) AS iDem FROM tbThoitiet WHERE ngayghi LIKE '%s';",NgayAll[i]);
			Recset = ObjConn.Execute(SqlString);
			Recset->GetFieldValue(L"iDem",sDem);
			Recset->Close();

			if(_wtoi(sDem)==0 && szGtri[0]==L"")
			{
				sThu = fThuMay(NgayAll[i]);
				m_listNDNK.InsertItem(gtlit,L"",0);
				m_listNDNK.SetItemText(gtlit,1,sThu);			// Thứ
				m_listNDNK.SetItemText(gtlit,2,NgayAll[i]);		// Ngày

				if(sThu==L"CN") m_listNDNK.SetRowColors(gtlit,clr1,clr2);
				if(f_Check_nghile(gtlit)==1) m_listNDNK.SetRowColors(gtlit,RED,WHITE);
				gtlit++;
			}
		}

		ObjConn.CloseDatabase();
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[x7.51] Lỗi load dữ liệu")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_Quanly_NDNK::f_Load_ALL(CString sdaybd, CString sdaykt, int itype)
{
	TRY
	{
		// Xác định vị trí bắt đầu và kết thúc
		int ibd = f_Set_batdau(sdaybd);
		int ikt = f_Set_ketthuc(sdaykt);

		// Xóa dữ liệu cũ (nếu có)
		f_Delete_list();

		// Load dữ liệu và truy vấn
		ObjConn.OpenAccessDB(pathNK, L"", L"");

		gtlit=0;
		int icheck = 0;
		for (int i = ibd; i <= ikt; i++)
		{
			// Nếu là combobox danh sách chưa có dữ liệu
			icheck = 0;
			if (itype == 1)
			{
				// Kiểm tra có dữ liệu chưa?
				if (CheckDulieu(NgayAll[i]) == 0) icheck = 1;
			}

			if (icheck == 0)
			{
				// Load vào list
				CString sThu = fThuMay(NgayAll[i]);
				m_listNDNK.InsertItem(gtlit, L"", 0);
				m_listNDNK.SetItemText(gtlit, 1, sThu);				// Thứ
				m_listNDNK.SetItemText(gtlit, 2, NgayAll[i]);		// Ngày			

				// Truy vấn dữ liệu
				f_Load_SQL_01(NgayAll[i], L"tbThoitiet", 4, 0);
				f_Load_SQL_01(NgayAll[i], L"tbNhietdo", 5, 1);

				f_Load_SQL_02(NgayAll[i], L"tbNhanluc", L"", 6, 11);
				f_Load_SQL_02(NgayAll[i], L"tbThietbi", L"", 7, 12);

				f_Load_SQL_02(NgayAll[i], L"tbBienphapTC", L"idBienphapTC", 8, 2);
				f_Load_SQL_02(NgayAll[i], L"tbTinhtrangVL", L"idTinhtrangVL", 9, 3);
				f_Load_SQL_02(NgayAll[i], L"tbChatluongTC", L"idChatluongTC", 10, 4);
				f_Load_SQL_02(NgayAll[i], L"tbTiendoTH", L"idTiendoTH", 11, 5);
				f_Load_SQL_02(NgayAll[i], L"tbCongtacVS", L"idCongtacVS", 12, 6);
				f_Load_SQL_02(NgayAll[i], L"tbAntoanLD", L"idAntoanLD", 13, 7);

				f_Load_SQL_02(NgayAll[i], L"tbNhanxetGSTC", L"", 14, 8);
				f_Load_SQL_02(NgayAll[i], L"tbNhanxetCDT", L"", 15, 9);
				f_Load_SQL_02(NgayAll[i], L"tbYkienNTTC", L"", 16, 10);				

				f_Load_SQL_03(NgayAll[i]);	// Load ghi chú & màu sắc

				f_Load_SQL_04(NgayAll[i]);	// Load đường dẫn hình ảnh

				CString szval = m_listNDNK.GetItemText(gtlit, 4);
				if (f_Check_nghile(gtlit) == 1) m_listNDNK.SetRowColors(gtlit, RED, WHITE);
				else
				{
					// Load ngày chủ nhật
					CString a1 = m_listNDNK.GetItemText(gtlit, 18);
					CString a2 = m_listNDNK.GetItemText(gtlit, 19);
					if (a1 != L"" || a2 != L"") f_Fill_color_row(gtlit, a1, a2);
					else
					{
						if (sThu == L"CN") m_listNDNK.SetRowColors(gtlit, clr1, clr2);
						else
						{
							if (szval == L"") m_listNDNK.SetRowColors(gtlit, WHITE, RGB(192, 192, 192));
						}
					}
				}

				if (itype == 1)
				{
					if (szval != L"") gtlit++;
				}
				else gtlit++;
			}			
		}

		ObjConn.CloseDatabase();
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[x7.69] Lỗi load dữ liệu")+e->m_strError);
	}
	END_CATCH;
}

void Dlg_Quanly_NDNK::f_Fill_color_row(int nRow, CString s1, CString s2)
{
	try
	{
		// Màu nền
		if(s1==L"")
		{
			dwR = 255;dwG = 255;dwB = 255;
		}
		else f_Tack_colorRGB(s1);
		COLORREF mcl1 = RGB(dwR,dwG,dwB);

		// Màu chữ
		if(s2==L"")
		{
			dwR = 0;dwG = 0;dwB = 0;
		}
		else f_Tack_colorRGB(s2);
		COLORREF mcl2 = RGB(dwR,dwG,dwB);

		// Tô màu
		m_listNDNK.SetRowColors(nRow,mcl1,mcl2);
	}
	catch(...){}
}

void Dlg_Quanly_NDNK::f_Load_SQL_ALL()
{
	TRY
	{
		for (int i = 0; i < 15; i++) szGtri[i]=L"";
		ObjConn.OpenAccessDB(pathNK, L"", L"");

		CString sDem=L"";
		CRecordset* Recset = ObjConn.Execute(L"SELECT COUNT(*) AS iDem FROM tbThoitiet WHERE ngayghi LIKE 'ALL';");
		Recset->GetFieldValue(L"iDem",sDem);
		Recset->Close();

		if(_wtoi(sDem)>0)
		{
			// Điều kiện thời tiết & nhiệt độ
			CString szDktt[2]={L"tbThoitiet",L"tbNhietdo"};
			CString sval[5]={L"",L"",L"",L"",L""};
			for (int i = 0; i < 2; i++)
			{
				SqlString.Format(L"SELECT * FROM %s WHERE ngayghi LIKE 'ALL';",szDktt[i]);
				Recset = ObjConn.Execute(SqlString);
				while( !Recset->IsEOF() )
				{
					Recset->GetFieldValue(L"dksang",sval[1]);
					Recset->GetFieldValue(L"dkchieu",sval[2]);
					Recset->GetFieldValue(L"dktoi",sval[3]);
					Recset->GetFieldValue(L"dkdem",sval[4]);				
					szGtri[i].Format(L"%s;%s;%s;%s",sval[1],sval[2],sval[3],sval[4]);
					break;
				}
				Recset->Close();
			}

			// Các dữ liệu khác
			CString szTbl[11]={L"tbBienphapTC",L"tbTinhtrangVL",L"tbChatluongTC",L"tbTiendoTH",L"tbCongtacVS",L"tbAntoanLD",
				L"tbNhanxetGSTC",L"tbNhanxetCDT",L"tbYkienNTTC",L"tbNhanluc",L"tbThietbi"};
		
			CString szFld[6]={L"idBienphapTC",L"idTinhtrangVL",L"idChatluongTC",L"idTiendoTH",L"idCongtacVS",L"idAntoanLD"};

			for (int i = 0; i < 11; i++)
			{
				if(i<6)
				{
					Recset = ObjConn.Execute(L"SELECT * FROM tbCheckbox WHERE ngayghi LIKE 'ALL';");
					while( !Recset->IsEOF() )
					{
						Recset->GetFieldValue(szFld[i],sval[1]);
						szGtri[i+2]+=sval[1];
						szGtri[i+2]+=L"|";
						break;
					}
					Recset->Close();
				}

				SqlString.Format(L"SELECT * FROM %s WHERE ngayghi LIKE 'ALL' ORDER BY maso ASC;",szTbl[i]);
				Recset = ObjConn.Execute(SqlString);
				while( !Recset->IsEOF() )
				{
					Recset->GetFieldValue(L"noidung",sval[1]);
					szGtri[i+2]+=sval[1];
					szGtri[i+2]+=L";";
					Recset->MoveNext();
				}
				Recset->Close();

				if(szGtri[i+2].Right(1)==L";") szGtri[i+2] = szGtri[i+2].Left(szGtri[i+2].GetLength()-1);
			}
		}

		ObjConn.CloseDatabase();
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[tb56] Lỗi load dữ liệu")+e->m_strError);
	}
	END_CATCH;
}

int Dlg_Quanly_NDNK::CheckDulieu(CString szngay)
{
	TRY
	{
		// Kiểm tra dữ liệu có tồn tại không?
		CString sDem = L"";
		SqlString.Format(L"SELECT COUNT(*) AS iDem FROM tbThoitiet WHERE ngayghi LIKE '%s';",szngay);
		CRecordset* Recset = ObjConn.Execute(SqlString);
		Recset->GetFieldValue(L"iDem",sDem);
		Recset->Close();

		if (_wtoi(sDem) > 0) return 1;
		return 0;
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[tb00] Lỗi check dữ liệu") + e->m_strError);
	}
	END_CATCH;
	return 0;
}

void Dlg_Quanly_NDNK::f_Load_SQL_01(CString szngay, CString szTbl, int col, int itypeall)
{
	TRY
	{
		CString sval[5]={L"",L"",L"",L"",L""};
		SqlString.Format(L"SELECT * FROM %s WHERE ngayghi LIKE '%s';",szTbl,szngay);
		CRecordset* Recset = ObjConn.Execute(SqlString);
		while( !Recset->IsEOF() )
		{
			Recset->GetFieldValue(L"dksang",sval[1]);
			Recset->GetFieldValue(L"dkchieu",sval[2]);
			Recset->GetFieldValue(L"dktoi",sval[3]);
			Recset->GetFieldValue(L"dkdem",sval[4]);
			sval[0].Format(L"%s;%s;%s;%s",sval[1],sval[2],sval[3],sval[4]);
			break;
		}
		Recset->Close();

		if(sval[0]!=L"") m_listNDNK.SetItemText(gtlit,col,sval[0]);
		else m_listNDNK.SetItemText(gtlit,col,szGtri[itypeall]);
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[tb01] Lỗi load dữ liệu")+e->m_strError);
	}
	END_CATCH;
}

void Dlg_Quanly_NDNK::f_Load_SQL_02(CString szngay, CString szTbl, CString Fld, int col, int itypeall)
{
	TRY
	{
		CRecordset* Recset;
		CString sval[2]={L"",L""};

		if(Fld!=L"")
		{
			SqlString.Format(L"SELECT * FROM tbCheckbox WHERE ngayghi LIKE '%s';",szngay);
			Recset = ObjConn.Execute(SqlString);
			while( !Recset->IsEOF() )
			{
				Recset->GetFieldValue(Fld,sval[1]);
				sval[0]+=sval[1];
				sval[0]+=L"|";
				break;
			}
			Recset->Close();
		}
		
		SqlString.Format(L"SELECT * FROM %s WHERE ngayghi LIKE '%s' ORDER BY maso ASC;",szTbl,szngay);
		Recset = ObjConn.Execute(SqlString);
		while( !Recset->IsEOF() )
		{
			Recset->GetFieldValue(L"noidung",sval[1]);
			sval[0]+=sval[1];
			sval[0]+=L";";
			Recset->MoveNext();
		}
		Recset->Close();

		if(sval[0]!=L"" && sval[0]!=L"|")
		{
			if(sval[0].Right(1)==L";") sval[0] = sval[0].Left(sval[0].GetLength()-1);
			m_listNDNK.SetItemText(gtlit,col,sval[0]);
		}
		else m_listNDNK.SetItemText(gtlit,col,szGtri[itypeall]);
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[tb02] Lỗi load dữ liệu")+e->m_strError);
	}
	END_CATCH;
}

void Dlg_Quanly_NDNK::f_Load_SQL_03(CString szngay)
{
	TRY
	{
		CString sval[3]={L"",L"",L""};
		SqlString.Format(L"SELECT * FROM tbGhichu WHERE ngayghi LIKE '%s';",szngay);
		CRecordset* Recset = ObjConn.Execute(SqlString);
		while( !Recset->IsEOF() )
		{
			Recset->GetFieldValue(L"ghichu",sval[0]);
			Recset->GetFieldValue(L"fbkgnk",sval[1]);
			Recset->GetFieldValue(L"ftxtnk",sval[2]);
			
			m_listNDNK.SetItemText(gtlit,3,sval[0]);
			m_listNDNK.SetItemText(gtlit,18,sval[1]);
			m_listNDNK.SetItemText(gtlit,19,sval[2]);

			break;
		}
		Recset->Close();

		SqlString.Format(L"SELECT * FROM tbBoqua WHERE ngayghi LIKE '%s';",szngay);
		Recset = ObjConn.Execute(SqlString);
		while( !Recset->IsEOF() )
		{
			Recset->GetFieldValue(L"ghichu",sval[0]);
			if(sval[0]!=L"") m_listNDNK.SetItemText(gtlit,3,sval[0]);
			m_listNDNK.SetItemText(gtlit,20,L"1");
			break;
		}
		Recset->Close();
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[tb03] Lỗi load dữ liệu")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_Quanly_NDNK::f_Load_SQL_04(CString szngay)
{
	TRY
	{
		CString szval = L"", szImage = L"";
		SqlString.Format(L"SELECT * FROM tbHinhAnh WHERE ngayghi LIKE '%s';",szngay);
		CRecordset* Recset = ObjConn.Execute(SqlString);
		while (!Recset->IsEOF())
		{
			Recset->GetFieldValue(L"duongdan", szval);
			szImage += szval;
			szImage += L"|";
			
			Recset->MoveNext();
		}
		Recset->Close();

		if (szImage != L"")
		{
			szImage = szImage.Left(szImage.GetLength() - 1);
			m_listNDNK.SetItemText(gtlit, 17, szImage);
		}
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[tb04] Lỗi load dữ liệu") + e->m_strError);
	}
	END_CATCH;
}


void Dlg_Quanly_NDNK::OnBnClickedB1()
{
	CString s1=L"",s2=L"";
	m_timeDlg[0].GetWindowTextW(s2);

	_iTimeDlg=1;
	AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
	Dlg_Timer _dlg;
	_dlg.DoModal();

	m_timeDlg[0].GetWindowTextW(s1);
	if(s1!=s2)
	{
		sday_bd = s1;
		m_timeDlg[1].GetWindowTextW(s2);
		f_Ktra_dkien();
	}
}


void Dlg_Quanly_NDNK::OnBnClickedB2()
{
	CString s1=L"",s2=L"";
	m_timeDlg[1].GetWindowTextW(s1);

	_iTimeDlg=2;
	AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
	Dlg_Timer _dlg;
	_dlg.DoModal();

	m_timeDlg[1].GetWindowTextW(s2);
	if(s2!=s1)
	{
		sday_kt = s2;
		m_timeDlg[0].GetWindowTextW(s1);
		f_Ktra_dkien();
	}
}


void Dlg_Quanly_NDNK::f_Ktra_dkien()
{
	try
	{
		nItem=-1;
		int nCk = m_chk2.GetCheck();
		if(nCk==1)
		{
			// Check
			CString s1=L"",s2=L"";
			m_timeDlg[0].GetWindowTextW(s1);
			m_timeDlg[1].GetWindowTextW(s2);

			nCk = m_cbbox.GetCurSel();
			if(nCk==0) f_Load_ALL(s1,s2,1);			// Có dữ liệu	
			else if(nCk==1) f_Load_NULL(s1,s2);		// Không dữ liệu
			else f_Load_ALL(s1,s2,0);				// Tất cả
		}
		else
		{
			// Uncheck
			nCk = m_cbbox.GetCurSel();
			if(nCk==0) f_Load_ALL(L"",L"",1);		// Có dữ liệu	
			else if(nCk==1) f_Load_NULL(L"",L"");	// Không dữ liệu
			else f_Load_ALL(L"",L"",0);				// Tất cả
		}
	}
	catch(...){}	
}


void Dlg_Quanly_NDNK::OnBnClickedCk2()
{
	ClickEsc=0;
	int nCk = m_chk2.GetCheck();
	m_timeDlg[0].EnableWindow(nCk);
	m_timeDlg[1].EnableWindow(nCk);
	m_btn1.EnableWindow(nCk);
	m_btn2.EnableWindow(nCk);

	// Kiểm tra điều kiện tìm kiếm
	f_Ktra_dkien();
}


void Dlg_Quanly_NDNK::f_TkiemT1()
{
	CString s1=L"",s2=L"";
	m_timeDlg[0].GetWindowTextW(s1);
	m_timeDlg[1].GetWindowTextW(s2);
	if(s1!=sday_bd)
	{
		sday_bd=s1;
		f_Ktra_dkien();
	}	
}


void Dlg_Quanly_NDNK::f_TkiemT2()
{
	CString s1=L"",s2=L"";
	m_timeDlg[0].GetWindowTextW(s1);
	m_timeDlg[1].GetWindowTextW(s2);
	if(s2!=sday_kt)
	{
		sday_kt=s2;
		f_Ktra_dkien();
	}
}


void Dlg_Quanly_NDNK::OnNMClickL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	ClickEsc=0;
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	nItem = pNMListView->iItem;
	nSubItem = pNMListView->iSubItem;
	*pResult = 0;
}


void Dlg_Quanly_NDNK::OnNh40007()
{
	_cpyNThg=L"";
	POSITION pos=m_listNDNK.GetFirstSelectedItemPosition();
	if(pos!=NULL)
	{
		for (int i = 0; i < m_listNDNK.GetItemCount(); i++)
		if (m_listNDNK.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
		{
			int iktr=0;
			for (int j = 4; j <= 16; j++)
			{
				schep[j] = m_listNDNK.GetItemText(i,j);
				if(schep[j]!=L"") iktr=1;
			}

			if(iktr==1) _cpyNThg = m_listNDNK.GetItemText(i,2);
			else _s(L"Không tồn tại nội dung sao chép!");
			
			return;
		}
	}
}

void Dlg_Quanly_NDNK::OnCbnSelchangeCb1()
{
	// Kiểm tra điều kiện tìm kiếm
	ClickEsc=0;
	f_Ktra_dkien();
}

void Dlg_Quanly_NDNK::OnSysCommand(UINT nID, LPARAM lParam)
{
	if(nID == SC_CLOSE)
	{
		f_Get_width();
		CDialog::OnCancel();
	}
	else CDialog::OnSysCommand(nID, lParam);
}

// Hàm tách màu RGB
void Dlg_Quanly_NDNK::f_Tack_colorRGB(CString sColor)
{
	try
	{
		CString sval = sColor;
		sval.Replace(L"(",L"");
		sval.Replace(L")",L"");
		int N = sval.Find(L",");
		if(N>0)
		{
			dwR = _wtoi(sval.Left(N));
			sColor = sval.Right(sval.GetLength()-N-1);
			N = sColor.Find(L",");
			if(N>0)
			{
				dwG = _wtoi(sColor.Left(N));
				dwB = _wtoi(sColor.Right(sColor.GetLength()-N-1));
			}
		}
	}
	catch(...){}
}

void Dlg_Quanly_NDNK::f_Change_color(int cot)
{
	TRY
	{
		// Thay đổi màu nền & chữ
		POSITION pos=m_listNDNK.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			
			CColorDialog dlg; 
			if (dlg.DoModal() == IDOK)
			{
				CString sValue=L"",sCot=L"";
				COLORREF clr = dlg.GetColor();
				COLORREF m_txtcolor = BLACK, m_bgcolor = WHITE;

				dwR = GetRValue(clr);
				dwG = GetGValue(clr);
				dwB = GetBValue(clr);
				sValue.Format(L"(%d,%d,%d)", dwR, dwG, dwB);

				// cot=18 là cột màu nền, 19 là màu chữ
				if(cot==18)
				{
					sCot = L"fbkgnk";
					m_bgcolor = clr;

					// Nền màu trắng
					if(sValue==L"(255,255,255)") sValue=L"";
				}
				else
				{
					sCot = L"ftxtnk";
					m_txtcolor = clr;

					// Chữ màu đen
					if(sValue==L"(0,0,0)") sValue=L"";
				}

				// Kiểm tra có phải màu ngày lễ không?
				if(m_bgcolor==RED && m_txtcolor==WHITE)
				{
					_s(L"Chữ trắng trên nền đỏ là ký hiệu màu ngày lễ tết, đặc biệt, nghỉ không in... Vui lòng chọn lại màu.");
					return;
				}

				// Lưu vào CSDL
				ObjConn.OpenAccessDB(pathNK, L"", L"");
				CRecordset* Recset;

				CString sngay=L"",sTak=L"",sDem=L"";
				int iCount = m_listNDNK.GetItemCount();
				for (int i = 0; i < iCount; i++)
				{
					if (m_listNDNK.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
					{
						sngay = m_listNDNK.GetItemText(i,2);	// ngày

						// Tách màu
						if(cot==18)
						{
							// cot=18 thì xác định màu chữ cột 19
							sTak = m_listNDNK.GetItemText(i,19);	// chữ
							if(sTak==L"")
							{
								dwR = 0;
								dwG = 0;
								dwB = 0;
							}
							else f_Tack_colorRGB(sTak);

							m_txtcolor = RGB(dwR,dwG,dwB);
						}
						else
						{
							// cot=19 thì xác định màu nền cột 18
							sTak = m_listNDNK.GetItemText(i,18);		// nền
							if(sTak==L"")
							{
								dwR = 255;
								dwG = 255;
								dwB = 255;
							}
							else f_Tack_colorRGB(sTak);

							m_bgcolor = RGB(dwR,dwG,dwB);
						}

						m_listNDNK.SetItemText(i,cot,sValue);
						m_listNDNK.SetRowColors(i,m_bgcolor,m_txtcolor);

						// Kiểm tra đã tồn tại dữ liệu chưa?
						SqlString.Format(L"SELECT COUNT(*) AS iDem FROM tbGhichu WHERE ngayghi LIKE '%s';",sngay);
						Recset = ObjConn.Execute(SqlString);
						Recset->GetFieldValue(L"iDem",sDem);
						Recset->Close();

						if(_wtoi(sDem)>0) SqlString.Format(L"UPDATE tbGhichu SET %s = '%s' WHERE ngayghi LIKE '%s';",sCot,sValue,sngay);
						else
						{
							if(cot==18)	SqlString.Format(L"INSERT INTO tbGhichu (ngayghi,fbkgnk) VALUES ('%s','%s');",sngay,sValue);
							else SqlString.Format(L"INSERT INTO tbGhichu (ngayghi,ftxtnk) VALUES ('%s','%s');",sngay,sValue);
						}

						ObjConn.ExecuteRB(SqlString);
					}
				}

				ObjConn.CloseDatabase();
			}				
		}
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(L"Lỗi cập nhật màu."+e->m_strError);
	}
	END_CATCH;
}


void Dlg_Quanly_NDNK::OnNh40005()
{
	// Đổi nền
	f_Change_color(18);
}


void Dlg_Quanly_NDNK::OnNh40006()
{
	// Đổi màu chữ
	f_Change_color(19);
}


void Dlg_Quanly_NDNK::OnNh40004()
{
	TRY
	{
		// Bỏ qua không in ngày này
		POSITION pos=m_listNDNK.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			// Lưu vào CSDL
			ObjConn.OpenAccessDB(pathNK, L"", L"");
			CRecordset* Recset;

			CString sngay=L"",sDem=L"";
			for (int i = 0; i < (int)m_listNDNK.GetItemCount(); i++)
			{
				if (m_listNDNK.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					// Kiểm tra đã tồn tại dữ liệu chưa?
					sngay = m_listNDNK.GetItemText(i,2);	// ngày
					SqlString.Format(L"SELECT COUNT(*) AS iDem FROM tbBoqua WHERE ngayghi LIKE '%s';",sngay);
					Recset = ObjConn.Execute(SqlString);
					Recset->GetFieldValue(L"iDem",sDem);
					Recset->Close();

					if(_wtoi(sDem)>0)
					{
						SqlString.Format(L"DELETE FROM tbBoqua WHERE ngayghi LIKE '%s';",sngay);
						ObjConn.ExecuteRB(SqlString);

						m_listNDNK.SetRowColors(i,WHITE,BLACK);
						m_listNDNK.SetItemText(i,3,L"");
						m_listNDNK.SetItemText(i,20,L"");
					}
					else
					{
						SqlString.Format(L"INSERT INTO tbBoqua (ngayghi,ghichu) VALUES ('%s','%s');",sngay,L"Nghỉ");
						ObjConn.ExecuteRB(SqlString);

						m_listNDNK.SetRowColors(i,RED,WHITE);
						m_listNDNK.SetItemText(i,3,L"Nghỉ");
						m_listNDNK.SetItemText(i,20,L"1");
					}					
				}
			}

			ObjConn.CloseDatabase();

			// Đọc từ file MDB
			gtriCSV = f_Read_Nhatky();

			// Đọc file dữ liệu csv
			CString spath=L"";
			spath.Format(L"%s\\%s",_fGPathFolder(L"ToolQLCL.dll"),_dFileNle);
			spath.Replace(L"\\\\",L"\\");
	
			gtriCSV += f_Read_file_CSV(spath,gtriCSV);
		}
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(L"Lỗi bỏ qua trang in."+e->m_strError);
	}
	END_CATCH;
}

void Dlg_Quanly_NDNK::OnNh40064()
{
	// Chỉnh sửa chi tiết nội dung
	switch (nSubItem)
	{
		case 3:
		case 8:
		case 9:
		case 10:
		case 11:
		case 12:
		case 13:
		case 14:
		case 15:
		case 16:
		{
			AFX_MANAGE_STATE(AfxGetStaticModuleState());
			Dlg_Nky_Noidung _dlg;
			_dlg.DoModal();
		}
		break;

		case 4:
		case 5:
		{
			AFX_MANAGE_STATE(AfxGetStaticModuleState());
			Dlg_Nky_Thoitiet _dlg;
			_dlg.DoModal();
		}
		break;

		case 6:
		case 7:
		{
			AFX_MANAGE_STATE(AfxGetStaticModuleState());
			Dlg_Nky_BangNLTB _dlg;
			_dlg.DoModal();
		}
		break;

		case 17:
		{
			AFX_MANAGE_STATE(AfxGetStaticModuleState());
			Dlg_Nky_Hinhanh _dlg;
			_dlg.DoModal();
		}
		break;

		default:
			break;
	}
}


void Dlg_Quanly_NDNK::OnEnKillfocusE1()
{
	CString sTruoc=L"", sSau=L"";
	sTruoc = m_listNDNK.GetItemText(CLRow,CLCol);
	txtE_NKTC.GetWindowTextW(sSau);
	
	sSau.Replace(L"'",L"");
	if(sSau==sTruoc) return;

	m_listNDNK.SetWindowText(sSau);
	_CheckUpdate(sSau);
}

void Dlg_Quanly_NDNK::_CheckUpdate(CString sGtri)
{
	try
	{
		CString sFd=L"",sFc=L"",sDem=L"";
		CString sngay = m_listNDNK.GetItemText(CLRow,2);
		ObjConn.OpenAccessDB(pathNK, L"", L"");
		CRecordset* Recset;

		if(CLCol>=3 && CLCol<=16)
		{
			CString sTb[17]={L"",L"",L"",L"tbGhichu",L"tbThoitiet",L"tbNhietdo",L"tbNhanluc",L"tbThietbi",
				L"tbBienphapTC",L"tbTinhtrangVL",L"tbChatluongTC",L"tbTiendoTH",L"tbCongtacVS",L"tbAntoanLD",
				L"tbNhanxetGSTC",L"tbNhanxetCDT",L"tbYkienNTTC"};
			
			CString sChk[17]={L"",L"",L"",L"",L"",L"",L"",L"",
				L"idBienphapTC",L"idTinhtrangVL",L"idChatluongTC",L"idTiendoTH",L"idCongtacVS",L"idAntoanLD",
				L"",L"",L""};

			sFd = sTb[CLCol];
			sFc = sChk[CLCol];

			sDem=L"";
			SqlString.Format(L"SELECT COUNT(*) AS iDem FROM %s WHERE ngayghi LIKE '%s';",sFd,sngay);
			Recset = ObjConn.Execute(SqlString);
			Recset->GetFieldValue(L"iDem",sDem);
			Recset->Close();

			if(_wtoi(sDem)>0)
			{
				SqlString.Format(L"DELETE FROM %s WHERE ngayghi LIKE '%s';",sFd,sngay);
				ObjConn.ExecuteRB(SqlString);
			}

			if(sGtri!=L"")
			{
				if(CLCol==3)
				{
					// Ghi chú, bỏ qua
					int iFor=1;
					CString szTblGC[2]={L"tbGhichu",L"tbBoqua"};
					CString szBq = m_listNDNK.GetItemText(CLRow,20);
					if(szBq==L"1") iFor++;

					for (int i = 0; i < iFor; i++)
					{
						SqlString.Format(L"SELECT COUNT(*) AS iDem FROM %s WHERE ngayghi LIKE '%s';",szTblGC[i],sngay);
						Recset = ObjConn.Execute(SqlString);
						Recset->GetFieldValue(L"iDem",sDem);
						Recset->Close();

						if(_wtoi(sDem)>0)
						{
							SqlString.Format(L"UPDATE %s SET ghichu = '%s' WHERE ngayghi LIKE '%s';",szTblGC[i],sGtri,sngay);
							ObjConn.ExecuteRB(SqlString);
						}
						else
						{
							SqlString.Format(L"INSERT INTO %s (ngayghi,ghichu) VALUES ('%s','%s');",szTblGC[i],sngay,sGtri);
							ObjConn.ExecuteRB(SqlString);
						}
					}					
				}
				else if(CLCol==4 || CLCol==5)
				{
					// Thời tiết, nhiệt độ
					for (int i = 1; i <= 4; i++) pKey[i]=L"";
					_TackTukhoaChuan(sGtri,L";",L";",0,0);
					SqlString.Format(L"INSERT INTO %s (ngayghi,dksang,dkchieu,dktoi,dkdem) "
						L"VALUES ('%s','%s','%s','%s','%s');",sFd,sngay,pKey[1],pKey[2],pKey[3],pKey[4]);
					ObjConn.ExecuteRB(SqlString);
				}
				else if(CLCol==6 || CLCol==7)
				{
					// Nhân lực,thiết bị thi công
					CString szNCMTC = L"tbNhanluc";
					if(CLCol == 7) szNCMTC = L"tbThietbi";
					
					CheckDeleteDulieu(sngay, szNCMTC);
					Themdulieu_02(sngay, szNCMTC, sGtri);					
				}
				else if(CLCol>=8)
				{
					CString szCheckbox=L"",szNoidung=sGtri;
					if(CLCol<=13)
					{
						int pos = sGtri.Find(L"|");
						if(pos>=0)
						{
							szCheckbox = sGtri.Left(pos);
							szNoidung = sGtri.Right(sGtri.GetLength()-pos-1);
						}
						else
						{
							if(sGtri.GetLength()<8 && sGtri.Find(L";")>0) szCheckbox = sGtri;
							else szNoidung = sGtri;
						}

						SqlString.Format(L"SELECT COUNT(*) AS iDem FROM tbCheckbox WHERE ngayghi LIKE '%s';",sngay);
						Recset = ObjConn.Execute(SqlString);
						Recset->GetFieldValue(L"iDem",sDem);
						Recset->Close();

						if(_wtoi(sDem)>0)
						{
							SqlString.Format(L"UPDATE tbCheckbox SET %s = '%s' WHERE ngayghi LIKE '%s';",sFc,szCheckbox,sngay);
							ObjConn.ExecuteRB(SqlString);
						}
						else
						{
							SqlString.Format(L"INSERT INTO tbCheckbox (ngayghi,%s) VALUES ('%s','%s');",sngay,sFc,szCheckbox);
							ObjConn.ExecuteRB(SqlString);
						}
					}

					if(szNoidung!=L"")
					{
						CString szma=L"";
						_TackTukhoaChuan(szNoidung,L";",L";",0,0);
						for (int i = 1; i <= iKey; i++)
						{
							if(i<=9) szma.Format(L"M0%d",i);
							else szma.Format(L"M%d",i);

							SqlString.Format(L"INSERT INTO %s (ngayghi,maso,noidung) "
								L"VALUES ('%s','%s','%s');",sFd,sngay,szma,pKey[i]);
							ObjConn.ExecuteRB(SqlString);
						}
					}
				}							
			}			
		}

		ObjConn.CloseDatabase();		
	}
	catch(...){}
}


void Dlg_Quanly_NDNK::OnNMRClickL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	ClickEsc = 0;
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	nItem = pNMListView->iItem;
	nSubItem = pNMListView->iSubItem;
	*pResult = 0;
}


void Dlg_Quanly_NDNK::PasteDulieu(int itype1, int itype2, int itype3, int itype4, int itype5)
{
	TRY
	{
		if(_cpyNThg==L"") return;

		// Giá trị sao chép
		POSITION pos;

		// Kiểm tra combobox sao chép
		int ncount = m_listNDNK.GetItemCount();
		int nCb = m_cbbox.GetCurSel();
		if(nCb==0 && nItem==-1)
		{
			// Hiện hộp thoại chọn ngày
			iKqua=0;
			if(ncount>0)
			{
				for (int i = 0; i < ncount; i++) XLXD[i+1].CDG[0] = m_listNDNK.GetItemText(i,2);
				iKqua = ncount;
			}

			CString szIns = L"";
			AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
			Dlg_Chonngay_paste _dlg;
			_dlg.DoModal();
			szIns = _dlg.sSaochepngay;
			if (szIns == L"") return;

			// Chèn thêm ngày vào dưới
			CString sThu = fThuMay(szIns);
			m_listNDNK.InsertItem(ncount, L"", 0);
			m_listNDNK.SetItemText(ncount,1,sThu);
			m_listNDNK.SetItemText(ncount,2, szIns);

			pos=m_listNDNK.GetFirstSelectedItemPosition();
			if(pos==NULL && ncount>0) m_listNDNK.SetItemState(ncount,LVIS_SELECTED,LVIS_SELECTED);

			ncount++;
		}

////////////////////////////////////////////////////////////////////////

		// Load dữ liệu và truy vấn
		ObjConn.OpenAccessDB(pathNK, L"", L"");
		
		// Dán nội dung
		pos=m_listNDNK.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			// Xác định giá trị hình ảnh
			slImage = 0, slNhanluc = 0, slThietbi = 0;
			if (itype2 == 1) GetNhancong(_cpyNThg);
			if (itype3 == 1) GetThietbi(_cpyNThg);
			if (itype5 == 1) GetHinhAnh(_cpyNThg);

			CString sCNgay=L"";
			for (int a = 0; a < ncount; a++)
			{
				sCNgay = m_listNDNK.GetItemText(a,2);
				if (sCNgay !=_cpyNThg && m_listNDNK.GetItemState(a, LVIS_SELECTED) == LVIS_SELECTED)
				{
					// code sao chép
					for (int j = 4; j <= 16; j++)
					{
						m_listNDNK.SetItemText(a, j, schep[j]);
					}

					// Paste vào CSDL
					if (itype1 == 1) PasteThoitietNhietdo(sCNgay);
					if (itype2 == 1) PasteNhancong(sCNgay);
					if (itype3 == 1) PasteMaythicong(sCNgay);
					if (itype4 == 1) PasteGhichuNhanxet(sCNgay);
					if (itype5 == 1) PasteHinhAnh(sCNgay);
				}
			}
		}

		ObjConn.CloseDatabase();
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[ql46] Lỗi dán dữ liệu")+e->m_strError);
	}
	END_CATCH;
}

void Dlg_Quanly_NDNK::GetNhancong(CString szNgay)
{
	TRY
	{
		SqlString.Format(L"SELECT * FROM tbBangNhanluc WHERE ngayghi LIKE '%s';", szNgay);
		CRecordset* Recset = ObjConn.Execute(SqlString);
		while (!Recset->IsEOF())
		{
			slNhanluc++;
			Recset->GetFieldValue(L"maso", MSNhanluc[slNhanluc].maso);
			Recset->GetFieldValue(L"cot1", MSNhanluc[slNhanluc].cot[1]);
			Recset->GetFieldValue(L"cot2", MSNhanluc[slNhanluc].cot[2]);
			Recset->GetFieldValue(L"cot3", MSNhanluc[slNhanluc].cot[3]);
			Recset->GetFieldValue(L"cot4", MSNhanluc[slNhanluc].cot[4]);
			Recset->MoveNext();
		}
		Recset->Close();
	}
		CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[ql319] Lỗi xác định hình ảnh") + e->m_strError);
	}
	END_CATCH;
}

void Dlg_Quanly_NDNK::GetThietbi(CString szNgay)
{
	TRY
	{
		SqlString.Format(L"SELECT * FROM tbBangThietbi WHERE ngayghi LIKE '%s';", szNgay);
		CRecordset* Recset = ObjConn.Execute(SqlString);
		while (!Recset->IsEOF())
		{
			slThietbi++;
			Recset->GetFieldValue(L"maso", MSThietbi[slThietbi].maso);
			Recset->GetFieldValue(L"cot1", MSThietbi[slThietbi].cot[1]);
			Recset->GetFieldValue(L"cot2", MSThietbi[slThietbi].cot[2]);
			Recset->GetFieldValue(L"cot3", MSThietbi[slThietbi].cot[3]);
			Recset->MoveNext();
		}
		Recset->Close();
	}
		CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[ql319] Lỗi xác định bảng nhân lực - thiết bị") + e->m_strError);
	}
	END_CATCH;
}

void Dlg_Quanly_NDNK::GetHinhAnh(CString szNgay)
{
	TRY
	{
		SqlString.Format(L"SELECT * FROM tbHinhAnh WHERE ngayghi LIKE '%s';", szNgay);
		CRecordset* Recset = ObjConn.Execute(SqlString);
		while (!Recset->IsEOF())
		{
			slImage++;
			Recset->GetFieldValue(L"maso", MSI[slImage].maso);
			Recset->GetFieldValue(L"duongdan", MSI[slImage].duongdan);
			Recset->GetFieldValue(L"vitri", MSI[slImage].vitri);
			Recset->GetFieldValue(L"kichthuoc", MSI[slImage].kichthuoc);
			Recset->GetFieldValue(L"ghichu", MSI[slImage].ghichu);
			Recset->MoveNext();
		}
		Recset->Close();
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[ql319] Lỗi xác định hình ảnh") + e->m_strError);
	}
	END_CATCH;
}

// Dán tất cả dữ liệu
void Dlg_Quanly_NDNK::OnD40053()
{
	PasteDulieu(1,1,1,1,0);		// Không paste hình ảnh
}

// Thời tiết, nhiệt độ
void Dlg_Quanly_NDNK::OnD40054()
{
	PasteDulieu(1,0,0,0,0);
}

// Nhân công
void Dlg_Quanly_NDNK::OnD40055()
{
	PasteDulieu(0,1,0,0,0);
}

// Máy
void Dlg_Quanly_NDNK::OnD40057()
{
	PasteDulieu(0,0,1,0,0);
}

// Ghi chú, nhận xét
void Dlg_Quanly_NDNK::OnD40056()
{
	PasteDulieu(0,0,0,1,0);
}

// Hình ảnh
void Dlg_Quanly_NDNK::OnD40059()
{
	PasteDulieu(0,0,0,0,1);
}


void Dlg_Quanly_NDNK::OnEnSetfocusE1()
{
	ClickEsc=1;
}


void Dlg_Quanly_NDNK::OnEnSetfocusT1()
{
	ClickEsc=0;
}


void Dlg_Quanly_NDNK::OnEnSetfocusT2()
{
	ClickEsc=0;
}

void Dlg_Quanly_NDNK::OnX40012()
{
	TRY
	{
		if(m_cbbox.GetCurSel()==1) return;

		POSITION pos=m_listNDNK.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			if(_yn(L"Bạn có chắc chắn xóa dữ liệu được chọn?",0)!=6) return;

			CString szval=L"",sDem=L"";
			CString sTbDk[2]={L"tbThoitiet",L"tbNhietdo"};
			CString sTb[17]={L"tbAntoanLD",L"tbBienphapTC",L"tbBoqua",L"tbCongtacVS",L"tbChatluongTC",L"tbGhichu",
				L"tbHinhAnh",L"tbNhanluc",L"tbNhanxetGSTC",L"tbNhanxetCDT",L"tbTiendoTH",L"tbTinhtrangVL",
				L"tbThietbi",L"tbYkienNTTC",L"tbCheckbox",L"tbBangNhanluc",L"tbBangThietbi"};

			// Load dữ liệu và truy vấn
			ObjConn.OpenAccessDB(pathNK, L"", L"");
			CRecordset* Recset;			
			
			for (int a = m_listNDNK.GetItemCount(); a >=0 ; a--)
			{
				if (m_listNDNK.GetItemState(a, LVIS_SELECTED) == LVIS_SELECTED)
				{
					// Xóa nội dung đã tồn tại
					szval = m_listNDNK.GetItemText(a,2);

					for (int j = 0; j < 2; j++)
					{
						SqlString.Format(L"SELECT COUNT(*) AS iDem FROM %s WHERE ngayghi LIKE '%s';",sTbDk[j],szval);
						Recset = ObjConn.Execute(SqlString);
						Recset->GetFieldValue(L"iDem",sDem);
						Recset->Close();

						if(_wtoi(sDem)>0)
						{
							SqlString.Format(L"UPDATE %s SET dksang='',dkchieu='',dktoi='',dkdem='' "
								L"WHERE ngayghi LIKE '%s';",sTbDk[j],szval);
							ObjConn.ExecuteRB(SqlString);
						}						
					}

					for (int j = 0; j < 17; j++)
					{
						SqlString.Format(L"DELETE FROM %s WHERE ngayghi LIKE '%s';",sTb[j],szval);
						ObjConn.ExecuteRB(SqlString);
					}

					// Xóa tất cả list control
					for (int j = 3; j <= 20; j++) m_listNDNK.SetItemText(a,j,L"");
					
					szval = m_listNDNK.GetItemText(a,1);
					if(szval==L"CN") m_listNDNK.SetRowColors(a,clr1,clr2);
					else m_listNDNK.SetRowColors(a,WHITE,RGB(192,192,192));
				}
			}

			ObjConn.CloseDatabase();
		}
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[ql66] Lỗi xóa dữ liệu")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_Quanly_NDNK::OnX40014()
{
	TRY
	{
		if(m_cbbox.GetCurSel()==1) return;

		POSITION pos=m_listNDNK.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			CString szval=L"",sDem=L"";
			CString sTbDk[2]={L"tbThoitiet",L"tbNhietdo"};

			// Load dữ liệu và truy vấn
			ObjConn.OpenAccessDB(pathNK, L"", L"");
			CRecordset* Recset;
			
			for (int a = m_listNDNK.GetItemCount(); a >=0 ; a--)
			{
				if (m_listNDNK.GetItemState(a, LVIS_SELECTED) == LVIS_SELECTED)
				{
					// Xóa nội dung đã tồn tại
					szval = m_listNDNK.GetItemText(a,2);
					for (int j = 0; j < 2; j++)
					{
						SqlString.Format(L"SELECT COUNT(*) AS iDem FROM %s WHERE ngayghi LIKE '%s';",sTbDk[j],szval);
						Recset = ObjConn.Execute(SqlString);
						Recset->GetFieldValue(L"iDem",sDem);
						Recset->Close();

						if(_wtoi(sDem)>0)
						{
							SqlString.Format(L"UPDATE %s SET dksang='',dkchieu='',dktoi='',dkdem='' "
								L"WHERE ngayghi LIKE '%s';",sTbDk[j],szval);
							ObjConn.ExecuteRB(SqlString);
						}						
					}

					// Xóa trên list control
					m_listNDNK.SetItemText(a,4,L"");
					m_listNDNK.SetItemText(a,5,L"");
				}
			}

			ObjConn.CloseDatabase();
		}
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[ql67] Lỗi xóa dữ liệu")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_Quanly_NDNK::OnX40016()
{
	TRY
	{
		if(m_cbbox.GetCurSel()==1) return;

		POSITION pos=m_listNDNK.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			CString sTb[2]={L"tbNhanluc",L"tbBangNhanluc"};

			// Load dữ liệu và truy vấn
			ObjConn.OpenAccessDB(pathNK, L"", L"");
			CString szval=L"";
			
			for (int a = m_listNDNK.GetItemCount(); a >=0 ; a--)
			{
				if (m_listNDNK.GetItemState(a, LVIS_SELECTED) == LVIS_SELECTED)
				{
					// Xóa nội dung đã tồn tại
					szval = m_listNDNK.GetItemText(a,2);
					for (int j = 0; j < 2; j++)
					{
						SqlString.Format(L"DELETE FROM %s WHERE ngayghi LIKE '%s';",sTb[j],szval);
						ObjConn.ExecuteRB(SqlString);
					}

					// Xóa trên list control
					m_listNDNK.SetItemText(a,6,L"");
				}
			}

			ObjConn.CloseDatabase();
		}
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[ql68] Lỗi xóa dữ liệu")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_Quanly_NDNK::OnX40017()
{
	TRY
	{
		if(m_cbbox.GetCurSel()==1) return;

		POSITION pos=m_listNDNK.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			CString sTb[2]={L"tbThietbi",L"tbBangThietbi"};

			// Load dữ liệu và truy vấn
			ObjConn.OpenAccessDB(pathNK, L"", L"");
			CString szval=L"";
			
			for (int a = m_listNDNK.GetItemCount(); a >=0 ; a--)
			{
				if (m_listNDNK.GetItemState(a, LVIS_SELECTED) == LVIS_SELECTED)
				{
					// Xóa nội dung đã tồn tại
					szval = m_listNDNK.GetItemText(a,2);
					for (int j = 0; j < 2; j++)
					{
						SqlString.Format(L"DELETE FROM %s WHERE ngayghi LIKE '%s';",sTb[j],szval);
						ObjConn.ExecuteRB(SqlString);
					}

					// Xóa trên list control
					m_listNDNK.SetItemText(a,7,L"");
				}
			}

			ObjConn.CloseDatabase();
		}
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[ql69] Lỗi xóa dữ liệu")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_Quanly_NDNK::OnX40018()
{
	TRY
	{
		if(m_cbbox.GetCurSel()==1) return;

		POSITION pos=m_listNDNK.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			CString sTb[10]={L"tbAntoanLD",L"tbBienphapTC",L"tbCongtacVS",L"tbChatluongTC",
				L"tbNhanxetGSTC",L"tbNhanxetCDT",L"tbTiendoTH",L"tbTinhtrangVL",L"tbYkienNTTC",L"tbCheckbox"};

			// Load dữ liệu và truy vấn
			ObjConn.OpenAccessDB(pathNK, L"", L"");
			CString szval=L"";
			
			for (int a = m_listNDNK.GetItemCount(); a >=0 ; a--)
			{
				if (m_listNDNK.GetItemState(a, LVIS_SELECTED) == LVIS_SELECTED)
				{
					// Xóa nội dung đã tồn tại
					szval = m_listNDNK.GetItemText(a,2);
					for (int j = 0; j < 10; j++)
					{
						SqlString.Format(L"DELETE FROM %s WHERE ngayghi LIKE '%s';",sTb[j],szval);
						ObjConn.ExecuteRB(SqlString);
					}

					// Xóa trên list control
					for (int j = 8; j <= 16; j++) m_listNDNK.SetItemText(a,j,L"");
				}
			}

			ObjConn.CloseDatabase();
		}		
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[ql70] Lỗi xóa dữ liệu")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_Quanly_NDNK::OnX40058()
{
	TRY
	{
		if(m_cbbox.GetCurSel()==1) return;

		POSITION pos=m_listNDNK.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			// Load dữ liệu và truy vấn
			ObjConn.OpenAccessDB(pathNK, L"", L"");
			CString szval=L"";
			
			for (int a = m_listNDNK.GetItemCount(); a >=0 ; a--)
			{
				if (m_listNDNK.GetItemState(a, LVIS_SELECTED) == LVIS_SELECTED)
				{
					// Xóa nội dung đã tồn tại
					szval = m_listNDNK.GetItemText(a,2);
					SqlString.Format(L"DELETE FROM tbHinhAnh WHERE ngayghi LIKE '%s';",szval);
					ObjConn.ExecuteRB(SqlString);

					// Xóa trên list control
					m_listNDNK.SetItemText(a,17,L"");
				}
			}

			ObjConn.CloseDatabase();
		}
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[ql93] Lỗi xóa dữ liệu")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_Quanly_NDNK::OnX40062()
{
	TRY
	{
		if(_yn(L"Bạn có chắc chắn xóa các dữ liệu chung trong nhật ký?",0)!=6) return;

		ObjConn.OpenAccessDB(pathNK, L"", L"");
		
		CString szTbl[19]={L"tbAntoanLD",L"tbBangNhanluc",L"tbBangThietbi",L"tbBienphapTC",
			L"tbBoqua",L"tbCongtacVS",L"tbChatluongTC",L"tbCheckbox",L"tbGhichu",L"tbHinhAnh",
			L"tbNhanluc",L"tbNhanxetGSTC",L"tbNhanxetCDT",L"tbNhietdo",L"tbTiendoTH",
			L"tbTinhtrangVL",L"tbThietbi",L"tbThoitiet",L"tbYkienNTTC"};

		for (int i = 0; i < 19; i++)
		{
			SqlString.Format(L"DELETE FROM %s WHERE ngayghi LIKE 'ALL';",szTbl[i]);
			ObjConn.ExecuteRB(SqlString);
		}

		ObjConn.CloseDatabase();
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[ql80] Lỗi xóa dữ liệu")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_Quanly_NDNK::OnX40063()
{
	TRY
	{
		if(_yn(L"Bạn có chắc chắn xóa toàn bộ dữ liệu nhật ký?",0)!=6) return;

		ObjConn.OpenAccessDB(pathNK, L"", L"");
		
		CString szTbl[19]={L"tbAntoanLD",L"tbBangNhanluc",L"tbBangThietbi",L"tbBienphapTC",
			L"tbBoqua",L"tbCongtacVS",L"tbChatluongTC",L"tbCheckbox",L"tbGhichu",L"tbHinhAnh",
			L"tbNhanluc",L"tbNhanxetGSTC",L"tbNhanxetCDT",L"tbNhietdo",L"tbTiendoTH",
			L"tbTinhtrangVL",L"tbThietbi",L"tbThoitiet",L"tbYkienNTTC"};

		for (int i = 0; i < 19; i++)
		{
			SqlString.Format(L"DELETE FROM %s;",szTbl[i]);
			ObjConn.ExecuteRB(SqlString);
		}

		ObjConn.CloseDatabase();
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[ql79] Lỗi xóa dữ liệu")+e->m_strError);
	}
	END_CATCH;
}

int Dlg_Quanly_NDNK::CheckDeleteDulieu(CString szNgay, CString szTbl)
{
	TRY
	{
		CString sDem = L"";
		SqlString.Format(L"SELECT COUNT(*) AS iDem FROM %s WHERE ngayghi LIKE '%s';",szTbl, szNgay);
		CRecordset* Recset = ObjConn.Execute(SqlString);
		Recset->GetFieldValue(L"iDem",sDem);
		Recset->Close();

		if (_wtoi(sDem) > 0)
		{
			SqlString.Format(L"DELETE FROM %s WHERE ngayghi LIKE '%s';", szTbl, szNgay);
			ObjConn.ExecuteRB(SqlString);
			return 1;
		}
		return 0;
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[ql262] Lỗi xóa dữ liệu sao chép") + e->m_strError);
	}
	END_CATCH;
	return 0;
}

void Dlg_Quanly_NDNK::Themdulieu_01(CString szNgay, CString szTbl, CString szGtri)
{
	TRY
	{
		if (szGtri != L"")
		{
			for (int i = 1; i <= 4; i++) pKey[i] = L"";
			_TackTukhoaChuan(szGtri, L";", L";", 0, 0);
			SqlString.Format(L"INSERT INTO %s (ngayghi,dksang,dkchieu,dktoi,dkdem) "
				L"VALUES ('%s','%s','%s','%s','%s');", szTbl, szNgay, pKey[1], pKey[2], pKey[3], pKey[4]);
			ObjConn.ExecuteRB(SqlString);
		}		
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[ql263] Lỗi thêm dữ liệu (1)") + e->m_strError);
	}
	END_CATCH;
}

void Dlg_Quanly_NDNK::Themdulieu_02(CString szNgay, CString szTbl, CString szGtri)
{
	TRY
	{
		if (szGtri != L"")
		{
			_TackTukhoaChuan(szGtri, L";", L";", 0, 0);
			CString szMaso = L"";
			for (int i = 1; i <= iKey; i++)
			{
				szMaso.Format(L"M%02d", i);
				SqlString.Format(L"INSERT INTO %s (ngayghi,maso,noidung) "
					L"VALUES ('%s','%s','%s');", szTbl, szNgay, szMaso, pKey[i]);
				ObjConn.ExecuteRB(SqlString);
			}
		}
	}
		CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[ql263] Lỗi thêm dữ liệu (2)") + e->m_strError);
	}
	END_CATCH;
}

void Dlg_Quanly_NDNK::Themdulieu_03(CString szNgay)
{
	TRY
	{
		int pos = -1;
		CString szgtr[6] = { L"",L"",L"",L"",L"",L"" };
		for (int i = 0; i < 6; i++)
		{
			pos = schep[i + 8].Find(L"|");
			if (pos > 0)
			{
				szgtr[i] = schep[i + 8].Left(pos);
				szgtr[i].Trim();
			}
		}

		SqlString.Format(L"INSERT INTO tbCheckbox (ngayghi,idBienphapTC,idTinhtrangVL,idChatluongTC,"
			L"idTiendoTH,idCongtacVS,idAntoanLD) VALUES ('%s','%s','%s','%s','%s','%s','%s');", 
			szNgay, szgtr[0], szgtr[1], szgtr[2], szgtr[3], szgtr[4], szgtr[5]);
		ObjConn.ExecuteRB(SqlString);
	}
		CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[ql263] Lỗi thêm dữ liệu (3)") + e->m_strError);
	}
	END_CATCH;
}

void Dlg_Quanly_NDNK::PasteThoitietNhietdo(CString szNgay)
{
	CheckDeleteDulieu(szNgay, L"tbThoitiet");
	CheckDeleteDulieu(szNgay, L"tbNhietdo");

	Themdulieu_01(szNgay, L"tbThoitiet", schep[4]);
	Themdulieu_01(szNgay, L"tbNhietdo", schep[5]);
}

void Dlg_Quanly_NDNK::PasteNhancong(CString szNgay)
{
	CheckDeleteDulieu(szNgay, L"tbNhanluc");
	Themdulieu_02(szNgay, L"tbNhanluc", schep[6]);
	PasteBangNhanluc(szNgay);
}

void Dlg_Quanly_NDNK::PasteMaythicong(CString szNgay)
{
	CheckDeleteDulieu(szNgay, L"tbThietbi");
	Themdulieu_02(szNgay, L"tbThietbi", schep[7]);
	PasteBangThietbi(szNgay);
}

void Dlg_Quanly_NDNK::PasteGhichuNhanxet(CString szNgay)
{
	TRY
	{
		int pos = -1;
		CString szval = L"";
		CString sTb[9] = {L"tbBienphapTC",L"tbTinhtrangVL",L"tbChatluongTC",L"tbTiendoTH",
			L"tbCongtacVS",	L"tbAntoanLD",L"tbNhanxetGSTC",L"tbNhanxetCDT",L"tbYkienNTTC"};

		for (int i = 0; i < 9; i++)
		{
			CheckDeleteDulieu(szNgay, sTb[i]);

			szval = schep[i + 8];
			pos = szval.Find(L"|");
			if (pos > 0)
			{
				szval = szval.Right(szval.GetLength() - pos - 1);
				szval.Trim();
			}

			Themdulieu_02(szNgay, sTb[i], szval);
		}

		Themdulieu_03(szNgay);
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[ql261] Lỗi paste ghi chú, nhận xét") + e->m_strError);
	}
	END_CATCH;
}


void Dlg_Quanly_NDNK::PasteHinhAnh(CString szNgay)
{
	TRY
	{
		if (slImage > 0)
		{
			for (int i = 1; i <= slImage; i++)
			{
				SqlString.Format(L"INSERT INTO tbHinhAnh (ngayghi,maso,duongdan,vitri,kichthuoc,ghichu) "
					L"VALUES ('%s','%s','%s','%s','%s','%s');", szNgay, 
					MSI[i].maso, MSI[i].duongdan, MSI[i].vitri, MSI[i].kichthuoc, MSI[i].ghichu);
				ObjConn.ExecuteRB(SqlString);
			}
		}
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[ql261] Lỗi paste hình ảnh") + e->m_strError);
	}
	END_CATCH;
}

void Dlg_Quanly_NDNK::PasteBangNhanluc(CString szNgay)
{
	TRY
	{
		if (slNhanluc > 0)
		{
			for (int i = 1; i <= slNhanluc; i++)
			{
				SqlString.Format(L"INSERT INTO tbBangNhanluc (ngayghi,maso,cot1,cot2,cot3,cot4) "
					L"VALUES ('%s','%s','%s','%s','%s','%s');", szNgay,
					MSNhanluc[i].maso, MSNhanluc[i].cot[1], MSNhanluc[i].cot[2],
					MSNhanluc[i].cot[3], MSNhanluc[i].cot[4]);
				ObjConn.ExecuteRB(SqlString);
			}
		}
	}
		CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[ql261] Lỗi paste bảng nhân lực") + e->m_strError);
	}
	END_CATCH;
}

void Dlg_Quanly_NDNK::PasteBangThietbi(CString szNgay)
{
	TRY
	{
		if (slThietbi > 0)
		{
			for (int i = 1; i <= slThietbi; i++)
			{
				SqlString.Format(L"INSERT INTO tbBangThietbi (ngayghi,maso,cot1,cot2,cot3) "
					L"VALUES ('%s','%s','%s','%s','%s');", szNgay,
					MSThietbi[i].maso, MSThietbi[i].cot[1],
					MSThietbi[i].cot[2], MSThietbi[i].cot[3]);
				ObjConn.ExecuteRB(SqlString);
			}
		}
	}
		CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[ql261] Lỗi paste bảng thiết bị") + e->m_strError);
	}
	END_CATCH;
}


void Dlg_Quanly_NDNK::OnHdnEndtrackL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMHEADER pNMListHeader = reinterpret_cast<LPNMHEADER>(pNMHDR);
	int pSubItem = pNMListHeader->iItem;

	switch (pSubItem)
	{
		case 0:
		case 18:
		case 19:
		case 20:
			pNMListHeader->pitem->cxy = 0;
			break;

		case 1:
			pNMListHeader->pitem->cxy = 40;
			break;

		case 2:
			pNMListHeader->pitem->cxy = 100;
			break;		

		default:
			break;
	}

	*pResult = 0;
}

void Dlg_Quanly_NDNK::ShowColumnWidth(int nColumn, int nWidth, int itype)
{
	LVCOLUMN col;
	col.mask = LVCF_WIDTH;
	col.cx = nWidth;

	if (itype == 0)
	{
		int icheck = m_listNDNK.GetColumnWidth(nColumn);
		if (icheck > 0) col.cx = 0;
	}

	m_listNDNK.SetColumn(nColumn, &col);
}


void Dlg_Quanly_NDNK::OnDwShowAll()
{
	int col[20] = { 0,40,100,100,150,150,150,150,150,150,150,150,150,150,150,150,150,150,0,0 };
	for (int i = 0; i < 20; i++) ShowColumnWidth(i, col[i], 1);
}

void Dlg_Quanly_NDNK::OnDwShowGhichu()
{
	ShowColumnWidth(3, 100, 0);
}

void Dlg_Quanly_NDNK::OnDwShowThoitiet()
{
	ShowColumnWidth(4, 150, 0);
	ShowColumnWidth(5, 150, 0);
}

void Dlg_Quanly_NDNK::OnDwShowNhancong()
{
	ShowColumnWidth(6, 150, 0);
}

void Dlg_Quanly_NDNK::OnDwShowMaythicong()
{
	ShowColumnWidth(7, 150, 0);
}

void Dlg_Quanly_NDNK::OnDwShowChitieu()
{
	for (int i = 8; i <= 11; i++) ShowColumnWidth(i, 150, 0);
}

void Dlg_Quanly_NDNK::OnDwShowVesinhATLD()
{
	ShowColumnWidth(12, 150, 0);
	ShowColumnWidth(13, 150, 0);
}

void Dlg_Quanly_NDNK::OnDwShowNhanxet()
{
	ShowColumnWidth(14, 150, 0);
	ShowColumnWidth(15, 150, 0);
	ShowColumnWidth(16, 150, 0);
}

void Dlg_Quanly_NDNK::OnDwShowPathHinhanh()
{
	ShowColumnWidth(17, 150, 0);
}

LRESULT Dlg_Quanly_NDNK::OnHeaderRightClick(WPARAM wParam, LPARAM lParam)
{
	try
	{
		nItem = -1;
		nSubItem = -1;

		ASSERT((HWND)wParam == m_list.GetSafeHwnd());

		CListCtrlEx::CHeaderRightClick *hit = (CListCtrlEx::CHeaderRightClick*) lParam;

		CMenu m_Menu;
		m_Menu.CreateMenu();

		CMenu subMenu;
		subMenu.CreatePopupMenu();
		subMenu.AppendMenuW(MF_STRING, IDC_COLUMN_POPUP, L"Hiển thị mặc định");
		subMenu.AppendMenuW(MF_SEPARATOR);

		subMenu.AppendMenuW(MF_STRING, IDC_COLUMN_GHICHU, L"Ghi chú");
		subMenu.AppendMenuW(MF_STRING, IDC_COLUMN_THOITIET, L"Thời tiết/ nhiệt độ");
		subMenu.AppendMenuW(MF_STRING, IDC_COLUMN_NHANCONG, L"Nhân công");
		subMenu.AppendMenuW(MF_STRING, IDC_COLUMN_MAYTHICONG, L"Máy móc thiết bị");
		subMenu.AppendMenuW(MF_STRING, IDC_COLUMN_CHITIEU, L"Các chỉ tiêu đánh giá");
		subMenu.AppendMenuW(MF_STRING, IDC_COLUMN_VSATLD, L"Vệ sinh/ ATLĐ");
		subMenu.AppendMenuW(MF_STRING, IDC_COLUMN_NHANXET, L"Các nhận xét/ đánh giá");
		subMenu.AppendMenuW(MF_STRING, IDC_COLUMN_HINHANH, L"Đường dẫn hình ảnh");
		
		int ncol[20];
		for (int i = 0; i < 20; i++) ncol[i] = (int)m_listNDNK.GetColumnWidth(i);
		
		if (ncol[3] > 0) subMenu.CheckMenuItem(IDC_COLUMN_GHICHU, MF_CHECKED);
		if (ncol[4] + ncol[5] > 0) subMenu.CheckMenuItem(IDC_COLUMN_THOITIET, MF_CHECKED);
		if (ncol[6] > 0) subMenu.CheckMenuItem(IDC_COLUMN_NHANCONG, MF_CHECKED);
		if (ncol[7] > 0) subMenu.CheckMenuItem(IDC_COLUMN_MAYTHICONG, MF_CHECKED);
		if (ncol[8] + ncol[9] + ncol[10] + ncol[11] > 0) subMenu.CheckMenuItem(IDC_COLUMN_CHITIEU, MF_CHECKED);
		if (ncol[12] + ncol[13] > 0) subMenu.CheckMenuItem(IDC_COLUMN_VSATLD, MF_CHECKED);
		if (ncol[14] + ncol[15] + ncol[16] > 0) subMenu.CheckMenuItem(IDC_COLUMN_NHANXET, MF_CHECKED);
		if (ncol[17] > 0) subMenu.CheckMenuItem(IDC_COLUMN_HINHANH, MF_CHECKED);

		m_Menu.AppendMenuW(MF_POPUP, reinterpret_cast<UINT_PTR>(subMenu.m_hMenu), L"");

		CMenu* popup = m_Menu.GetSubMenu(0);
		popup->TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON,
			hit->m_mousePos.x, hit->m_mousePos.y, this);

		return 0;
	}
	catch (...) {}

	return 0;
}
