#include "Dlg_Open_DSVL.h"
#include "Dlg_all_tieuchuan.h"
#include "Dlg_trauu_tieuchuan.h"
#include "Dlg_open_progress.h"

Dlg_Open_DSVL::Dlg_Open_DSVL(CWnd* pParent /*=NULL*/)
	: CDialog(Dlg_Open_DSVL::IDD, pParent)
{
	
}

void Dlg_Open_DSVL::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, DSVL_T1, m_Txt);
	DDX_Control(pDX, DSVL_L1, lit_dmvl);
	DDX_Control(pDX, DSVL_L2, lit_tcvl);
	DDX_Control(pDX, DSVL_L3, lit_ppvl);
	DDX_Control(pDX, DSVL_ED1, txtE_DMVL);
	DDX_Control(pDX, DSVL_ED2, txtE_TCVL);
	DDX_Control(pDX, DSVL_ED3, txtE_PPVL);
	DDX_Control(pDX, DSVL_S111, m_Group2);
	DDX_Control(pDX, DSVL_CBS1, m_cbk2);
	DDX_Control(pDX, DSVL_CBS2, m_cbk1);
	DDX_Control(pDX, DSVL_B3, bxx3);
	DDX_Control(pDX, DSVL_CBS3, m_chkColor);
	DDX_Control(pDX, DSVL_B12, m_btnColor);
	DDX_Control(pDX, DSVL_S10, m_txtColor);	
	DDX_Control(pDX, DSVL_CBS5, m_chkLMau);
	DDX_Control(pDX, DSVL_CBS4, m_chkBTC);
	DDX_Control(pDX, DSVL_CBB1, m_cbbox);

	DDX_Control(pDX, DSVL_CKLUU, m_chkLuuTC);
	DDX_Control(pDX, DSVL_T2, m_txtTChon);

	DDX_Control(pDX, DSVL_PRE, m_btnPre);
	DDX_Control(pDX, DSVL_NXT, m_btnNxt);
	DDX_Control(pDX, DSVL_CBBFL, m_cbbFile);	
}

BEGIN_MESSAGE_MAP(Dlg_Open_DSVL, CDialog)
	ON_WM_SIZE()
	ON_WM_SIZING()
	ON_WM_SYSCOMMAND()
	ON_WM_CONTEXTMENU()
	ON_BN_CLICKED(DSVL_B1, &Dlg_Open_DSVL::OnBnClickedB1)
	ON_BN_CLICKED(DSVL_B2, &Dlg_Open_DSVL::OnBnClickedB2)
	ON_BN_CLICKED(DSVL_B4, &Dlg_Open_DSVL::OnBnClickedB4)
	ON_BN_CLICKED(DSVL_B5, &Dlg_Open_DSVL::OnBnClickedB5)
	ON_BN_CLICKED(DSVL_B3, &Dlg_Open_DSVL::OnBnClickedB3)
	ON_BN_CLICKED(DSVL_B6, &Dlg_Open_DSVL::OnBnClickedB6)
	ON_EN_KILLFOCUS(DSVL_ED1, &Dlg_Open_DSVL::OnEnKillfocusEd1)
	ON_NOTIFY(NM_CLICK, DSVL_L1, &Dlg_Open_DSVL::OnNMClickL1)
	ON_BN_CLICKED(DSVL_B7, &Dlg_Open_DSVL::OnBnClickedB7)
	ON_BN_CLICKED(DSVL_B8, &Dlg_Open_DSVL::OnBnClickedB8)
	ON_BN_CLICKED(DSVL_B9, &Dlg_Open_DSVL::OnBnClickedB9)
	ON_BN_CLICKED(DSVL_B10, &Dlg_Open_DSVL::OnBnClickedB10)
	ON_BN_CLICKED(DSVL_B11, &Dlg_Open_DSVL::OnBnClickedB11)
	ON_NOTIFY(LVN_KEYDOWN, DSVL_L1, &Dlg_Open_DSVL::OnLvnKeydownL1)
	ON_NOTIFY(LVN_KEYDOWN, DSVL_L2, &Dlg_Open_DSVL::OnLvnKeydownL2)
	ON_NOTIFY(LVN_KEYDOWN, DSVL_L3, &Dlg_Open_DSVL::OnLvnKeydownL3)
	ON_NOTIFY(NM_CLICK, DSVL_L2, &Dlg_Open_DSVL::OnNMClickL2)
	ON_NOTIFY(NM_CLICK, DSVL_L3, &Dlg_Open_DSVL::OnNMClickL3)
	ON_EN_KILLFOCUS(DSVL_ED2, &Dlg_Open_DSVL::OnEnKillfocusEd2)
	ON_EN_KILLFOCUS(DSVL_ED3, &Dlg_Open_DSVL::OnEnKillfocusEd3)	
	ON_EN_SETFOCUS(DSVL_T1, &Dlg_Open_DSVL::OnEnSetfocusT1)
	ON_BN_CLICKED(DSVL_CBS1, &Dlg_Open_DSVL::OnBnClickedCbs1)
	ON_BN_CLICKED(DSVL_CBS2, &Dlg_Open_DSVL::OnBnClickedCbs2)
	ON_EN_SETFOCUS(DSVL_ED1, &Dlg_Open_DSVL::OnEnSetfocusEd1)
	ON_EN_SETFOCUS(DSVL_ED2, &Dlg_Open_DSVL::OnEnSetfocusEd2)
	ON_EN_SETFOCUS(DSVL_ED3, &Dlg_Open_DSVL::OnEnSetfocusEd3)
	ON_BN_CLICKED(DSVL_CBS3, &Dlg_Open_DSVL::OnBnClickedCbs3)
	ON_BN_CLICKED(DSVL_B12, &Dlg_Open_DSVL::OnBnClickedB12)	
	ON_BN_CLICKED(DSVL_CBS4, &Dlg_Open_DSVL::OnBnClickedCbs4)	
	ON_NOTIFY(NM_RCLICK, DSVL_L2, &Dlg_Open_DSVL::OnNMRClickL2)
	ON_COMMAND(ID_TI40039_6, &Dlg_Open_DSVL::OnTi40039)
	ON_COMMAND(ID_TI40040_6, &Dlg_Open_DSVL::OnTi40040)
	ON_COMMAND(ID_TI40052, &Dlg_Open_DSVL::OnTi40052)
	ON_NOTIFY(NM_DBLCLK, DSVL_L2, &Dlg_Open_DSVL::OnNMDblclkL2)
	ON_NOTIFY(NM_DBLCLK, DSVL_L3, &Dlg_Open_DSVL::OnNMDblclkL3)
	ON_BN_CLICKED(DSVL_PRE, &Dlg_Open_DSVL::OnBnClickedPre)
	ON_BN_CLICKED(DSVL_NXT, &Dlg_Open_DSVL::OnBnClickedNxt)
	ON_CBN_SELCHANGE(DSVL_CBBFL, &Dlg_Open_DSVL::OnCbnSelchangeCbbfl)
	ON_BN_CLICKED(DSVL_CKLUU, &Dlg_Open_DSVL::OnBnClickedCkluu)
	ON_NOTIFY(LVN_ENDSCROLL, DSVL_L1, &Dlg_Open_DSVL::OnLvnEndScrollL1)
END_MESSAGE_MAP()

BEGIN_EASYSIZE_MAP(Dlg_Open_DSVL)
	EASYSIZE(DSVL_S1,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(DSVL_T1,ES_BORDER,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(DSVL_PRE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(DSVL_NXT,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(DSVL_CBBFL,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(DSVL_S111,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(DSVL_L1,ES_BORDER,ES_BORDER,ES_BORDER,ES_BORDER,0)
	EASYSIZE(DSVL_B1,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(DSVL_B2,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(DSVL_B3,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(DSVL_B4,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(DSVL_B5,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(DSVL_B6,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)

	EASYSIZE(DSVL_G2,ES_BORDER,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(DSVL_L2,ES_BORDER,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(DSVL_L3,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)

	EASYSIZE(DSVL_CBS4,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)
	EASYSIZE(DSVL_CBB1,ES_BORDER,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)

	EASYSIZE(DSVL_B7,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(DSVL_B8,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(DSVL_B9,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(DSVL_B10,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(DSVL_B11,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)

	EASYSIZE(DSVL_S9,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)
	EASYSIZE(DSVL_CBS3,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)
	EASYSIZE(DSVL_B12,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)
	EASYSIZE(DSVL_S10,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)

	EASYSIZE(DSVL_CBS1,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)
	EASYSIZE(DSVL_CBS2,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)
	EASYSIZE(DSVL_CBS5,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)

	EASYSIZE(DSVL_CKLUU,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(DSVL_T2,ES_BORDER,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
END_EASYSIZE_MAP

BOOL Dlg_Open_DSVL::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	iTCclick=-1;
	sDownTC=L"",sFullTC=L"";
	sLuuMaCT=L"",sLuuTenCT=L"",sLuuDVT=L"",sLuuLink=L"";
	sBefore=L"",sAfter=L"";

	iStopload=1;
	tslkq=0,lanshow=1,ibuocnhay=50;
	_iTab=0,nClick=-1;
	_isl1=0,_isl2=0;
	f_Load_ToolTip();

	m_chkLuuTC.SetCheck(iLuuTCVL);
	m_txtTChon.SubclassDlgItem(DSVL_T2,this);
	m_txtTChon.SetBkColor(WHITE);
	m_txtTChon.SetTextColor(BLUE);

	m_Txt.SubclassDlgItem(DSVL_T1,this);
	m_Txt.SetBkColor(WHITE);
	m_Txt.SetTextColor(BLUE);

	m_Txt.SetWindowText(xl_tukhoa);
	m_Txt.SetCueBanner(L"Nhập từ khóa tìm kiếm (Ví dụ: V1014, ống nước D50,... )");

	m_chkColor.SetCheck(chkClrVL);
	m_btnColor.EnableWindow(chkClrVL);
	m_txtColor.ShowWindow(chkClrVL);

	m_txtTChon.EnableWindow(iLuuTCVL);
	m_cbbox.EnableWindow(jNhomTC);
	m_chkBTC.SetCheck(jNhomTC);

	m_txtColor.SubclassDlgItem(DSCV_S10,this);
	m_txtColor.SetBkColor(fclrVL);
	m_txtColor.SetTextColor(BLACK);

	m_Group2.SubclassDlgItem(DSVL_S111,this);
	m_Group2.SetBkColor(YELLOW);
	m_Group2.SetTextColor(BLACK);

	f_StyleList();
	f_StyleList2();
	f_StyleList3();
	f_Load_DL();	

	if(iCpy==0)
	{
		m_cbk1.SetCheck(1);
		m_cbk2.SetCheck(1);
	}
	else if(iCpy==1) m_cbk1.SetCheck(1);
	else if(iCpy==2) m_cbk2.SetCheck(1);

	this->SetWindowTextW(L"Danh sách vật liệu | " + sFullPath);

	_GetDanhsachnhom();
	_LoadlistFile();

	INIT_EASYSIZE;

	// Set size dialog mfc
	if(irWVL>0 && irHVL>0)
	{
		this->SetWindowPos(NULL,0,0,irWVL,irHVL,SWP_NOMOVE | SWP_NOZORDER);

		CRect r;
		GetWindowRect(&r);
		ScreenToClient(&r);
		int isx = GetSystemMetrics(SM_CXSCREEN);
		int isy = GetSystemMetrics(SM_CYSCREEN);
		MoveWindow((isx-r.Width())/2, (isy-r.Height())/2 , r.right - r.left, r.bottom - r.top, TRUE);
	}

	return TRUE; 
}

// Bắt sự kiện click Enter
BOOL Dlg_Open_DSVL::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if (pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(DSVL_T1))
	{
		OnBnClickedB1();
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(DSVL_L1))
	{
		OnBnClickedB2();
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		if(ClickEsc==0) OnBnClickedB6();
		else GotoDlgCtrl(GetDlgItem(DSVL_L1));
		ClickEsc=0;	
		
		return TRUE; 
	}
	else if( pMsg->message == WM_LBUTTONDOWN ||
		pMsg->message == WM_LBUTTONUP ||
		pMsg->message == WM_MOUSEMOVE )
	{
		m_ToolTips.RelayEvent(pMsg);
		return CDialog::PreTranslateMessage(pMsg);
	}
	else if(pMsg->message == WM_CHAR)
	{
		TCHAR chr = static_cast<TCHAR>(pMsg->wParam);
		if(pWndWithFocus == GetDlgItem(DSVL_L2))
		{
			switch (chr)
			{
				case 0x04:
				{
					OnTi40039();
					break;
				}
				case 0x07:
				{
					OnTi40040();
					break;
				}

				default:
					break;
			}
		}
	}

	return FALSE;
}

void Dlg_Open_DSVL::OnSize(UINT nType, int cx, int cy) 
{
	CDialog::OnSize(nType, cx, cy);
	UPDATE_EASYSIZE;
}

void Dlg_Open_DSVL::OnSizing(UINT fwSide, LPRECT pRect) 
{
	CDialog::OnSizing(fwSide, pRect);
	EASYSIZE_MINSIZE(320,240,fwSide,pRect);
}


void Dlg_Open_DSVL::f_Load_ToolTip()
{
	CButton * btn0 = (CButton *)GetDlgItem(DSVL_B7);
	CButton * btn2 = (CButton *)GetDlgItem(DSVL_B8);
	CButton * btn3 = (CButton *)GetDlgItem(DSVL_B9);
	CButton * btn4 = (CButton *)GetDlgItem(DSVL_B10);
	CButton * btn5 = (CButton *)GetDlgItem(DSVL_B11);
	CButton * btn6 = (CButton *)GetDlgItem(DSVL_CBS4);

	CListCtrl * pls1 = (CListCtrl *)GetDlgItem(DSVL_L1);
	CListCtrl * pls2 = (CListCtrl *)GetDlgItem(DSVL_L2);

	m_ToolTips.Create(this);
	m_ToolTips.SetMaxTipWidth(600);
	m_ToolTips.SetDelayTime(900);

	m_ToolTips.AddTool(btn0, L"Thêm mới dữ liệu tiêu chuẩn hoặc nội dung vào vị trí mong muốn."
		L"\nKích chọn dòng đầu tiên để chèn lên trên đầu."
		L"\nKích chọn trống (NULL) dưới cùng để chèn xuống dưới cùng.");

	m_ToolTips.AddTool(btn2, L"Xóa dữ liệu tiêu chuẩn hoặc nội dung.");
	m_ToolTips.AddTool(btn3, L"Cập nhật các dữ liệu chỉnh sửa.");
	m_ToolTips.AddTool(btn4, L"Di chuyển tiêu chuẩn/nội dung lên trên.");
	m_ToolTips.AddTool(btn5, L"Di chuyển tiêu chuẩn/nội dung xuống dưới.");
	m_ToolTips.AddTool(btn6, L"Tích chọn để đưa tiêu chuẩn sang 1 bảng riêng biệt.");

	m_ToolTips.AddTool(pls1, L"Kích để áp dụng vật liệu. Giữ Shift/Ctrl để chọn nhiều."
		L"\nKích đúp để chỉnh sửa dữ liệu.");

	m_ToolTips.AddTool(pls2, L"Kích đúp để nhập từ khóa tra cứu tiêu chuẩn.");
}


void Dlg_Open_DSVL::f_Get_size()
{
	CRect rectCtrl;
	GetWindowRect(&rectCtrl);
	irHVL = rectCtrl.Height();
	irWVL = rectCtrl.Width();
}


void Dlg_Open_DSVL::f_StyleList()
{
	lit_dmvl.InsertColumn(0,L"",LVCFMT_LEFT,20);
	lit_dmvl.InsertColumn(1,L"Mã VL",LVCFMT_CENTER,_wV1[0]);
	lit_dmvl.InsertColumn(2,L"Tên vật liệu",LVCFMT_LEFT,_wV1[1]);
	lit_dmvl.InsertColumn(3,L"Kích thước mẫu",LVCFMT_LEFT,_wV1[2]);
	lit_dmvl.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES |  LVS_EX_CHECKBOXES | LVS_EX_LABELTIP);

	lit_dmvl.SetColumnReadOnly(0);
	lit_dmvl.SetColumnReadOnly(4);
	lit_dmvl.SetColumnColors(0,WHITE,WHITE);
	lit_dmvl.SetDefaultEditor(NULL, NULL, &txtE_DMVL);
	txtE_DMVL.SubclassDlgItem(DSVL_ED1,this);txtE_DMVL.SetBkColor(YELLOW153);txtE_DMVL.SetTextColor(BLUE);
}


void Dlg_Open_DSVL::f_StyleList2()
{
	lit_tcvl.InsertColumn(0,L"",LVCFMT_CENTER,0);
	lit_tcvl.InsertColumn(1,L"Mã TC",LVCFMT_LEFT,_wV2[0]);
	lit_tcvl.InsertColumn(2,L"Tiêu chuẩn áp dụng",LVCFMT_LEFT,_wV2[1]);
	lit_tcvl.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);

	lit_tcvl.SetColumnReadOnly(0);
	lit_tcvl.SetColumnReadOnly(1);
	lit_tcvl.SetDefaultEditor(NULL, NULL, &txtE_TCVL);
	txtE_TCVL.SubclassDlgItem(DSVL_ED2,this);txtE_TCVL.SetBkColor(YELLOW153);txtE_TCVL.SetTextColor(BLUE);
}


void Dlg_Open_DSVL::f_StyleList3()
{
	lit_ppvl.InsertColumn(0,L"",LVCFMT_CENTER,0);
	lit_ppvl.InsertColumn(1,L"Mã LM",LVCFMT_LEFT,_wV3[0]);
	lit_ppvl.InsertColumn(2,L"Phương pháp lấy mẫu",LVCFMT_LEFT,_wV3[1]);
	lit_ppvl.InsertColumn(3,L"Tần suất",LVCFMT_LEFT,_wV3[2]);
	lit_ppvl.InsertColumn(4,L"Đơn vị",LVCFMT_LEFT,_wV3[3]);
	lit_ppvl.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);

	lit_ppvl.SetColumnReadOnly(0);
	lit_ppvl.SetColumnReadOnly(1);
	lit_ppvl.SetDefaultEditor(NULL, NULL, &txtE_PPVL);
	txtE_PPVL.SubclassDlgItem(DSVL_ED3,this);txtE_PPVL.SetBkColor(YELLOW153);txtE_PPVL.SetTextColor(BLUE);
}


void Dlg_Open_DSVL::f_SaveWidth()
{
	for (int i = 0; i < 3; i++) _wV1[i] = lit_dmvl.GetColumnWidth(i+1);
	for (int i = 0; i < 2; i++) _wV2[i] = lit_tcvl.GetColumnWidth(i+1);
	for (int i = 0; i < 4; i++) _wV3[i] = lit_ppvl.GetColumnWidth(i+1);
}


void Dlg_Open_DSVL::xl_xdcot_DMVL(_WorksheetPtr pSh)
{
	_iCot[1] = getColumn(pSh,"DMVL_STT");
	_iCot[2] = getColumn(pSh,"DMVL_MAVL");
	_iCot[3] = getColumn(pSh,"DMVL_MHDG");
	_iCot[4] = getColumn(pSh,"DMVL_MAHS");
	_iCot[5] = getColumn(pSh,"DMVL_ND");
	_iCot[6] = getColumn(pSh,"DMVL_TC");

	_iCot[7] = getColumn(pSh,"DMVL_NK_NG");
	_iCot[8] = getColumn(pSh,"DMVL_LM_NG");
	_iCot[9] = getColumn(pSh,"DMVL_NTNB_NG");
	_iCot[10] = getColumn(pSh,"DMVL_YC");
	_iCot[11] = getColumn(pSh,"DMVL_AB_NG");

	_iCot[13] = getColumn(pSh,"DMVL_TC2");
	_iCot[26] = getColumn(pSh,"DMVL_KBB");

	_iCot[27] = getColumn(pSh,L"DMVL_XXU");
	_iCot[28] = getColumn(pSh,L"DMVL_DOT");

	_iCotEND = getColumn(pSh, L"DMVL_MAUNTVL_END");
}


void Dlg_Open_DSVL::f_Danh_sachfile()
{
	try
	{
		CString spath = GIOR(iRowDLGoc,iColumnDLGoc,prTS,L"");
		CString szfolder = GIOR(iRowDLGoc+1,iColumnDLGoc,prTS,L"");
		if(spath.Right(1)!=L"\\") spath+=L"\\";		
		sDDanfile = spath + szfolder;

		slfileDL=0;
		CString szval=L"", scheck=L"";
		for (int i = iRowDLGoc+1; i < iRowDLGoc+100; i++)
		{
			szval = GIOR(i,iColumnDLGoc+2,prTS,L"");
			if(szval==L"") break;

			scheck = GIOR(i,iColumnDLGoc+3,prTS,L"");
			if(_wtoi(scheck)==1)
			{
				if(szval.Find(L"\\")==-1) scheck.Format(L"%s\\%s",sDDanfile,szval);
				else
				{
					scheck = szval;
					szval = _TackTenFile(szval,1);
				}

				if(_FileExists(scheck,0)==1)
				{
					slfileDL++;
					sFiledulieu[slfileDL] = szval;
					sFpathdulieu[slfileDL] = scheck;
				}
			}
		}

		if(slfileDL==0)
		{
			_s(L"Không tồn tại tệp tin dữ liệu. Vui lòng kiểm tra lại đường dẫn.");
			return;
		}
	}
	catch(...){}
}


void Dlg_Open_DSVL::f_Tim_kiem_DMVL()
{
	try
	{
		CString val = L"", szFormula = L"";

		psDMVL = pWb->ActiveSheet;
		iColumn = psDMVL->Application->ActiveCell->Column;

		// Xác định cột
		xl_xdcot_DMVL(psDMVL);

/////// Kiểm tra vị trí activate = cột tra mã
		// Xác định dòng
		iRow = psDMVL->Application->ActiveCell->Row;
		prDMVL = psDMVL->Cells;

		// Xác định vị trí END
		xde = FindComment(psDMVL,_iCot[1],"END");
		if(iRow<8 || iRow>=xde) return;

		if(iColumn==_iCot[7] || iColumn==_iCot[8] || iColumn==_iCot[9] 
		|| iColumn==_iCot[10] || iColumn==_iCot[11])
		{
			// Leo edit 01.06.2018
			val= GIOR(iRow,1,prDMVL,L"");
			if(_wtoi(val)>0) CheckNhomKy(psDMVL,iRow,getColumn(psDMVL,L"DMVL_KBB"));
			return;
		}
		else if(iColumn==_iCot[6] || iColumn==_iCot[13])
		{
			Dlg_trauu_tieuchuan _dlg;
			_dlg.xl_timkiem_tieuchuan();
			return;
		}
		else if(iColumn==_iCot[5])
		{
			// Đổi mã nếu đổi tên
			val= GIOR(iRow,1,prDMVL,L"");
			if(_wtoi(val)>0)
			{
				val = GIOR(iRow,_iCot[2],prDMVL,L"");
				if(val!=L"")
				{
					CString szMaL=val,szTenBF=L"",szTenAT=L"";
					if(szMaL.GetLength()>5 && (szMaL.Right(3)).Left(1)==L".") szMaL = val.Left(val.GetLength()-3);
					int vtL=0,demL=0,posL=0;

					while (true)
					{
						vtL = FindValue(psDMVL,_iCot[2],8,iRow-1,(_bstr_t)val);
						if(vtL==0) vtL = FindValue(psDMVL,_iCot[2],iRow+1,xde,(_bstr_t)val);
						if(vtL==0)
						{
							posL=1;
							break;
						}
						else
						{
							// Nếu đã tồn tại mã như vậy, tiếp tục kiểm tra xem tên có trùng không?
							// Nếu trùng bỏ qua, ngược lại tạo mã mới
							szTenBF = GIOR(vtL,_iCot[5],prDMVL,L"");
							szTenAT = GIOR(iRow,_iCot[5],prDMVL,L"");
							if(szTenBF==szTenAT) break;
						}

						demL++;
						if(demL<=9) val.Format(L"%s.0%d",szMaL,demL);
						else val.Format(L"%s.%d",szMaL,demL);
					}

					if (posL==1) prDMVL->PutItem(iRow,_iCot[2],(_bstr_t)val);
				}
			}

			return;
		}
		else if (iColumn == _iCot[2] || iColumn == _iCot[3])
		{
			// Get Formula
			if (iColumn == _iCot[2])
			{
				PRS = prDMVL->GetItem(iRow, iColumn);
				szFormula = PRS->GetFormula();
			}

			// UPPER
			xl_tukhoa = GIOR(iRow, iColumn, prDMVL, L"");
			xl_tukhoa.MakeUpper();
			prDMVL->PutItem(iRow, iColumn, (_bstr_t)xl_tukhoa);
		}

		if(iColumn!=_iCot[2]) return;

		// Tồn tại giá trị cột tiêu chuẩn
		val = GIOR(iRow,_iCot[1],prDMVL,L"");
		val.Trim();

		// Tồn tại giá trị cột tiêu chuẩn hoặc diễn giải
		CString sz5 = GIOR(iRow,_iCot[5],prDMVL,L""); sz5.Trim();
		CString sz6 = GIOR(iRow,_iCot[6],prDMVL,L""); sz6.Trim();

		if(val==L"" && (sz5!=L"" || sz6!=L"") && szFormula.Left(1)!=L"=") return;

// Hàm check bản quyền có điều kiện đếm
		if(f_Mod_check()!=1) return;

//////// Tất cả thỏa mãn, tiến hành tra cứu

		// Bổ sung 27.05.2019
		if(szFormula.Left(1)==L"=")
		{
			szFormula.Replace(L"=",L"");
			szFormula.Replace(L"$",L"");

			int iRowStart=0,iRowCopy=0;
			int posR = szFormula.Find(L":");
			if(posR==-1)
			{
				iRowCopy = _wtoi(_xlConvertA1toRC(szFormula,1));
				iRowStart = iRowCopy;
			}
			else
			{
				CString szLf = szFormula.Left(posR);
				CString szRg = szFormula.Right(szFormula.GetLength()-posR-1);

				iRowStart = _wtoi(_xlConvertA1toRC(szLf,1));
				iRowCopy = _wtoi(_xlConvertA1toRC(szRg,1));
			}

			if(iRowStart>0 && iRowCopy>0)
			{
				int ichecknull=0;
				val = GIOR(iRow,_iCot[2],prDMVL,L"");
				val.MakeUpper();
				if(val==L"")
				{
					val = L"'" + szFormula;
					prDMVL->PutItem(iRow,_iCot[2],(_bstr_t)val);
				}
				else if(val==L"0") ichecknull=1;

				int vtkt = iRow;
				if(iRowCopy>iRow) vtkt = xde;

				int iRowEndCopy=vtkt;
				if(ichecknull==0)
				{
					if(val==L"HM")
					{
						for (int i = iRowCopy+1; i < vtkt; i++)
						{
							val = GIOR(i,_iCot[2],prDMVL,L"");
							if(val==L"HM")
							{
								iRowEndCopy=i;
								break;
							}
						}
					}
					else if(val==L"GD")
					{
						for (int i = iRowCopy+1; i < vtkt; i++)
						{
							val = GIOR(i,_iCot[2],prDMVL,L"");
							if(val==L"HM" || val==L"GD")
							{
								iRowEndCopy=i;
								break;
							}
						}
					}
					else
					{
						for (int i = iRowCopy+1; i < vtkt; i++)
						{
							val = GIOR(i,_iCot[2],prDMVL,L"");
							if(val!=L"")
							{
								iRowEndCopy=i;
								break;
							}
						}
					}
				}
				else iRowEndCopy = iRowCopy+1;

				if(iRowEndCopy > iRowCopy+1)
				{
					for (int i = iRowEndCopy-1; i >= iRowCopy; i--)
					{
						sz5 = GIOR(i,_iCot[5],prDMVL,L""); sz5.Trim();
						sz6 = GIOR(i,_iCot[6],prDMVL,L""); sz6.Trim();
						if(sz5!=L"" || sz6!=L"")
						{
							iRowEndCopy=i+1;
							break;
						}
					}
				}

				int slcp = iRowEndCopy-iRowStart;
				if(iRowStart>iRow)
				{
					iRowStart+=slcp;
					iRowEndCopy+=slcp;
				}
				xde += slcp;

				PRS = psDMVL->GetRange(prDMVL->GetItem(iRow+1,1),prDMVL->GetItem(iRow+slcp,1));
				PRS->EntireRow->Insert(xlShiftDown);
					
				PRS = psDMVL->GetRange(prDMVL->GetItem(iRowStart, 1), prDMVL->GetItem(iRowEndCopy - 1, 1));

				bool blCopyAll = true;	// Tạm thời để copy tất cả (giá trị, công thức,...)
				if (blCopyAll)
				{
					PRS->EntireRow->Copy(prDMVL->GetItem(iRow, 1));
				}
				else
				{
					PRS->EntireRow->Copy();

					PRS2 = prDMVL->GetItem(iRow, 1);
					PRS2->PasteSpecial(xlPasteValues, Excel::xlPasteSpecialOperationNone);
					PRS2->PasteSpecial(xlPasteFormats, Excel::xlPasteSpecialOperationNone);
					PRS2->PasteSpecial(xlPasteComments, Excel::xlPasteSpecialOperationNone);
				}
				
				xl->PutCutCopyMode(XlCutCopyMode(false));

				shTS = name_sheet("shTS");
				psTS = xl->Sheets->GetItem(shTS);
				prTS = psTS->GetCells();

				int ntype = _wtoi(getVCell(psTS,L"TS_ATOSTT"));
				if(ntype!=0 && ntype!=1 && ntype!=2) ntype=0;
				f_Auto_stt_dmuc(prDMVL,ntype,0,8,xde,_iCot[1],_iCot[2],_iCot[4]);

				if(slcp>5) PRS = prDMVL->GetItem(iRow,_iCot[2]);
				else
				{
					if(posR>0) PRS = prDMVL->GetItem(iRow+slcp,_iCot[2]);
					else PRS = prDMVL->GetItem(iRow+slcp-1,_iCot[2]);
				}
				PRS->Select();

				return;			
			}
		}

/*********************************************************************/

		getPathQLCL();
		f_Danh_sachfile();

		jNhomTC = _wtoi(getVCell(psTS,L"TS_GRPTC"));
		iCpy = _wtoi(getVCell(psTS,L"CF_DMCOPY"));

		shBTC = name_sheet("shNhomTC");
		psBTC = xl->Sheets->GetItem(shBTC);
		prBTC = psBTC->GetCells();
		int iVisible = psBTC->GetVisible();
		if (iVisible!=-1) psBTC->PutVisible(XlSheetVisibility::xlSheetVisible);

		// Xác định từ khóa tìm kiếm
		xl_tukhoa = GIOR(iRow,iColumn,prDMVL,L"");
		xl_tukhoa.Replace(L"'",L"");
		xl_tukhoa.Trim();
		if(xl_tukhoa==L"")
		{
			SqlString.Format(L"SELECT * FROM tbTuDienVatTu WHERE MSVT LIKE 'V%s' ORDER BY MSVT ASC;",L"%");

			//if(sFiledulieu[1].Find(L"\\")>0) sFullPath = sFiledulieu[1];
			//else sFullPath.Format(L"%s\\%s",sDDanfile,sFiledulieu[1]);

			sFullPath = sFpathdulieu[1];
			
			AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
			xl->PutScreenUpdating(VARIANT_TRUE);
			DoModal();
			xl->PutScreenUpdating(VARIANT_FALSE);

			return;
		}

		// Trường hợp là ghi chú (21.03.2019)
		if(xl_tukhoa.Find(L"*")>=0)
		{
			PRS = prDMVL->GetItem(iRow,1);
			PRS->EntireRow->Font->PutBold(1);
			PRS->EntireRow->Font->PutItalic(1);
			return;
		}

		// Trường hợp là nhóm công tác
		if (xl_tukhoa.Left(1) == L"#")
		{
			typedef bool(__stdcall *p)();
			HINSTANCE loadDLL = LoadLibrary(L"groupdata.dll");
			if (loadDLL != NULL)
			{
				p pc = (p)GetProcAddress(loadDLL, "CallTranhomcongviec");
				if (pc != NULL) pc();
			}
			FreeLibrary(loadDLL);
			return;
		}

		// Các trường hợp đặc biệt (18.03.2016)
		CString sDbiet = xl_tukhoa.Left(1);
		if(sDbiet==L"*" || sDbiet==L"#" || sDbiet==L"@" || sDbiet==L"$" 
			|| sDbiet==L"+" || sDbiet==L"-" || sDbiet==L".") return;

		sDbiet = xl_tukhoa;
		sDbiet.MakeUpper();

		// Trường hợp là chữ cái
		for (int i = 0; i < 23; i++)
		{
			if(sDbiet==cSabc[i])
			{
				prDMVL->PutItem(iRow,iColumn,(_bstr_t)sDbiet);
				return;
			}
		}

		// Trường hợp là số La Mã
		CString valLM=L"";
		for (int i = 0; i < 20; i++)
		{
			valLM = cSpr[i]+L".";
			if(sDbiet==cSpr[i] || sDbiet==valLM)
			{
				prDMVL->PutItem(iRow,iColumn,(_bstr_t)sDbiet);
				return;
			}
		}

		// Các trường hợp đặc biệt khác (DG,DS)
		sDbiet = xl_tukhoa.Left(2);
		sDbiet.MakeUpper();
		if (sDbiet == L"HM" || sDbiet == L"GD")
		{
			PRS = psDMVL->GetRange(prDMVL->GetItem(iRow, 1), prDMVL->GetItem(iRow, _iCotEND));
			PRS->EntireRow->Font->PutBold(1);
			PRS->EntireRow->Font->PutItalic(0);
			PRS->EntireRow->Font->PutColor(RGB(0, 0, 0));

			if (sDbiet == L"HM")
			{
				PRS->EntireRow->Interior->PutColor(RGB(217, 217, 217));
				PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->Weight = xlThin;
				PRS->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->Weight = xlThin;
				PRS->Borders->GetItem(XlBordersIndex::xlInsideVertical)->LineStyle = xlNone;

				int demhm = 1;
				for (int j = 8; j < iRow; j++)
				{
					val = GIOR(j, _iCot[2], prDMVL, L"");
					if (val == L"HM") demhm++;
				}

				val.Format(L"HM%d", demhm);
				prDMVL->PutItem(iRow, _iCot[4], (_bstr_t)val);

				val.Format(L"Hạng mục %d", demhm);
				prDMVL->PutItem(iRow, _iCot[5], (_bstr_t)val);
			}
			else
			{
				PRS->EntireRow->Interior->PutColor(RGB(255, 255, 255));

				PRS->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->PutWeight(xlThin);
				PRS->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->PutLineStyle(xlDot);
				PRS->Borders->GetItem(XlBordersIndex::xlInsideVertical)->Weight = xlThin;

				val = GIOR(iRow - 1, _iCot[2], prDMVL, L"");
				if (val == L"HM")
				{
					PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->Weight = xlThin;
				}
				else
				{
					PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->PutWeight(xlThin);
					PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->PutLineStyle(xlDot);
				}

				int demgd = 1;
				for (int j = iRow - 1; j >= 8; j--)
				{
					val = GIOR(j, _iCot[2], prDMVL, L"");

					if (val == L"HM") break;
					if (val == L"GD") demgd++;
				}

				val.Format(L"GD%d", demgd);
				prDMVL->PutItem(iRow, _iCot[4], (_bstr_t)val);

				val.Format(L"Giai đoạn %d", demgd);
				prDMVL->PutItem(iRow, _iCot[5], (_bstr_t)val);
			}

			if (iRow + 1 == xde)
			{
				PRS = psDMVL->GetRange(prDMVL->GetItem(xde, 1), prDMVL->GetItem(xde + 4, 1));
				PRS->EntireRow->Insert(xlShiftDown);

				PRS = psDMVL->GetRange(prDMVL->GetItem(xde, 1), prDMVL->GetItem(xde + 4, _iCotEND));
				PRS->EntireRow->Font->PutBold(1);
				PRS->EntireRow->Font->PutItalic(0);
				PRS->EntireRow->Font->PutColor(RGB(0, 0, 0));
				PRS->EntireRow->Interior->PutColor(RGB(255, 255, 255));

				PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->PutWeight(xlThin);
				PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->PutLineStyle(xlDot);

				PRS->Borders->GetItem(XlBordersIndex::xlInsideHorizontal)->PutWeight(xlThin);
				PRS->Borders->GetItem(XlBordersIndex::xlInsideHorizontal)->PutLineStyle(xlDot);

				PRS->Borders->GetItem(XlBordersIndex::xlInsideVertical)->Weight = xlThin;

				xde += 5;
				PRS = psDMVL->GetRange(prDMVL->GetItem(xde, 1), prDMVL->GetItem(xde, _iCotEND));
				PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->Weight = xlThin;
			}

			return;
		}

		// Bổ sung 20.08.2018
		// Tìm mã tương tự ở trên bảng tính
		CString szval=L"";

		int irTrung=0;
		CString sktrU = xl_tukhoa;
		sktrU.MakeUpper();
		for (int i = 8; i < iRow-1; i++)
		{
			szval = GIOR(i,iColumn,prDMVL,L"");
			szval.MakeUpper();
			if(szval==sktrU)
			{
				irTrung=i;
				break;
			}
		}

		if(irTrung==0)
		{
			for (int i = iRow+1; i < xde; i++)
			{
				szval = GIOR(i,iColumn,prDMVL,L"");
				szval.MakeUpper();
				if(szval==sktrU)
				{
					irTrung=i;
					break;
				}
			}
		}

		if(irTrung>0)
		{
			szval.Format(L"Đã tồn tại mã '%s' trong bảng tính. "
				L"\nBạn có muốn vận dụng luôn mã này?",xl_tukhoa);
			int gtz = _yn(szval,1,1);
			if(gtz!=6 && gtz!=7) return;
			if(gtz==6)
			{
				PRS = prDMVL->GetItem(irTrung,1);
				PRS->EntireRow->Copy();

				PRS2 = prDMVL->GetItem(iRow,1);
				PRS2->PasteSpecial(xlPasteValues,Excel::xlPasteSpecialOperationNone);
				xl->PutCutCopyMode(XlCutCopyMode(false));

				PRS = prDMVL->GetItem(iRow,1);
				PRS->EntireRow->Font->PutColor(RGB(0,0,0));
				PRS->EntireRow->Font->PutBold(0);
				PRS->EntireRow->Font->PutItalic(0);
				PRS->EntireRow->Rows->AutoFit();
				if((int)PRS->GetRowHeight()>80) PRS->PutWrapText(FALSE);

				PRS = psDMVL->GetRange(prDMVL->GetItem(iRow,1),prDMVL->GetItem(iRow, _iCotEND));
				if(chkClrVL==1) PRS->Interior->PutColor(fclrVL);
				else PRS->EntireRow->Interior->PutColor(xlNone);

				// Thay đổi mã công việc và mã hồ sơ nghiệm thu
				/*CString bsMaCV = GIOR(irTrung,_iCot[2],prDMVL,L"");
				if(bsMaCV.GetLength()>5 && (bsMaCV.Right(3)).Left(1)==L".") bsMaCV = bsMaCV.Left(bsMaCV.GetLength()-3);

				for (int i = 1; i < 100; i++)
				{
					if(i<=9) szval.Format(L"%s.0%d",bsMaCV,i);
					else szval.Format(L"%s.%d",bsMaCV,i);

					 if((int)FindValue(psDMVL,_iCot[2],1,0,(_bstr_t)szval)==0)
					 {
						 bsMaCV = L"'" + szval;
						 prDMVL->PutItem(iRow,_iCot[2],(_bstr_t)bsMaCV);
						 break;
					 }
				}*/

				CString bsMaHSNT = GIOR(irTrung,_iCot[4],prDMVL,L"");
				for (int i = 1; i < 100; i++)
				{
					if(i<=9) szval.Format(L"%s.0%d",bsMaHSNT,i);
					else szval.Format(L"%s.%d",bsMaHSNT,i);

					 if((int)FindValue(psDMVL,_iCot[4],1,0,(_bstr_t)szval)==0)
					 {
						 bsMaHSNT = L"'" + szval;
						 prDMVL->PutItem(iRow,_iCot[4],(_bstr_t)bsMaHSNT);
						 break;
					 }
				}

				// Sắp xếp lại stt
				int ntype = _wtoi(getVCell(psTS,L"TS_ATOSTT"));
				if(ntype!=0 && ntype!=1 && ntype!=2) ntype=0;
				f_Auto_stt_dmuc(prDMVL,ntype,0,8,xde,_iCot[1],_iCot[2],_iCot[4]);				

				return;
			}
		}

		// Bổ sung 18.03.2019
		if((int)xl_tukhoa.Find(L" ")==-1 && (int)xl_tukhoa.Find(L"+")==-1)
		{
			if(xl_tukhoa.GetLength()>5 && (xl_tukhoa.Right(3)).Left(1)==L".") xl_tukhoa = xl_tukhoa.Left(xl_tukhoa.GetLength()-3);
		}

		// Bổ sung 16.12.2017
		if((int)xl_tukhoa.Find(L" ")==-1 && (int)xl_tukhoa.Find(L"+")==-1 && (int)xl_tukhoa.Find(L"|")>0)
		{
			_TackTukhoa(xl_tukhoa,L"|",L" ");
			SqlString = L"SELECT * FROM tbTuDienVatTu WHERE ((";
			for (int i = 1; i <= iKey; i++)
			{
				szval.Format(L"MSVT LIKE '%s%s%s' OR ",L"%",pKey[i],L"%");
				SqlString+=szval;
			}

			if(SqlString.Right(4)==L" OR ") SqlString = SqlString.Left(SqlString.GetLength()-4);
			szval.Format(L") AND MSVT LIKE 'V%s') ORDER BY MSVT ASC;",L"%");
			SqlString+=szval;
		}
		else
		{
			// Truy vấn & load dữ liệu
			f_Tack_tukhoa(xl_tukhoa);
			f_Tao_truyvan();
		}

///////// Truy vấn dữ liệu
		f_Truy_van_DL();

		if(iKqua>=1)
		{
			// Có nhiều kết quả -> hiển thị hộp thoại
			AFX_MANAGE_STATE(AfxGetStaticModuleState());
			xl->PutScreenUpdating(VARIANT_TRUE);
			DoModal();
			xl->PutScreenUpdating(VARIANT_FALSE);
		}
		else
		{
			// Không tồn tại dữ liệu
			PRS = prDMVL->GetItem(iRow,iColumn);
			PRS->Font->PutColor(RGB(255,0,0));
			PRS->Activate();
		}
	}
	catch(...){_s(L"Lỗi tra cứu dữ liệu vật liệu.");}
}

void Dlg_Open_DSVL::f_Tack_tukhoa(CString sKey)
{
	try
	{
		sKey+=L"+";
		iKey=0;
		int vtri=0;
		CString val=L"";
		for (int i = 0; i < sKey.GetLength(); i++)
		{
			val = sKey.Mid(i,1);
			if((val==L" " || val==L"+") && i>vtri)
			{
				iKey++;
				pKey[iKey] = sKey.Mid(vtri,i-vtri);
				vtri=i+1;

				if(iKey==98) break;
			}
		}
	}
	catch(...){}
}

void Dlg_Open_DSVL::f_Tao_truyvan()
{
	try
	{
		CString szval=L"";
		if(iKey==1) SqlString.Format(L"SELECT * FROM tbTuDienVatTu WHERE (MSVT LIKE '%s%s%s' OR TVT LIKE '%s%s%s');",L"%",pKey[1],L"%",L"%",pKey[1],L"%");
		else
		{
			SqlString.Format(L"SELECT * FROM tbTuDienVatTu WHERE (((MSVT LIKE '%s%s%s' OR TVT LIKE '%s%s%s')",L"%",pKey[1],L"%",L"%",pKey[1],L"%");
			for (int i = 2; i <= iKey; i++)
			{
				szval.Format(L" AND (MSVT LIKE '%s%s%s' OR TVT LIKE '%s%s%s')",L"%",pKey[i],L"%",L"%",pKey[i],L"%");
				SqlString+=szval;
			}

			szval.Format(L") AND MSVT LIKE 'V%s') ORDER BY MSVT ASC;",L"%");
			SqlString+=szval;
		}
	}
	catch(...){}
}

void Dlg_Open_DSVL::f_Truy_van_DL()
{
	TRY
	{
		CString szval=L"";
		for (int i = 1; i <= slfileDL; i++)
		{
			//if(sFiledulieu[i].Find(L"\\")>0) szval = sFiledulieu[i];
			//else szval.Format(L"%s\\%s",sDDanfile,sFiledulieu[i]);

			szval = sFpathdulieu[1];
			ObjConn.OpenAccessDB(szval,L"",L"");
			CRecordset* Recset = ObjConn.Execute(SqlString);
			iKqua=0;
			while( !Recset->IsEOF() )
			{
				iKqua=1;
				Recset->GetFieldValue(L"MSVT",XLXD[1].CDG[0]);
				Recset->GetFieldValue(L"TVT",XLXD[1].CDG[1]);
				sFullPath = szval;
				break;
			}
			Recset->Close();
			ObjConn.CloseDatabase();
			if(iKqua>0) return;
		}		
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[1] Lỗi tìm kiếm dữ liệu vật liệu")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_Open_DSVL::f_Load_DL()
{
	TRY
	{
		// Kết nối SQL
		ObjConn.OpenAccessDB(sFullPath,L"",L"");
		CRecordset* Recset = ObjConn.Execute(SqlString);

		tslkq=0;
		while( !Recset->IsEOF() )
		{
			tslkq++;
			Recset->GetFieldValue(L"MSVT",XLXD[tslkq].CDG[0]);
			Recset->GetFieldValue(L"TVT",XLXD[tslkq].CDG[1]);
			Recset->MoveNext();

			if (tslkq >= 4000) break;
		}
		Recset->Close();
		ObjConn.CloseDatabase();

		// Chỉnh sửa 03.07.2017 --> Thêm nút hiển thị thêm kết quả (chỉ show ibuocnhay kq)
		lanshow=1;
		int iz=ibuocnhay;
		if(tslkq<=iz)
		{
			iz=tslkq;
			iStopload=0;
		}
		else iStopload=1;

		for (int i = 0; i < iz; i++)
		{
			lit_dmvl.InsertItem(i,L"",0);
			lit_dmvl.SetItemText(i,1,XLXD[i+1].CDG[0]);
			lit_dmvl.SetItemText(i,2,XLXD[i+1].CDG[1]);
		}

		// Kết quả tìm kiếm
		CString szval=L"";
		if (tslkq >= 4000) szval = L"Tìm thấy rất nhiều kết quả";
		else szval.Format(L"Tổng số: %d kết quả ",tslkq);
		m_Group2.SetWindowText(szval);

		if(iz>0) GotoDlgCtrl(GetDlgItem(DSVL_L1));
		else GotoDlgCtrl(GetDlgItem(DSVL_T1));
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[2] Lỗi tìm kiếm dữ liệu vật liệu")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_Open_DSVL::OnBnClickedB1()
{
	try
	{
		ClickEsc=0;

		// Xác định từ khóa
		CString val=L"";
		m_Txt.GetWindowTextW(val);
		val.Replace(L"'",L"");
		val.Trim();

		_iTab=0,nClick=-1;
		_isl1=0,_isl2=0;

		// Xóa toàn bộ dữ liệu trên List Control
		lit_dmvl.DeleteAllItems();
		ASSERT(lit_dmvl.GetItemCount() == 0);

		lit_tcvl.DeleteAllItems();
		ASSERT(lit_tcvl.GetItemCount() == 0);

		lit_ppvl.DeleteAllItems();
		ASSERT(lit_ppvl.GetItemCount() == 0);

		if(val!=L"")
		{
			// Bổ sung 18.03.2019
			if((int)val.Find(L" ")==-1 && (int)val.Find(L"+")==-1)
			{
				if(val.GetLength()>5 && (val.Right(3)).Left(1)==L".") val = val.Left(val.GetLength()-3);
			}

			// Bổ sung 16.12.2017
			if((int)val.Find(L" ")==-1 && (int)val.Find(L"+")==-1 && (int)val.Find(L"|")>0)
			{
				_TackTukhoa(val,L"|",L" ");
				CString szval=L"";

				SqlString = L"SELECT * FROM tbTuDienVatTu WHERE ((";
				for (int i = 1; i <= iKey; i++)
				{
					szval.Format(L"MSVT LIKE '%s%s%s' OR ",L"%",pKey[i],L"%");
					SqlString+=szval;
				}

				if(SqlString.Right(4)==L" OR ") SqlString = SqlString.Left(SqlString.GetLength()-4);
				
				val.Format(L") AND MSVT LIKE 'V%s') ORDER BY MSVT ASC;",L"%");
				SqlString+=val;
			}
			else
			{
				// Truy vấn & load dữ liệu
				f_Tack_tukhoa(val);
				f_Tao_truyvan();
			}
		}
		else SqlString.Format(L"SELECT * FROM tbTuDienVatTu WHERE MSVT LIKE 'V%s' ORDER BY MSVT ASC;",L"%");
		
		f_Load_DL();
	}
	catch(...){}
}

// 09.06.2017: Kiểm tra tiêu chuẩn trùng
int Dlg_Open_DSVL::f_Check_tieuchuan(CString sKtraTC)
{
	try
	{
		if(slMaTC==0) return 0;
		for (int i = 1; i <= slMaTC; i++)
			if(sKtraTC==sKTMaTC[i]) return 1;
	}
	catch(...){return 0;}

	return 0;
}


void Dlg_Open_DSVL::_GetDanhsachnhom()
{
	try
	{
		int pRowEnd = prBTC->SpecialCells(xlCellTypeLastCell)->GetRow();
		int cotMa = getColumn(psBTC,L"NHTC_TEN");
		int dongMa = getRow(psBTC,L"NHTC_TEN");
		if(pRowEnd<=dongMa) return;

		m_cbbox.AdjustDroppedWidth();
		m_cbbox.SetMode(CComboBoxExt::MODE_AUTOCOMPLETE);

		CString szval=L"";
		for (int i = dongMa+1; i <= pRowEnd; i++)
		{
			szval = GIOR(i,cotMa,prBTC,L"");
			szval.Trim();
			if(szval!=L"") m_cbbox.AddString(szval);
		}

		m_cbbox.SetCueBanner(L"Đặt tên nhóm (Bạn có thể bỏ qua)");
	}
	catch(...){}
}

void Dlg_Open_DSVL::_LoadlistFile()
{
	try
	{
		int ivt=0,ilen=0;
		CString szval=L"";
		for (int i = 1; i <= slfileDL; i++)
		{
			/*if(sFiledulieu[i].Find(L"\\")==-1) szval = sFiledulieu[i];
			else
			{
				ilen = sFiledulieu[i].GetLength();
				for (int j = ilen-1; j > 0; j--)
				{
					if(sFiledulieu[i].Mid(j,1)==L"\\")
					{
						szval = sFiledulieu[i].Right(ilen-j-1);
						break;
					}
				}
			}*/

			m_cbbFile.AddString(sFiledulieu[i]);
			if(sFullPath==sFpathdulieu[i]) ivt=i-1;
		}

		m_cbbFile.SetCurSel(ivt);

		if(slfileDL==1)
		{
			m_btnPre.EnableWindow(0);
			m_btnNxt.EnableWindow(0);
			m_cbbFile.EnableWindow(0);
		}
	}
	catch(...){}
}

void Dlg_Open_DSVL::OnBnClickedB2()
{
	try
	{
		ClickEsc=0;

		int iChekLmau = m_chkLMau.GetCheck();

		// Kiểm tra có tích chọn hay không?
		int iTCCVluutam = iLuuTCVL;
		if(iTCCVluutam==1 && sLuuMaCT==L"") iTCCVluutam=0;

		// Kiểm tra có checkbox hay không?
		int tonglkq = lit_dmvl.GetItemCount();
		int ktrchk=0;
		for (int i = 0; i < tonglkq; i++)
		{
			if((int)lit_dmvl.GetCheck(i)==1)
			{
				ktrchk=1;
				break;
			}
		}

		POSITION pos=lit_dmvl.GetFirstSelectedItemPosition();
		if(ktrchk==0 && pos==NULL)
		{
			_s(L"Bạn chưa chọn dữ liệu. Vui lòng kiểm tra lại.");
			return;
		}

		xl->PutScreenUpdating(VARIANT_FALSE);

		// Kiểm tra tên đã tồn tại chưa?
		CString szval=L"";
		int icTen = getColumn(psBTC,L"NHTC_TEN");

		vtnhom=0;
		jNhomTC = m_chkBTC.GetCheck();
		if(jNhomTC==1)
		{
			m_cbbox.GetWindowTextW(sAfter);
			sAfter.Trim();
			if(sAfter!=L"") vtnhom = FindValue(psBTC,icTen,1,0,(_bstr_t)sAfter);
		}

////////////////////////////////////////////////////////////////////
	// 23.01.2016
	int Nb1 = m_cbk1.GetCheck();
	int Nb2 = m_cbk2.GetCheck();
	if(Nb1==1 && Nb2==1) iCpy=0;
	else if(Nb1==1 && Nb2==0) iCpy=1;
	else if(Nb1==0 && Nb2==1) iCpy=2;
	else iCpy=-1;

	putVal(psTS,"CF_DMCOPY",iCpy);
	putVal(psTS,"TS_GRPTC",jNhomTC);

	psDMVL = pWb->ActiveSheet;
	int rowac = psDMVL->Application->ActiveCell->Row;
	int colac = psDMVL->Application->ActiveCell->Column;
	prDMVL=  psDMVL->Cells;

	xl_xdcot_DMVL(psDMVL);	
	int iEnd = FindComment(psDMVL,_iCot[1],"END");

/////////////////////////////////////////////////////////////////////

	ObjConn2.OpenAccessDB(pathMDB,L"",L"");
	CRecordset *RecsetTC, *RecsetND;
	CString sMSCV=L"",sNA=L"",sDVT=L"",sLLink=L"",val=L"",stemp=L"",valtmp=L"";
	int fla=0;
	slMaTC=0;

	if(iTCCVluutam==1)
	{
		// Bổ sung 16.12.2017
		_TackTukhoa(sLuuMaCT,L"|",L" ");
		sMSCV = sLuuMaCT;
		sNA = sLuuTenCT;
		//sDVT = sLuuDVT;
		//sLLink = sLuuLink;

		if(sMSCV.Right(1)==L"|") sMSCV = sMSCV.Left(sMSCV.GetLength()-1);
		sMSCV.Replace(L"|",L"|\n");

		if(sNA.Right(1)==L"|") sNA = sNA.Left(sNA.GetLength()-1);
		sNA.Replace(L"|",L";\n");

		//if(sDVT.Right(1)==L"|") sDVT = sDVT.Left(sDVT.GetLength()-1);
		//sDVT.Replace(L"|",L"|\n");

		//if(sLLink.Right(1)==L"|") sLLink = sLLink.Left(sLLink.GetLength()-1);
		//sLLink.Replace(L"|",L";\n");

		for(int i=1; i<=iKey; i++)
		{
			fla=1;

			//Get TC
			SqlString.Format(L"SELECT * FROM tbMaVL_Tieuchuan WHERE MaVL LIKE '%s';",pKey[i]);
			RecsetTC = ObjConn2.Execute(SqlString);

			while( !RecsetTC->IsEOF())
			{
				RecsetTC->GetFieldValue(L"MaTC",stemp);
				if(f_Check_tieuchuan(stemp)==0)
				{
					slMaTC++;
					sKTMaTC[slMaTC] = stemp;
					sKTTenTC[slMaTC] = stemp + L": Không tìm thấy tên tiêu chuẩn";

					SqlString.Format(L"SELECT * FROM tbTentieuchuan WHERE MaTC LIKE '%s';",stemp);
					RecsetND = ObjConn2.Execute(SqlString);
					while( !RecsetND->IsEOF())
					{
						RecsetND->GetFieldValue(L"TenTC",sKTTenTC[slMaTC]);
						break;
					}
					RecsetND->Close();
				}
				RecsetTC->MoveNext();
			}
			RecsetTC->Close();
		}
	}
	else
	{
		if(ktrchk==1)
		{
			for(int i=0;i<lit_dmvl.GetItemCount();i++)
			if ((int)lit_dmvl.GetCheck(i)==1)
			{
				fla=1;
				val = lit_dmvl.GetItemText(i,1);

				//Get MaVL & NA--------------
				if(sMSCV==L"") sMSCV.Format(L"%s",val);
				else sMSCV.Format(L"%s|\n%s",sMSCV,val);

				valtmp = lit_dmvl.GetItemText(i,2);
				if(sNA==L"") sNA.Format(L"%s",valtmp);
				else sNA.Format(L"%s;\n%s",sNA,valtmp);

				//if(sLLink==L"") sLLink.Format(L"%s",sFullPath);
				//else sLLink.Format(L"%s|\n%s",sLLink,sFullPath);

				//Get MaVL & NA--------------/

				//Get TC
				SqlString.Format(L"SELECT *FROM tbMaVL_Tieuchuan WHERE MaVL LIKE '%s';",val);
				RecsetTC = ObjConn2.Execute(SqlString);
				//Fill TC

				while( !RecsetTC->IsEOF())
				{
					RecsetTC->GetFieldValue(L"MaTC",stemp);
					if(f_Check_tieuchuan(stemp)==0)
					{
						slMaTC++;
						sKTMaTC[slMaTC] = stemp;

						SqlString.Format(L"SELECT *FROM tbTentieuchuan WHERE MaTC LIKE '%s';",stemp);
						RecsetND = ObjConn2.Execute(SqlString);
						while( !RecsetND->IsEOF())
						{
							RecsetND->GetFieldValue(L"TenTC",sKTTenTC[slMaTC]);
							break;
						}
						RecsetND->Close();					
					}
					RecsetTC->MoveNext();
				}
				RecsetTC->Close();
			}
		}
		else
		{
			for(int i=0;i<lit_dmvl.GetItemCount();i++)
			if (lit_dmvl.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
			{
				fla=1;
				val = lit_dmvl.GetItemText(i,1);

				//Get MaVL & NA--------------
				if(sMSCV==L"") sMSCV.Format(L"%s",val);
				else sMSCV.Format(L"%s|\n%s",sMSCV,val);

				valtmp = lit_dmvl.GetItemText(i,2);
				if(sNA==L"") sNA.Format(L"%s",valtmp);
				else sNA.Format(L"%s;\n%s",sNA,valtmp);

				//if(sLLink==L"") sLLink.Format(L"%s",sFullPath);
				//else sLLink.Format(L"%s|\n%s",sLLink,sFullPath);

				//Get MaVL & NA--------------/

				//Get TC
				SqlString.Format(L"SELECT *FROM tbMaVL_Tieuchuan WHERE MaVL LIKE '%s';",val);
				RecsetTC = ObjConn2.Execute(SqlString);
				//Fill TC

				while( !RecsetTC->IsEOF())
				{
					RecsetTC->GetFieldValue(L"MaTC",stemp);
					if(f_Check_tieuchuan(stemp)==0)
					{
						slMaTC++;
						sKTMaTC[slMaTC] = stemp;

						SqlString.Format(L"SELECT *FROM tbTentieuchuan WHERE MaTC LIKE '%s';",stemp);
						RecsetND = ObjConn2.Execute(SqlString);
						while( !RecsetND->IsEOF())
						{
							RecsetND->GetFieldValue(L"TenTC",sKTTenTC[slMaTC]);
							break;
						}
						RecsetND->Close();					
					}
					RecsetTC->MoveNext();
				}
				RecsetTC->Close();
			}
		}
	}	

	ObjConn2.CloseDatabase();

	// Sửa 12.06.2017
	val = GIOR(rowac+1,_iCot[2],prDMVL,L"");
	int cong=0;
	if(rowac+1>=iEnd || val==L"HM" || val==L"GD")
	{
		cong=1;
		PRS = psDMVL->GetRange(prDMVL->GetItem(rowac+1,1),prDMVL->GetItem(rowac+cong+1,1));
		PRS->EntireRow->Insert(xlShiftDown);

		PRS = psDMVL->GetRange(prDMVL->GetItem(rowac+1,1),prDMVL->GetItem(rowac+cong+1,1));
		PRS->EntireRow->Interior->PutColor(xlNone);
	}

////////////////////////////////////////////////////////////////////////////

	// Sử dụng bảng tiêu chuẩn hay không? <-- Leo 08.02.2018
	// Nếu sử dụng (jNhomTC=1) --> đưa tất cả tiêu chuẩn sang sheet Bang tieu chuan
	// Tiếp đó đặt tên nhóm và link sang DMCV/VL

		// Lưu tên công việc
		CString sluuTenCV = GIOR(rowac,_iCot[5],prDMVL,L"");

		// Lưu đơn vị tính
		//CString sluuDVTx = GIOR(rowac,_iCot[7],prDMVL,L"");

		// Lưu mã HSNK (nếu có)
		CString sluuMHSNT = GIOR(rowac,_iCot[4],prDMVL,L"");

		// Xóa công tác (17.06.2017)
		int iXoa=0;
		CString sMax=L"",sXoa= GIOR(rowac,_iCot[1],prDMVL,L"");
		if(_wtoi(sXoa)>0)
		{
			for (int i = rowac+1; i < iEnd; i++)
			{
				sXoa = GIOR(i,_iCot[1],prDMVL,L"");
				sMax = GIOR(i,_iCot[2],prDMVL,L"");
				if(_wtoi(sXoa)>0 || sMax==L"HM" || sMax==L"GD")
				{
					iXoa=i;
					break;
				}
			}

			if(iXoa>0)
			{
				if(rowac+1<iXoa)
				{
					PRS = psDMVL->GetRange(prDMVL->GetItem(rowac+1,1),prDMVL->GetItem(iXoa-1,1));
					PRS->EntireRow->Delete(xlShiftUp);
				}

				PRS = prDMVL->GetItem(rowac,1);
				PRS->EntireRow->ClearContents();

				iEnd = FindComment(psDMVL,_iCot[1],"END");
			}
		}
		else
		{
			if(slMaTC>1 && jNhomTC==0)
			{
				// Chèn dòng
				PRS = psDMVL->GetRange(prDMVL->GetItem(rowac+1,1),prDMVL->GetItem(rowac+slMaTC-1,1));
				PRS->EntireRow->Insert(xlShiftDown);

				PRS = psDMVL->GetRange(prDMVL->GetItem(rowac+1,1),prDMVL->GetItem(rowac+slMaTC-1,1));
				PRS->EntireRow->Interior->PutColor(xlNone);
			}
		}		

		// Chèn nội dung tiêu chuẩn
		if(slMaTC>0)
		{
			if(jNhomTC==0)
			{
				for (int i = 0; i < slMaTC; i++)
				{
					prDMVL->PutItem(rowac+i,_iCot[13],(_bstr_t)sKTMaTC[i+1]);
					prDMVL->PutItem(rowac+i,_iCot[6],(_bstr_t)sKTTenTC[i+1]);
					val = prDMVL->GetItem(rowac+i,_iCot[6]+1);
					if(val==L"") prDMVL->PutItem(rowac+i,_iCot[6]+1,(_bstr_t)L" ");
				}
			}
			else
			{
				// Đổ dữ liệu sang sheet Bang tieu chuan
				int icSTT = getColumn(psBTC,L"NHTC_STT");				
				int icMa = getColumn(psBTC,L"NHTC_MA");
				int icTC = getColumn(psBTC,L"NHTC_TC");
				int icGC = getColumn(psBTC,L"NHTC_GC");
				int irSTT = getRow(psBTC,L"NHTC_STT");			
				

				if(vtnhom==0)
				{
					// Kiểm tra xem nhóm tiêu chuẩn đã tồn tại bên sheet BTC chưa?
					int ivtr=1,iOk=0,icheck=0;
					int iEndBTC = prBTC->SpecialCells(xlCellTypeLastCell)->GetRow();
					for (int i = 1; i <= iEndBTC; i++)
					{
						szval = GIOR(i,icTen,prBTC,L"");
						if(szval!=L"")
						{
							szval = GIOR(i,icTC,prBTC,L"");
							if(szval==sKTTenTC[1])
							{
								ivtr=i;
								icheck=1;
								if(slMaTC>1)
								{
									for (int j = 2; j <= slMaTC; j++)
									{
										szval = GIOR(i+j-1,icTC,prBTC,L"");
										if(szval==sKTTenTC[j]) icheck++;
										else break;
									}
								}

								// Đã tồn tại tên nhóm tiêu chuẩn
								if(icheck==slMaTC)
								{							
									iOk=1;
									break;
								}
							}
						}
					}

					if(iOk==1)
					{
						szval.Format(L"='%s'!%s",(LPCTSTR)shBTC,GAR(ivtr,icTen,prBTC,0));
						prDMVL->PutItem(rowac,_iCot[6],(_bstr_t)szval);
					}
					else
					{
						// Chưa tồn tại --> Thêm mới vào đây
						ivtr=irSTT;
						for (int i = iEndBTC; i >= irSTT; i--)
						{
							szval = GIOR(i,icTC,prBTC,L"");
							szval.Trim();
							if(szval!=L"")
							{
								ivtr=i;
								break;
							}
						}

						PRS = psBTC->GetRange(prBTC->GetItem(ivtr+1,1),prBTC->GetItem(ivtr+1,icGC));
						PRS->Interior->PutColor(fRandomRGB());

						m_cbbox.GetWindowTextW(szval);
						szval.Trim();
						if(szval==L"")
						{
							time_t now = time(0);
							tm *ltm = localtime(&now);
							szval.Format(L"%d%d",ltm->tm_min,ltm->tm_sec);
							int _nRand = _wtoi(szval);

							szval.Format(L"Nhóm TC%d%d%d",_nRand,rand()%10,rand()%10);
						}
						prBTC->PutItem(ivtr+1,icTen,(_bstr_t)szval);

						szval.Format(L"=COUNTA(%s:%s)",GAR(irSTT+1,icTen,prBTC,3),GAR(ivtr+1,icTen,prBTC,0));
						prBTC->PutItem(ivtr+1,icSTT,(_bstr_t)szval);

						szval.Format(L"='%s'!%s",(LPCTSTR)shBTC,GAR(ivtr+1,icTen,prBTC,3));
						prDMVL->PutItem(rowac,_iCot[6],(_bstr_t)szval);

						for (int i = 1; i <= slMaTC; i++)
						{
							prBTC->PutItem(ivtr+i,icMa,(_bstr_t)sKTMaTC[i]);
							prBTC->PutItem(ivtr+i,icTC,(_bstr_t)sKTTenTC[i]);
						}
					}
				}
				else
				{
					szval.Format(L"='%s'!%s",(LPCTSTR)shBTC,GAR(vtnhom,icTen,prBTC,0));
					prDMVL->PutItem(rowac,_iCot[6],(_bstr_t)szval);
				}
			}
		}
		else
		{
			if(vtnhom>0)
			{
				szval.Format(L"='%s'!%s",(LPCTSTR)shBTC,GAR(vtnhom,icTen,prBTC,0));
				prDMVL->PutItem(rowac,_iCot[6],(_bstr_t)szval);
			}			
		}

		if(fla==1)
		{
			prDMVL->PutItem(rowac,_iCot[2],(_bstr_t)sMSCV);

			if(sluuTenCV!=L"")
			{
				int iktten = _wtoi(getVCell(psTS,L"CF_TRATEN"));
				if(iktten==1) prDMVL->PutItem(rowac,_iCot[5],(_bstr_t)sNA);
				else prDMVL->PutItem(rowac,_iCot[5],(_bstr_t)sluuTenCV);
			}
			else prDMVL->PutItem(rowac,_iCot[5],(_bstr_t)sNA);

			//if(sluuDVTx!=L"") prDMVL->PutItem(rowac,_iCot[7],(_bstr_t)sluuDVTx);
			//else prDMVL->PutItem(rowac,_iCot[7],(_bstr_t)sDVT);

			// Copy ngày tháng
			int iktrNT=0;
			for (int j = rowac-1; j >= 7; j--)
			{
				val = prDMVL->GetItem(j,1);
				if(_wtoi(val)>0)
				{
					if(iCpy==0 || iCpy==1)
					{
						// ngày
						int gtr[6];
						gtr[0] = getColumn(psDMVL,"DMVL_NK_NG");
						gtr[1] = getColumn(psDMVL,"DMVL_LM_NG");
						gtr[2] = getColumn(psDMVL,"DMVL_LM_KQ");												
						gtr[3] = getColumn(psDMVL,"DMVL_NTNB_NG");
						gtr[4] = getColumn(psDMVL,"DMVL_YC");
						gtr[5] = getColumn(psDMVL,"DMVL_AB_NG");

						for (int k = 0; k < 6; k++)
						{
							szval = GIOR(j,gtr[k],prDMVL,L"");
							f_ktra_date(szval,prDMVL,rowac,gtr[k]);
						}
					}

					if(iCpy==0 || iCpy==2)
					{
						// giờ
						int gtr[3];
						gtr[0] = getColumn(psDMVL,"DMVL_LM_GIO");
						gtr[1] = getColumn(psDMVL,"DMVL_NTNB_GIO");
						gtr[2] = getColumn(psDMVL,"DMVL_AB_GIO");

						for (int k = 0; k < 3; k++)
						{
							szval = GIOR(j,gtr[k],prDMVL,L"");
							prDMVL->PutItem(rowac,gtr[k],(_bstr_t)szval);
						}
					}				

					xl->PutCutCopyMode(XlCutCopyMode(false));
					iktrNT=1;
					break;
				}
			}

			if (iktrNT == 0 || (iCpy != 0 && iCpy != 1))
			{
				// ngày
				int gtr[6];
				gtr[0] = getColumn(psDMVL,"DMVL_NK_NG");
				gtr[1] = getColumn(psDMVL,"DMVL_LM_NG");
				gtr[2] = getColumn(psDMVL,"DMVL_LM_KQ");												
				gtr[3] = getColumn(psDMVL,"DMVL_NTNB_NG");
				gtr[4] = getColumn(psDMVL,"DMVL_YC");
				gtr[5] = getColumn(psDMVL,"DMVL_AB_NG");

				prDMVL->PutItem(rowac,gtr[0],(_bstr_t)L"=NOW()");
				val = GIOR(rowac,gtr[0],prDMVL,L"");
				int pos2= val.Find(L" ");
				if(pos2>0) val = val.Left(pos2);
				val.Trim();

				f_ktra_date(val,prDMVL,rowac,gtr[0]);
				f_ktra_date(val,prDMVL,rowac,gtr[1]);
				f_ktra_date(val,prDMVL,rowac,gtr[2]);
				f_ktra_date(val,prDMVL,rowac,gtr[3]);

				val.Format(L"=%s",GAR(rowac,gtr[2],prDMVL,0));
				prDMVL->PutItem(rowac,gtr[4],(_bstr_t)val);

				val.Format(L"=%s+1",GAR(rowac,gtr[4],prDMVL,0));
				prDMVL->PutItem(rowac,gtr[5],(_bstr_t)val);
			}

			PRS = prDMVL->GetItem(rowac,_iCot[5]);
			PRS->PutWrapText(TRUE);

			PRS = psDMVL->GetRange(prDMVL->GetItem(rowac,1),prDMVL->GetItem(rowac, _iCotEND));
			if(chkClrVL==1) PRS->Interior->PutColor(fclrVL);
			else PRS->EntireRow->Interior->PutColor(xlNone);
		}		

		// Định dạng khung
		// Kiểm tra công tác kế trên có phải hạng mục hay giai đoạn không?
		val = GIOR(rowac-1,_iCot[2],prDMVL,L"");
		val.MakeUpper();
		val.Trim();

		int iktx=slMaTC;
		if(jNhomTC==1 || slMaTC==0) iktx=1; 
		PRS = psDMVL->GetRange(prDMVL->GetItem(rowac,_iCot[1]),prDMVL->GetItem(rowac+iktx+cong, _iCotEND));
		
		if (val == L"HM" || val == L"GD")
		{
			PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->Weight = xlThin;
		}
		else
		{
			PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->PutWeight(xlThin);
			PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->PutLineStyle(xlDot);
		}
		
		PRS->Borders->GetItem(XlBordersIndex::xlEdgeLeft)->Weight=xlThin;
		PRS->Borders->GetItem(XlBordersIndex::xlEdgeRight)->Weight=xlThin;
		PRS->Borders->GetItem(XlBordersIndex::xlInsideVertical)->Weight=xlThin;

		PRS->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->PutWeight(xlThin);
		PRS->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->PutLineStyle(xlDot);

		PRS->Borders->GetItem(XlBordersIndex::xlInsideHorizontal)->PutWeight(xlThin);
		PRS->Borders->GetItem(XlBordersIndex::xlInsideHorizontal)->PutLineStyle(xlDot);


		PRS->Font->PutBold(0);
		PRS->Font->PutItalic(0);

		// Chèn dòng & đổ dữ liệu lấy mẫu
		if(iChekLmau==1)
		{
			PRS = prDMVL->GetItem(rowac+1,1);
			PRS->EntireRow->Insert(xlShiftDown);
			prDMVL->PutItem(rowac+1,_iCot[3],(_bstr_t)L"LM");
			PRS = prDMVL->GetItem(rowac+1,1);
			PRS->EntireRow->Interior->PutColor(xlNone);
		}

		// Bổ sung nhóm ký biên bản (31.05.2018)
		CheckNhomKy(psDMVL,rowac,getColumn(psDMVL,L"DMVL_KBB"));

		// Autofit
		PRS = prDMVL->GetItem(rowac,_iCot[5]);
		PRS->EntireRow->Rows->AutoFit();

		// Sắp xếp lại stt
		iEnd = FindComment(psDMVL,_iCot[1],"END");
		int ntype = _wtoi(getVCell(psTS,L"TS_ATOSTT"));
		if(ntype!=0 && ntype!=1 && ntype!=2) ntype=0;
		f_Auto_stt_dmuc(prDMVL,ntype,0,8,iEnd,_iCot[1],_iCot[2],_iCot[4]);

		if(sluuMHSNT!=L"") prDMVL->PutItem(rowac,_iCot[4],(_bstr_t)sluuMHSNT);
		else
		{
			CString mm = GIOR(rowac,_iCot[1],prDMVL,L"");
			CString qq = mm;
			if(_wtoi(mm)<10) mm.Format(L"'0%s",qq);
			else mm.Format(L"'%s",qq);
			prDMVL->PutItem(rowac,_iCot[4],(_bstr_t)mm);
		}

		// Leo add 07.02.2018
		if(slMaTC>1 && jNhomTC==0)
		{
			xl->ActiveWindow->PutScrollRow(rowac-5);
			PRS = prDMVL->GetItem(rowac+slMaTC-1,_iCot[2]);
			PRS->Select();
		}

		xl->PutScreenUpdating(VARIANT_TRUE);
		CDialog::OnOK();

		AFX_MANAGE_STATE(AfxGetStaticModuleState());
		Dlg_open_progress p;
		p.DoModal();
	}
	catch(...){_s(L"Lỗi đổ dữ liệu vật liệu.");}	
}


void Dlg_Open_DSVL::OnBnClickedB4()
{
	TRY
	{
		ClickEsc=0;

		int nItem=-1;
		int nCount = lit_dmvl.GetItemCount();

		POSITION pos = lit_dmvl.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			for (int i = 0; i < nCount; i++)
			if (lit_dmvl.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
			{
				nItem=i;
				nClick++;
				break;
			}
		}
		if(nItem==-1) nItem = nCount;

		// Kết nối SQL
		ObjConn.OpenAccessDB(sFullPath,L"",L"");
		ObjConn2.OpenAccessDB(pathMDB,L"",L"");
		CRecordset* Recset;

		// Số ngẫu nhiên theo thời gian
		CString val=L"";
		time_t now = time(0);
		tm *ltm = localtime(&now);
		val.Format(L"%d",ltm->tm_sec);
		int _nRand = _wtoi(val);

		// Kiểm tra dữ liệu có trong CSDL không?
		CString sTvan=L"",sDem=L"",sDem2=L"",jNew=L"";
		jNew.Format(L"VL.%d%d%d",_nRand,rand()%10,rand()%10);
		sTvan.Format(L"SELECT COUNT(*) AS iDem FROM tbTuDienVatTu WHERE MSVT LIKE '%s';",jNew);
		Recset = ObjConn.Execute(sTvan);
		Recset->GetFieldValue(L"iDem",sDem);
		Recset->Close();

		sTvan.Format(L"SELECT COUNT(*) AS iDem FROM tbMaVL_tenVL WHERE MaVL LIKE '%s';", jNew);
		Recset = ObjConn2.Execute(sTvan);
		Recset->GetFieldValue(L"iDem", sDem2);
		Recset->Close();

		while (_wtoi(sDem)>0 && _wtoi(sDem2)>0)
		{
			sDem=L"";
			jNew.Format(L"VL.%d%d%d",_nRand,rand()%10,rand()%10);
			sTvan.Format(L"SELECT COUNT(*) AS iDem FROM tbTuDienVatTu WHERE MSVT LIKE '%s';",jNew);
			Recset = ObjConn.Execute(sTvan);
			Recset->GetFieldValue(L"iDem",sDem);
			Recset->Close();

			sTvan.Format(L"SELECT COUNT(*) AS iDem FROM tbMaVL_tenVL WHERE MaVL LIKE '%s';", jNew);
			Recset = ObjConn2.Execute(sTvan);
			Recset->GetFieldValue(L"iDem", sDem2);
			Recset->Close();
		}

		// Đổ dữ liệu vào list
		sDem = L"Nhập tên vật liệu mới";
		lit_dmvl.InsertItem(nItem,L"",0);
		lit_dmvl.SetItemText(nItem,1,jNew);
		lit_dmvl.SetItemText(nItem,2,sDem);
		lit_dmvl.SetRowColors(nItem,RGB(0,128,0),RGB(255,255,255));

		// Lưu vào trong CSDL
		SqlString.Format(L"INSERT INTO tbTuDienVatTu (MSVT,TVT) VALUES ('%s','%s');",jNew,sDem);
		ObjConn.ExecuteRB(SqlString);

		// Lưu vào trong CSDL
		SqlString.Format(L"INSERT INTO tbMaVL_tenVL (MaVL) VALUES ('%s');", jNew);
		ObjConn2.ExecuteRB(SqlString);

		ObjConn.CloseDatabase();
		ObjConn2.CloseDatabase();
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[11] Lỗi thêm mới dữ liệu nội dung vật liệu")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_Open_DSVL::OnBnClickedB5()
{
	TRY
	{
		ClickEsc=0;

		int nCount = lit_dmvl.GetItemCount();
		if(nCount<=0) return;

		POSITION pos=lit_dmvl.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			int dem=0;
			for (int i = 0; i < nCount; i++)
				if (lit_dmvl.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED) dem++;

			CString sTvan,val = L"";
			if(dem==1) val = L"Bạn có chắc chắn sẽ xóa vật liệu này trong dữ liệu?";
			else val = L"Bạn có chắc chắn sẽ xóa những vật liệu này trong dữ liệu?";
			
			int n = _yn(val,0);
			if (n == 6)
			{
				// Kết nối SQL
				ObjConn.OpenAccessDB(sFullPath,L"",L"");
				CRecordset* Recset;
				
				for (int i = nCount-1; i >=0; i--)
				{
					if (lit_dmvl.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
					{
						// Xóa trong CSDL
						val = lit_dmvl.GetItemText(i,1);
						sTvan.Format(L"DELETE FROM tbMaVL_tenVL WHERE MaVL LIKE '%s';",val);
						ObjConn.ExecuteRB(sTvan);

						// Xóa trên list control
						lit_dmvl.DeleteItem(i);
					}
				}

				ObjConn.CloseDatabase();
			}
		}
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[28] Lỗi xóa dữ liệu vật liệu")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_Open_DSVL::OnBnClickedB3()
{
	try
	{
		int count = lit_dmvl.GetItemCount();
		if(count==0) return;

		int ktrchk=-1;
		for (int i = 0; i < count; i++)
		{
			if((int)lit_dmvl.GetCheck(i)==1)
			{
				ktrchk=i;
				break;
			}
		}

		if(ktrchk==-1)
		{
			_s(L"Bạn chưa tích chọn vật liệu cần cập nhật dữ liệu.");
			return;
		}

		if(_yn(L"Bạn chắc chắn cập nhật tiêu chuẩn và nội dung cho các vật liệu được tích chọn?",0)==6)
		{
			ObjConn.OpenAccessDB(sFullPath,L"",L"");
			
			for (int i = ktrchk; i < count; i++)
			{
				if((int)lit_dmvl.GetCheck(i)==1) _UpdateDL(i,0);				
			}

			ObjConn.CloseDatabase();

			_s(L"Cập nhật dữ liệu thành công!");
		}
	}
	catch(...){}
}


void Dlg_Open_DSVL::OnBnClickedB6()
{
	ClickEsc=0;

	f_SaveWidth();
	f_Get_size();
	CDialog::OnCancel();
}


void Dlg_Open_DSVL::OnEnKillfocusEd1()
{
	sBefore = lit_dmvl.GetItemText(CLRow,CLCol);
	sBefore.Replace(L"'",L"");

	txtE_DMVL.GetWindowTextW(sAfter);
	sAfter.Replace(L"'",L"");
	if(sAfter==sBefore) return;
	
	ObjConn.OpenAccessDB(sFullPath, L"", L"");
	ObjConn2.OpenAccessDB(pathMDB, L"", L"");

	if(CLCol==2 || CLCol==3)
	{
		// Kiểm tra mã đó có trong CSDL chưa? 
		// Chưa có thì thêm mới vào
		CString sdem=L"";
		CString sId = lit_dmvl.GetItemText(CLRow,1);
		CRecordset*	Recset;

		if (CLCol == 2)
		{
			SqlString.Format(L"SELECT COUNT(*) AS iDem FROM tbTuDienVatTu WHERE MSVT LIKE '%s';", sId);
			Recset = ObjConn.Execute(SqlString);
			Recset->GetFieldValue(L"iDem", sdem);
			Recset->Close();

			if (_wtoi(sdem) > 0)
			{
				SqlString.Format(L"UPDATE tbTuDienVatTu SET TVT='%s' WHERE MSVT LIKE '%s';", sAfter, sId);
				ObjConn.ExecuteRB(SqlString);
			}
		}

		sdem = L"";
		SqlString.Format(L"SELECT COUNT(*) AS iDem FROM tbMaVL_tenVL WHERE MaVL LIKE '%s';", sId);
		Recset = ObjConn2.Execute(SqlString);
		Recset->GetFieldValue(L"iDem", sdem);
		Recset->Close();

		if (_wtoi(sdem) > 0)
		{
			if (CLCol == 2) sdem = L"TenVL";
			else sdem = L"Kichthuocmau";

			SqlString.Format(L"UPDATE tbMaVL_tenVL SET %s='%s' WHERE MaVL LIKE '%s';", sdem, sAfter, sId);
			ObjConn2.ExecuteRB(SqlString);
		}
	}
	else if(CLCol==1)
	{
		SqlString.Format(L"UPDATE tbTuDienVatTu SET MSVT='%s' WHERE MSVT LIKE '%s';", sAfter, sBefore);
		ObjConn.ExecuteRB(SqlString);

		SqlString.Format(L"UPDATE tbMaVL_tenVL SET MaVL='%s' WHERE MaVL LIKE '%s';",sAfter,sBefore);
		ObjConn2.ExecuteRB(SqlString);
	}

	ObjConn2.CloseDatabase();
	ObjConn.CloseDatabase();
}


int Dlg_Open_DSVL::f_Get_change()
{
	try
	{
		CString val=L"";
		int ktr=0,nCount = lit_tcvl.GetItemCount();
		if(_isl1!=nCount) return 1;
		if(nCount>0)
		{
			for (int i = 0; i < nCount; i++)
			{
				val = lit_tcvl.GetItemText(i,0);
				if(val!=L"") return 1;
			}
		}

		nCount = lit_ppvl.GetItemCount();
		if(_isl2!=nCount) return 1;
		if(nCount>0)
		{
			for (int i = 0; i < nCount; i++)
			{
				val = lit_ppvl.GetItemText(i,0);
				if(val!=L"") return 1;
			}
		}

		return 0;
	}
	catch(...){}

	return 0;
}



void Dlg_Open_DSVL::f_tk_tchuan(CString sVal)
{
	TRY
	{
		// Kết nối SQL
		SqlString.Format(L"SELECT * FROM tbMaVL_Tieuchuan WHERE MaVL LIKE '%s' ORDER BY ID ASC;",sVal);
		CRecordset* Recset = ObjConn2.Execute(SqlString);

		int dem=0;
		CString sGan=L"";
		while( !Recset->IsEOF() )
		{
			Recset->GetFieldValue(L"MaTC",sGan);
			sGan.Trim();
			if(sGan!=L"")
			{
				lit_tcvl.InsertItem(dem,L"",0);
				lit_tcvl.SetItemText(dem,1,sGan);
				lit_tcvl.SetItemText(dem,2,f_ten_tchuan(sGan));
				dem++;
			}		

			Recset->MoveNext();
		}

		Recset->Close();
		_isl1 = dem;
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[8] Lỗi tìm kiếm dữ liệu tiêu chuẩn vật liệu")+e->m_strError);
	}
	END_CATCH;
}


CString Dlg_Open_DSVL::f_ten_tchuan(CString sVal)
{
	TRY
	{
		CString sGan=L"",sTruyvan=L"";
		sTruyvan.Format(L"SELECT * FROM tbTentieuchuan WHERE MaTC LIKE '%s';",sVal);
		CRecordset* Recset = ObjConn2.Execute(sTruyvan);
		while( !Recset->IsEOF() )
		{
			Recset->GetFieldValue(L"TenTC",sGan);
			break;
		}
		Recset->Close();
		
		sGan.Trim();
		if(sGan==L"") sGan = sVal + L": <?> Kiểm tra lại tên tiêu chuẩn";
		return sGan;
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[9] Lỗi tìm kiếm tên tiêu chuẩn")+e->m_strError);
		return L"";
	}
	END_CATCH;

	return L"";
}


void Dlg_Open_DSVL::f_tk_ndung(CString sVal)
{
	TRY
	{
		// Kết nối SQL
		SqlString.Format(L"SELECT * FROM tbMaVL_Quydinhlaymau WHERE MaVL LIKE '%s' ORDER BY ID ASC;",sVal);
		CRecordset* Recset = ObjConn2.Execute(SqlString);
		int dem=0;
		CString sGan=L"";

		while( !Recset->IsEOF() )
		{
			Recset->GetFieldValue(L"MaLayMau",sGan);
			sGan.Trim();
			if(sGan!=L"")
			{
				lit_ppvl.InsertItem(dem,L"",0);
				lit_ppvl.SetItemText(dem,1,sGan);

				f_ten_mauvl(sGan);
				lit_ppvl.SetItemText(dem,2,sMau[0]);	// Phương pháp
				lit_ppvl.SetItemText(dem,3,sMau[1]);	// Tần suất
				lit_ppvl.SetItemText(dem,4,sMau[2]);	// Đơn vị
				dem++;
			}

			Recset->MoveNext();
		}

		Recset->Close();
		_isl2 = dem;
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[8] Lỗi tìm kiếm dữ liệu tiêu chuẩn vật liệu")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_Open_DSVL::f_ten_mauvl(CString sVal)
{
	TRY
	{
		// Kết nối SQL
		CString sTruyvan=L"";
		sMau[0]=L"",sMau[1]=L"",sMau[2]=L"";
		sTruyvan.Format(L"SELECT * FROM tbMau WHERE MaLayMau LIKE '%s';",sVal);
		CRecordset* Recset = ObjConn2.Execute(sTruyvan);
		while( !Recset->IsEOF() )
		{
			Recset->GetFieldValue(L"Phuongphaplaymau",sMau[0]);
			Recset->GetFieldValue(L"TansuatLaymau",sMau[1]);
			Recset->GetFieldValue(L"Donvi",sMau[2]);

			sMau[0].Trim();
			sMau[1].Trim();
			if(sMau[0]==L"") sMau[0] = sVal + L": ?";

			break;
		}
		Recset->Close();
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[19] Lỗi tìm kiếm tên tiêu chuẩn")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_Open_DSVL::OnNMClickL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	ClickEsc=0;
	iTCclick=-1;
	sDownTC=L"",sFullTC=L"";

	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	int nItem = pNMListView->iItem;
	int nSubItem = pNMListView->iSubItem;

	iLuuTCVL = m_chkLuuTC.GetCheck();
	CString szval=L"",szval2=L"",szval3=L"";

	if(nItem>=0)
	{
		if(nSubItem==0)
		{
			int _pos=-1;
			szval = lit_dmvl.GetItemText(nItem,1);
			CString szcong = szval+L"|";

			szval2 = lit_dmvl.GetItemText(nItem,2);
			szval2.Replace(L"|",L" ");
			if(szval2==L"") szval2=L"Vật tư chưa đặt tên";

			//szval3 = lit_dmvl.GetItemText(nItem,3);
			//CString szcong3 = szval3+L"|";

			if(sLuuMaCT!=L"")
			{
				 if(sLuuMaCT.Right(1)!=L"|") sLuuMaCT+=L"|";
				 if(sLuuTenCT.Right(1)!=L"|") sLuuTenCT+=L"|";
				 //if(sLuuDVT.Right(1)!=L"|") sLuuDVT+=L"|";
				 //if(sLuuLink.Right(1)!=L"|") sLuuLink+=L"|";			 

				_pos = sLuuMaCT.Find(szcong);
			}

			if((int)lit_dmvl.GetCheck(nItem)==0)
			{
				// Tích chọn
				lit_dmvl.SetRowColors(nItem,RGB(255,255,150),BLACK);

				if(_pos==-1 && iLuuTCVL==1)
				{
					sLuuMaCT+=szval;
					m_txtTChon.SetWindowText(sLuuMaCT);

					sLuuTenCT+=szval2;
					//sLuuDVT+=szval3;
					//sLuuLink+=sFullPath;
				}								
			}
			else
			{
				// Bỏ tích
				lit_dmvl.SetRowColors(nItem,WHITE,BLACK);

				if(_pos>=0 && iLuuTCVL==1)
				{
					int demT=0;
					CString szgan = sLuuMaCT;
					if(szgan.Right(1)!=L"|") szgan+=L"|";
					for (int i = _pos; i >=0; i--)
					{
						if(szgan.Mid(i,1)==L"|") demT++;
					}

					if(demT>0)
					{
						szgan = sLuuMaCT;
						if(szgan.Right(1)!=L"|") szgan+=L"|";

						int demS=0,ilenf=0;
						CString szgg[3]={L"",sLuuMaCT,sLuuTenCT};
						//CString szgg[5]={L"",sLuuMaCT,sLuuTenCT,sLuuDVT,sLuuLink};
						for (int a = 1; a <= 2; a++)
						{
							demS=0;							
							szgan = szgg[a];
							ilenf = szgan.GetLength();
							if(szgan.Right(1)!=L"|") szgan+=L"|";
							for (int i = 0; i < ilenf; i++)
							{
								if(szgan.Mid(i,1)==L"|")
								{
									demS++;
									if(demS==demT)
									{
										for (int j = i+1; j < ilenf; j++)
										{
											if(szgan.Mid(j,1)==L"|") 
											{
												szgg[a].Format(L"%s|%s",szgan.Left(i),szgan.Right(ilenf-j-1));
												break;
											}
										}

										break;
									}
								}
							}
						}

						sLuuMaCT = szgg[1];
						sLuuTenCT = szgg[2];
						//sLuuDVT = szgg[3];
						//sLuuLink = szgg[4];

						/*if(sLuuMaCT.Right(1)==L"|") sLuuMaCT = sLuuMaCT.Left(sLuuMaCT.GetLength()-1);
						if(sLuuTenCT.Right(1)==L"|") sLuuTenCT = sLuuTenCT.Left(sLuuTenCT.GetLength()-1);
						if(sLuuDVT.Right(1)==L"|") sLuuDVT = sLuuDVT.Left(sLuuDVT.GetLength()-1);
						if(sLuuLink.Right(1)==L"|") sLuuLink = sLuuLink.Left(sLuuLink.GetLength()-1);*/

						m_txtTChon.SetWindowText(sLuuMaCT);
					}
					else
					{
						sLuuMaCT= L"";
						sLuuTenCT= L"";
						//sLuuDVT= L"";
						//sLuuLink= L"";
					}
				}
			}

			return;
		}

		// Kiểm tra vị trí click có trùng không?
		if(nItem==nClick) return;

		// Kiểm tra nội dung & pp nghiệm thu có thay đổi không?
		if(f_Get_change()==1)
		{
			CString val = L"Dữ liệu chỉnh sửa nội dung & phương pháp nghiệm thu chưa được lưu."
				L"\nBạn vẫn muốn xem công việc khác mà không lưu?";
			int n0 = _yn(val,0);
			if(n0==7)
			{
				lit_dmvl.SetItemState(nItem,0,LVIS_SELECTED);
				lit_dmvl.SetItemState(nClick,LVIS_SELECTED,LVIS_SELECTED);
				return;
			}
		}
		
		// Xóa toàn bộ dữ liệu trên List Control
		lit_tcvl.DeleteAllItems();
		ASSERT(lit_tcvl.GetItemCount() == 0);

		lit_ppvl.DeleteAllItems();
		ASSERT(lit_ppvl.GetItemCount() == 0);

		nClick = nItem;
		CString sMavl = lit_dmvl.GetItemText(nClick,1);

		ObjConn2.OpenAccessDB(pathMDB,L"",L"");

		// Kích thước mẫu
		SqlString.Format(L"SELECT * FROM tbMaVL_tenVL WHERE MaVL LIKE '%s';",sMavl);
		CRecordset* Recset = ObjConn2.Execute(SqlString);
		while( !Recset->IsEOF() )
		{
			Recset->GetFieldValue(L"Kichthuocmau",szval);
			break;
		}
		Recset->Close();
		lit_dmvl.SetItemText(nClick,3,szval);

		// Tìm kiếm tiêu chuẩn
		f_tk_tchuan(sMavl);

		// Tìm kiếm nội dung vật liệu
		f_tk_ndung(sMavl);

		ObjConn2.CloseDatabase();
	}

	*pResult = 0;
}


void Dlg_Open_DSVL::OnBnClickedB7()
{
	try
	{
		ClickEsc=0;
		int nCount=0,nItem=-1;
		if(_iTab==0) return;
		else if(_iTab==1)
		{
			// Thêm mới tiêu chuẩn
			nCount = lit_tcvl.GetItemCount();
			POSITION pos=lit_tcvl.GetFirstSelectedItemPosition();
			if(pos!=NULL)
			{
				for (int i = 0; i < nCount; i++)
				if (lit_tcvl.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					nItem=i;
					break;
				}
			}

			if(nItem==-1) nItem = nCount;

			// Đổ dữ liệu vào list
			lit_tcvl.InsertItem(nItem,L"",0);
			lit_tcvl.SetItemText(nItem,2,L"Gõ tìm kiếm tiêu chuẩn");
			lit_tcvl.SetRowColors(nItem,RGB(0,128,0),RGB(255,255,255));
		}
		else if(_iTab==2)
		{
			// Thêm mới phương pháp
			nCount = lit_ppvl.GetItemCount();
			POSITION pos=lit_ppvl.GetFirstSelectedItemPosition();
			if(pos!=NULL)
			{
				for (int i = 0; i < nCount; i++)
				if (lit_ppvl.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					nItem=i;
					break;
				}
			}

			if(nItem==-1) nItem = nCount;

			// Đổ dữ liệu vào list
			lit_ppvl.InsertItem(nItem,L"",0);
			lit_ppvl.SetItemText(nItem,2,L"Gõ tìm kiếm phương pháp lấy mẫu");
			lit_ppvl.SetRowColors(nItem,RGB(0,128,0),RGB(255,255,255));
		}
	}
	catch(...){_s(L"Lỗi thêm mới nội dung & phương pháp lấy mẫu.");}
}


void Dlg_Open_DSVL::OnBnClickedB8()
{
	try
	{
		ClickEsc=0;
		int nCount=0,nItem=-1;
		CString sMsg = L"Bạn có chắc chắn xóa dữ liệu hiển thị?"
			L"\n(Dữ liệu này sẽ vẫn tồn tại trong CSDL nếu bạn xóa)";

		if(_iTab==0) return;
		else if(_iTab==1)
		{
			// Xóa tiêu chuẩn
			nCount = lit_tcvl.GetItemCount();
			POSITION pos=lit_tcvl.GetFirstSelectedItemPosition();
			if(pos!=NULL)
			{
				if(_yn(sMsg,0)==6)
				{
					for (int i = nCount-1; i >= 0; i--)
						if (lit_tcvl.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED) lit_tcvl.DeleteItem(i);
				}				
			}
		}
		else if(_iTab==2)
		{
			// Xóa phương pháp
			nCount = lit_ppvl.GetItemCount();
			POSITION pos=lit_ppvl.GetFirstSelectedItemPosition();
			if(pos!=NULL)
			{
				if(_yn(sMsg,0)==6)
				{
					for (int i = nCount-1; i >= 0; i--)
						if (lit_ppvl.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED) lit_ppvl.DeleteItem(i);
				}
			}
		}
	}
	catch(...){_s(L"Lỗi xóa dữ liệu nội dung & phương pháp lấy mẫu.");}
}


void Dlg_Open_DSVL::OnBnClickedB9()
{
	ClickEsc=0;

	// Kiểm tra nội dung & pp nghiệm thu có thay đổi không?
	if(f_Get_change()==1)
	{
		// Kết nối SQL
		ObjConn.OpenAccessDB(sFullPath,L"",L"");
		ObjConn2.OpenAccessDB(pathMDB,L"",L"");

		_UpdateDL(nClick,1);		

		ObjConn2.CloseDatabase();
		ObjConn.CloseDatabase();
	}
	else _s(L"Không có dữ liệu cập nhật.");
}


void Dlg_Open_DSVL::_UpdateDL(int iItemR, int itype)
{
	try
	{
		CRecordset* Recset;
		CString sKtra=L"",sDem=L"",sTvan=L"",sGan[3];
		CString val = lit_dmvl.GetItemText(iItemR,1);

		// Tạo mã vật liệu mới
		int iRes = 0;
		if(itype==1)
		{
			sKtra=L"Mã vật liệu đã tồn tại. Bạn có muốn tự động mã mới?"
			L"\nChọn 'Yes' để tạo mã mới hoặc 'No' để thay thế dữ liệu cũ. \nHoặc nhấn Cancel để hủy cập nhật.";
			iRes = _yn(sKtra,0);
		}
		else iRes=7;

		if(iRes==6)
		{
			// Số ngẫu nhiên theo thời gian
			CString sRd=L"";
			time_t now = time(0);
			tm *ltm = localtime(&now);
			sRd.Format(L"%d",ltm->tm_sec);
			int _nRand = _wtoi(sRd);

			CString jNew=L"",sGn=val;
			if(val.GetLength()>5) sGn = val.Left(5);

			// Kiểm tra dữ liệu có trong CSDL không?
			jNew.Format(L"%s.%d%d",sGn,_nRand,rand()%10);
			sTvan.Format(L"SELECT COUNT(*) AS iDem FROM tbTuDienVatTu WHERE MSVT LIKE '%s';",jNew);
			Recset = ObjConn.Execute(sTvan);
			Recset->GetFieldValue(L"iDem",sDem);
			Recset->Close();

			while (_wtoi(sDem)>0)
			{
				sDem=L"";
				jNew.Format(L"%s.%d%d%d",sGn,_nRand,rand()%10,rand()%10);
				sTvan.Format(L"SELECT COUNT(*) AS iDem FROM tbTuDienVatTu WHERE MSVT LIKE '%s';",jNew);
				Recset = ObjConn.Execute(sTvan);
				Recset->GetFieldValue(L"iDem",sDem);
				Recset->Close();
			}

			val = jNew;

			// Gán lại mã công việc
			sGan[1] = lit_dmvl.GetItemText(iItemR,2);
			sGan[2] = lit_dmvl.GetItemText(iItemR,3);
				
			sTvan.Format(L"INSERT INTO tbTuDienVatTu (MSVT,TVT) VALUES ('%s','%s');",val,sGan[1]);
			ObjConn.ExecuteRB(sTvan);

			sTvan.Format(L"SELECT COUNT(*) AS iDem FROM tbMaVL_tenVL WHERE MaVL LIKE '%s';",val);
			Recset = ObjConn.Execute(sTvan);
			Recset->GetFieldValue(L"iDem",sDem);
			Recset->Close();
			if(_wtoi(sDem)>0) SqlString.Format(L"UPDATE tbMaVL_tenVL SET Kichthuocmau='%s' WHERE MaVL LIKE '%s';",sGan[2],val);
			else sTvan.Format(L"INSERT INTO tbMaVL_tenVL (MaVL,Kichthuocmau) VALUES ('%s','%s');",val,sGan[2]);
			ObjConn2.ExecuteRB(sTvan);

			// Cập nhật trạng thái trong list
			lit_dmvl.InsertItem(iItemR+1,L"",0);
			lit_dmvl.SetItemText(iItemR+1,1,val);
			lit_dmvl.SetItemText(iItemR+1,2,sGan[1]);
			lit_dmvl.SetItemText(iItemR+1,3,sGan[2]);
			lit_dmvl.SetRowColors(iItemR+1,RGB(0,128,0),RGB(255,255,255));
		}
		else if(iRes==7)
		{
			// Đã tồn tại --> Xóa các dữ liệu cũ (tiêu chuẩn & nội dung)
			// Xóa các dữ liệu tiêu chuẩn cũ
			sTvan.Format(L"DELETE FROM tbMaVL_Tieuchuan WHERE MaVL LIKE '%s';",val);
			ObjConn2.ExecuteRB(sTvan);

			// Xóa các dữ liệu phương pháp nghiệm thu cũ
			sTvan.Format(L"DELETE FROM tbMaVL_Quydinhlaymau WHERE MaVL LIKE '%s';",val);
			ObjConn2.ExecuteRB(sTvan);
		}
		else return;

		// Thêm mới dữ liệu tiêu chuẩn
		_isl1 = lit_tcvl.GetItemCount();
		for (int i = 0; i < _isl1; i++)
		{
			// Lưu dữ liệu
			if(i<9) sDem.Format(L"0%d",i+1);
			else sDem.Format(L"%d",i+1);
			sGan[1] = lit_tcvl.GetItemText(i,1);
			sTvan.Format(L"INSERT INTO tbMaVL_Tieuchuan (ID,MaVL,MaTC) VALUES ('%s','%s','%s');",sDem,val,sGan[1]);
			ObjConn2.ExecuteRB(sTvan);

			// Cập nhật trạng thái trong list
			lit_tcvl.SetItemText(i,0,L"");
		}
			
		// Thêm mới dữ liệu nội dung
		_isl2 = lit_ppvl.GetItemCount();
		for (int i = 0; i < _isl2; i++)
		{
			// Lưu dữ liệu
			if(i<9) sDem.Format(L"0%d",i+1);
			else sDem.Format(L"%d",i+1);
			sGan[1] = lit_ppvl.GetItemText(i,1);
			sTvan.Format(L"INSERT INTO tbMaVL_Quydinhlaymau (ID,MaVL,MaLayMau) VALUES ('%s','%s','%s');",sDem,val,sGan[1]);
			ObjConn2.ExecuteRB(sTvan);

			// Cập nhật trạng thái trong list
			lit_ppvl.SetItemText(i,0,L"");
		}
	}
	catch(...){}
}


void Dlg_Open_DSVL::OnBnClickedB10()
{
	ClickEsc=0;

	if(_iTab==0) return;
	else if(_iTab==1) f_UpDown_listTC(-1);
	else f_UpDown_listND(-1);
}


void Dlg_Open_DSVL::OnBnClickedB11()
{
	ClickEsc=0;

	if(_iTab==0) return;
	else if(_iTab==1) f_UpDown_listTC(1);
	else f_UpDown_listND(1);
}


void Dlg_Open_DSVL::f_UpDown_listTC(int iStyle)
{
	try
	{
		int nCount=0,nItem=-1,nCol=4;
		nCount = lit_tcvl.GetItemCount();
		POSITION pos=lit_tcvl.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			for (int i = 0; i < nCount; i++)
				if (lit_tcvl.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					nItem = i;
					break;
				}

			// Kiểm tra điều kiện
			if((iStyle==-1 && nItem<=0) || (iStyle==1 && nItem>=nCount-1)) return;

			// Hoán đổi dữ liệu
			CString sGan=L"";
			for (int j = 0; j < nCol; j++)
			{
				sGan = lit_tcvl.GetItemText(nItem,j);
				lit_tcvl.SetItemText(nItem,j,lit_tcvl.GetItemText(nItem+iStyle,j));
				lit_tcvl.SetItemText(nItem+iStyle,j,sGan);
			}

			lit_tcvl.SetItemText(nItem+iStyle,0,L"1");
			lit_tcvl.SetRowColors(nItem,RGB(255,255,255),RGB(0,0,0));
			lit_tcvl.SetRowColors(nItem+iStyle,RGB(255,255,255),RGB(0,0,0));
			lit_tcvl.SetItemState(nItem,0,LVIS_SELECTED);
			lit_tcvl.SetItemState(nItem+iStyle,LVIS_SELECTED,LVIS_SELECTED);
		}
	}
	catch(...){_s(L"Lỗi hoán đổi vị trí tiêu chuẩn.");}
}


void Dlg_Open_DSVL::f_UpDown_listND(int iStyle)
{
	try
	{
		int nCount=0,nItem=-1,nCol=5;
		nCount = lit_ppvl.GetItemCount();
		POSITION pos=lit_ppvl.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			for (int i = 0; i < nCount; i++)
				if (lit_ppvl.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					nItem = i;
					break;
				}

			// Kiểm tra điều kiện
			if((iStyle==-1 && nItem<=0) || (iStyle==1 && nItem>=nCount-1)) return;

			// Hoán đổi dữ liệu
			CString sGan=L"";
			for (int j = 0; j < nCol; j++)
			{
				sGan = lit_ppvl.GetItemText(nItem,j);
				lit_ppvl.SetItemText(nItem,j,lit_ppvl.GetItemText(nItem+iStyle,j));
				lit_ppvl.SetItemText(nItem+iStyle,j,sGan);
			}

			lit_ppvl.SetItemText(nItem+iStyle,0,L"1");
			lit_ppvl.SetRowColors(nItem,RGB(255,255,255),RGB(0,0,0));
			lit_ppvl.SetRowColors(nItem+iStyle,RGB(255,255,255),RGB(0,0,0));
			lit_ppvl.SetItemState(nItem,0,LVIS_SELECTED);
			lit_ppvl.SetItemState(nItem+iStyle,LVIS_SELECTED,LVIS_SELECTED);
		}
	}
	catch(...){_s(L"Lỗi hoán đổi vị trí tiêu chuẩn.");}
}

void Dlg_Open_DSVL::OnLvnKeydownL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);
	if(pLVKeyDow->wVKey == VK_SPACE) OnBnClickedB2();
	else if(pLVKeyDow->wVKey == VK_DELETE) OnBnClickedB5();

	*pResult = 0;
}


void Dlg_Open_DSVL::OnLvnKeydownL2(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);
	if(pLVKeyDow->wVKey == VK_DELETE) OnBnClickedB8();

	*pResult = 0;
}


void Dlg_Open_DSVL::OnLvnKeydownL3(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);
	if(pLVKeyDow->wVKey == VK_DELETE) OnBnClickedB8();

	*pResult = 0;
}


void Dlg_Open_DSVL::OnNMClickL2(NMHDR *pNMHDR, LRESULT *pResult)
{
	ClickEsc=0;
	_iTab=1;

	sDownTC=L"",sFullTC=L"";
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	iTCclick = pNMListView->iItem;
	if(iTCclick>=0)
	{
		sDownTC = lit_tcvl.GetItemText(iTCclick,1);
		sFullTC = lit_tcvl.GetItemText(iTCclick,2);
	}
}


void Dlg_Open_DSVL::OnNMClickL3(NMHDR *pNMHDR, LRESULT *pResult)
{
	ClickEsc=0;
	_iTab=2;
}


void Dlg_Open_DSVL::OnEnKillfocusEd2()
{
	CString sEdit = L"";
	txtE_TCVL.GetWindowTextW(sEdit);

	CString sKtra = lit_tcvl.GetItemText(CLRow,CLCol);
	if(sEdit != sKtra)
	{
		// Thông tin đã được chỉnh sửa
		lit_tcvl.SetItemText(CLRow,0,L"1");
	}
}


void Dlg_Open_DSVL::OnEnKillfocusEd3()
{
	CString sEdit = L"";
	txtE_PPVL.GetWindowTextW(sEdit);

	CString sKtra = lit_ppvl.GetItemText(CLRow,CLCol);
	if(sEdit != sKtra)
	{
		// Thông tin đã được chỉnh sửa
		lit_ppvl.SetItemText(CLRow,0,L"1");
	}
}

void Dlg_Open_DSVL::OnEnSetfocusT1()
{
	ClickEsc=0;
}


void Dlg_Open_DSVL::OnBnClickedCbs1()
{
	ClickEsc=0;
}


void Dlg_Open_DSVL::OnBnClickedCbs2()
{
	ClickEsc=0;
}


void Dlg_Open_DSVL::OnEnSetfocusEd1()
{
	ClickEsc=1;
}


void Dlg_Open_DSVL::OnEnSetfocusEd2()
{
	ClickEsc=1;
}


void Dlg_Open_DSVL::OnEnSetfocusEd3()
{
	ClickEsc=1;
}


void Dlg_Open_DSVL::OnBnClickedCbs3()
{
	chkClrVL = m_chkColor.GetCheck();
	m_btnColor.EnableWindow(chkClrVL);
	m_txtColor.ShowWindow(chkClrVL);
}


void Dlg_Open_DSVL::OnBnClickedB12()
{
	CColorDialog dlg; 
	if (dlg.DoModal() == IDOK) 
	{ 
		fclrVL = dlg.GetColor(); 
		m_txtColor.SetBkColor(chkClrVL);
		m_txtColor.SetTextColor(BLACK);
	}
}


void Dlg_Open_DSVL::OnSysCommand(UINT nID, LPARAM lParam)
{
	if(nID == SC_CLOSE) OnBnClickedB6();
	else CDialog::OnSysCommand(nID, lParam);
}


void Dlg_Open_DSVL::OnBnClickedCbs4()
{
	int n = m_chkBTC.GetCheck();
	m_cbbox.EnableWindow(n);
	if(n==1) GotoDlgCtrl(GetDlgItem(DSVL_CBB1));
}


void Dlg_Open_DSVL::OnContextMenu(CWnd* /*pWnd*/, CPoint point)
{
	try
	{
		if(iTCclick==-1) return;

		CMenu mnu2;
		mnu2.LoadMenu(IDR_MENU6);
		
		CListCtrlEx *pClist;
		pClist = reinterpret_cast<CListCtrlEx *>(GetDlgItem(DSVL_L2));

		CRect rectSubmitButton;
		pClist->GetWindowRect(&rectSubmitButton);

		CMenu *mnuP2 = mnu2.GetSubMenu(0);
		ASSERT(mnuP2);

		if( rectSubmitButton.PtInRect(point) )
			mnuP2->TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON, point.x, point.y, this);
	}
	catch(...){}
}


void Dlg_Open_DSVL::OnNMRClickL2(NMHDR *pNMHDR, LRESULT *pResult)
{
	ClickEsc=0;
	sDownTC=L"",sFullTC=L"";
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	iTCclick = pNMListView->iItem;
	if(iTCclick>=0)
	{
		sDownTC = lit_tcvl.GetItemText(iTCclick,1);
		sFullTC = lit_tcvl.GetItemText(iTCclick,2);
	}
}


CString Dlg_Open_DSVL::_GetMaTieuchuan()
{
	try
	{
		iTCclick=-1;
		CString szval=L"";
		POSITION pos = lit_tcvl.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			for (int i = 0; i < lit_tcvl.GetItemCount(); i++)
			if (lit_tcvl.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
			{
				iTCclick=i;
				szval = lit_tcvl.GetItemText(i,1);
				return szval;
			}
		}		
	}
	catch(...){return L"";}
	return L"";
}


void Dlg_Open_DSVL::OnTi40039()
{
	sDownTC = _GetMaTieuchuan();
	f_DocFileTC(sDownTC, 1);
}


void Dlg_Open_DSVL::OnTi40040()
{
	sDownTC = _GetMaTieuchuan();
	f_SearchGoogle(L"",sDownTC,0);
}


void Dlg_Open_DSVL::OnTi40052()
{
	try
	{
		POSITION pos = lit_tcvl.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			for (int i = 0; i < lit_tcvl.GetItemCount(); i++)
			if (lit_tcvl.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
			{
				CLRow = i;
				qThemTC=0;
				jTTTC=1;
				_pcaltc=1;
				xl_tukhoa = L"";
				Dlg_all_tieuchuan _dlg;
				_dlg.f_Tim_kiem_TC();
				return;
			}
		}
	}
	catch(...){}
}


void Dlg_Open_DSVL::OnNMDblclkL2(NMHDR *pNMHDR, LRESULT *pResult)
{
	ClickEsc=0;
	_iTab=1;

	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	int nItem = pNMListView->iItem;
	int nSubItem = pNMListView->iSubItem;

	if(nItem==-1 || nSubItem==-1) OnBnClickedB7();
}


void Dlg_Open_DSVL::OnNMDblclkL3(NMHDR *pNMHDR, LRESULT *pResult)
{
	ClickEsc=0;
	_iTab=2;

	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	int nItem = pNMListView->iItem;
	int nSubItem = pNMListView->iSubItem;

	if(nItem==-1 || nSubItem==-1) OnBnClickedB7();
}

void Dlg_Open_DSVL::OnBnClickedPre()
{
	int num = m_cbbFile.GetCurSel();
	num--;

	if(num<0) num = slfileDL-1;
	m_cbbFile.SetCurSel(num);

	//if(sFiledulieu[num+1].Find(L"\\")>0) sFullPath = sFiledulieu[num+1];
	//else sFullPath.Format(L"%s\\%s",sDDanfile,sFiledulieu[num+1]);

	sFullPath = sFpathdulieu[num+1];
	this->SetWindowTextW(L"Danh sách vật liệu | " + sFullPath);	

	OnBnClickedB1();
}


void Dlg_Open_DSVL::OnBnClickedNxt()
{
	int num = m_cbbFile.GetCurSel();
	num++;

	if(num>=slfileDL) num=0;
	m_cbbFile.SetCurSel(num);

	//if(sFiledulieu[num+1].Find(L"\\")>0) sFullPath = sFiledulieu[num+1];
	//else sFullPath.Format(L"%s\\%s",sDDanfile,sFiledulieu[num+1]);

	sFullPath = sFpathdulieu[num+1];
	this->SetWindowTextW(L"Danh sách vật liệu | " + sFullPath);

	OnBnClickedB1();
}


void Dlg_Open_DSVL::OnCbnSelchangeCbbfl()
{
	int num = m_cbbFile.GetCurSel();
	if(num==-1) return;	
	
	//if(sFiledulieu[num+1].Find(L"\\")>0) sFullPath = sFiledulieu[num+1];
	//else sFullPath.Format(L"%s\\%s",sDDanfile,sFiledulieu[num+1]);

	sFullPath = sFpathdulieu[num+1];
	this->SetWindowTextW(L"Danh sách vật liệu | " + sFullPath);

	OnBnClickedB1();
}


void Dlg_Open_DSVL::OnBnClickedCkluu()
{
	sLuuMaCT=L"",sLuuTenCT=L"",sLuuDVT=L"",sLuuLink=L"";
	iLuuTCVL = m_chkLuuTC.GetCheck();
	m_txtTChon.EnableWindow(iLuuTCVL);

	if(iLuuTCVL==1)
	{
		int ncount = lit_dmvl.GetItemCount();
		if(ncount>0)
		{
			CString szval=L"";
			for (int i = 0; i < ncount; i++)
			{
				if((int)lit_dmvl.GetCheck(i)==1)
				{
					szval = lit_dmvl.GetItemText(i,1);
					sLuuMaCT+=szval;
					sLuuMaCT+=L"|";

					szval = lit_dmvl.GetItemText(i,2);
					sLuuTenCT+=szval;
					sLuuTenCT+=L"|";

					//szval = lit_dmvl.GetItemText(i,3);
					//sLuuDVT+=szval;
					//sLuuDVT+=L"|";

					//sLuuLink+=sFullPath;
					//sLuuLink+=L"|";
				}				
			}

			if(sLuuMaCT.Right(1)==L"|") sLuuMaCT = sLuuMaCT.Left(sLuuMaCT.GetLength()-1);			
			if(sLuuTenCT.Right(1)==L"|") sLuuTenCT = sLuuTenCT.Left(sLuuTenCT.GetLength()-1);
			//if(sLuuDVT.Right(1)==L"|") sLuuDVT = sLuuDVT.Left(sLuuDVT.GetLength()-1);
			//if(sLuuLink.Right(1)==L"|") sLuuLink = sLuuLink.Left(sLuuLink.GetLength()-1);

			m_txtTChon.SetWindowText(sLuuMaCT);
		}			
	}
	else m_txtTChon.SetWindowText(sLuuMaCT);
}


void Dlg_Open_DSVL::OnLvnEndScrollL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	try
	{
		if(iStopload==0) return;

		RECT r;
		CRect rectCtrl;
		lit_dmvl.GetItemRect(0, &r, LVIR_BOUNDS);
		lit_dmvl.GetWindowRect(&rectCtrl);
		int a = r.bottom - r.top;	
		int b = rectCtrl.Height();
		int topIndex = lit_dmvl.GetTopIndex();
		int ncount = lit_dmvl.GetItemCount();
		if(topIndex+(int)(b/a) >= ncount)
		{
			lanshow++;
			int iz=ibuocnhay*lanshow;
			if(tslkq<=iz)
			{
				iz=tslkq;
				iStopload=0;
			}
			else iStopload=1;

			for (int i = ibuocnhay*(lanshow-1); i < iz; i++)
			{
				lit_dmvl.InsertItem(i,XLXD[i+1].CDG[0],0);
				lit_dmvl.SetItemText(i,1,XLXD[i+1].CDG[0]);
				lit_dmvl.SetItemText(i,2,XLXD[i+1].CDG[1]);
				lit_dmvl.SetItemText(i,3,XLXD[i+1].CDG[2]);
			}

			lit_dmvl.EnsureVisible(ncount+(int)(b/a)-5, FALSE);
		}
	}
	catch(...){}
}
