#include "Dlg_Open_DMCV.h"
#include "Dlg_all_tieuchuan.h"
#include "Dlg_trauu_tieuchuan.h"
#include "Dlg_open_progress.h"

Dlg_Open_DMCV::Dlg_Open_DMCV(CWnd* pParent /*=NULL*/)
	: CDialog(Dlg_Open_DMCV::IDD, pParent)
{
	
}

void Dlg_Open_DMCV::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, DSCV_T1, m_Txt);
	DDX_Control(pDX, DSCV_L1, lit_dmcv);
	DDX_Control(pDX, DSCV_L2, lit_tccv);
	DDX_Control(pDX, DSCV_L3, lit_ndkt);
	DDX_Control(pDX, DSCV_ED1, txtE_DMCV);
	DDX_Control(pDX, DSCV_ED2, txtE_TCCV);
	DDX_Control(pDX, DSCV_ED3, txtE_NDCV);
	DDX_Control(pDX, DSCV_CK2, m_btn[1]);
	DDX_Control(pDX, DSCV_CK1, m_btn[0]);
	DDX_Control(pDX, DSCV_S2, m_Group);
	DDX_Control(pDX, DSCV_CBS1, m_cbk2);
	DDX_Control(pDX, DSCV_CBS2, m_cbk1);
	DDX_Control(pDX, DSCV_CBS3, m_cbk3);

	DDX_Control(pDX, DSCV_B1, bxx[0]);
	DDX_Control(pDX, DSCV_B3, bxx[1]);
	DDX_Control(pDX, DSCV_B4, bxx[2]);
	DDX_Control(pDX, DSCV_B5, bxx[3]);
	DDX_Control(pDX, DSCV_B7, bxx[4]);
	DDX_Control(pDX, DSCV_B8, bxx[5]);
	DDX_Control(pDX, DSCV_B9, bxx[6]);
	DDX_Control(pDX, DSCV_B10, bxx[7]);
	DDX_Control(pDX, DSCV_B11, bxx[8]);

	
	DDX_Control(pDX, DSCV_CKDMCVCON, m_chkDMCVCon);
	DDX_Control(pDX, DSCV_CKLUU, m_chkLuuTC);
	DDX_Control(pDX, DSCV_T2, m_txtTChon);

	DDX_Control(pDX, DSCV_CBS4, m_chkColor);
	DDX_Control(pDX, DSCV_B13, m_btnColor);
	DDX_Control(pDX, DSCV_S10, m_txtColor);

	DDX_Control(pDX, DSCV_CBS5, m_chkBTC);
	DDX_Control(pDX, DSCV_CBB1, m_cbbox);

	DDX_Control(pDX, DSCV_PRE, m_btnPre);
	DDX_Control(pDX, DSCV_NXT, m_btnNxt);
	DDX_Control(pDX, DSCV_CBBFL, m_cbbFile);	
}

BEGIN_MESSAGE_MAP(Dlg_Open_DMCV, CDialog)
	ON_WM_SIZE()
	ON_WM_SIZING()
	ON_WM_SYSCOMMAND()
	ON_WM_CONTEXTMENU()
	ON_BN_CLICKED(DSCV_B1, &Dlg_Open_DMCV::OnBnClickedB1)
	ON_BN_CLICKED(DSCV_B3, &Dlg_Open_DMCV::OnBnClickedB3)
	ON_BN_CLICKED(DSCV_B2, &Dlg_Open_DMCV::OnBnClickedB2)
	ON_NOTIFY(NM_CLICK, DSCV_L1, &Dlg_Open_DMCV::OnNMClickL1)
	ON_EN_KILLFOCUS(DSCV_ED1, &Dlg_Open_DMCV::OnEnKillfocusEd1)
	ON_BN_CLICKED(DSCV_B4, &Dlg_Open_DMCV::OnBnClickedB4)
	ON_BN_CLICKED(DSCV_B5, &Dlg_Open_DMCV::OnBnClickedB5)
	ON_BN_CLICKED(DSCV_B6, &Dlg_Open_DMCV::OnBnClickedB6)
	ON_NOTIFY(NM_CLICK, DSCV_L2, &Dlg_Open_DMCV::OnNMClickL2)
	ON_NOTIFY(NM_CLICK, DSCV_L3, &Dlg_Open_DMCV::OnNMClickL3)
	ON_BN_CLICKED(DSCV_B7, &Dlg_Open_DMCV::OnBnClickedB7)
	ON_BN_CLICKED(DSCV_B8, &Dlg_Open_DMCV::OnBnClickedB8)
	ON_BN_CLICKED(DSCV_B9, &Dlg_Open_DMCV::OnBnClickedB9)
	ON_BN_CLICKED(DSCV_B10, &Dlg_Open_DMCV::OnBnClickedB10)
	ON_BN_CLICKED(DSCV_B11, &Dlg_Open_DMCV::OnBnClickedB11)
	ON_BN_CLICKED(DSCV_CK1, &Dlg_Open_DMCV::OnBnClickedCk1)
	ON_EN_KILLFOCUS(DSCV_ED2, &Dlg_Open_DMCV::OnEnKillfocusEd2)
	ON_EN_KILLFOCUS(DSCV_ED3, &Dlg_Open_DMCV::OnEnKillfocusEd3)
	ON_NOTIFY(LVN_KEYDOWN, DSCV_L1, &Dlg_Open_DMCV::OnLvnKeydownL1)
	ON_NOTIFY(LVN_KEYDOWN, DSCV_L2, &Dlg_Open_DMCV::OnLvnKeydownL2)
	ON_NOTIFY(LVN_KEYDOWN, DSCV_L3, &Dlg_Open_DMCV::OnLvnKeydownL3)	
	ON_EN_SETFOCUS(DSCV_ED1, &Dlg_Open_DMCV::OnEnSetfocusEd1)
	ON_EN_SETFOCUS(DSCV_T1, &Dlg_Open_DMCV::OnEnSetfocusT1)
	ON_EN_SETFOCUS(DSCV_ED2, &Dlg_Open_DMCV::OnEnSetfocusEd2)
	ON_EN_SETFOCUS(DSCV_ED3, &Dlg_Open_DMCV::OnEnSetfocusEd3)
	ON_BN_CLICKED(DSCV_CBS1, &Dlg_Open_DMCV::OnBnClickedCbs1)
	ON_BN_CLICKED(DSCV_CBS2, &Dlg_Open_DMCV::OnBnClickedCbs2)
	ON_BN_CLICKED(DSCV_CBS3, &Dlg_Open_DMCV::OnBnClickedCbs3)
	ON_BN_CLICKED(DSCV_CKLUU, &Dlg_Open_DMCV::OnBnClickedCkluu)
	ON_BN_CLICKED(DSCV_CBS4, &Dlg_Open_DMCV::OnBnClickedCbs4)
	ON_BN_CLICKED(DSCV_B13, &Dlg_Open_DMCV::OnBnClickedB13)	
	ON_BN_CLICKED(DSCV_CBS5, &Dlg_Open_DMCV::OnBnClickedCbs5)
	ON_NOTIFY(NM_DBLCLK, DSCV_L3, &Dlg_Open_DMCV::OnNMDblclkL3)
	ON_NOTIFY(NM_DBLCLK, DSCV_L2, &Dlg_Open_DMCV::OnNMDblclkL2)	
	ON_NOTIFY(NM_RCLICK, DSCV_L2, &Dlg_Open_DMCV::OnNMRClickL2)
	ON_COMMAND(ID_TI40039_5, &Dlg_Open_DMCV::OnTi40039)
	ON_COMMAND(ID_TI40040_5, &Dlg_Open_DMCV::OnTi40040)
	ON_NOTIFY(NM_RCLICK, DSCV_L1, &Dlg_Open_DMCV::OnNMRClickL1)
	ON_BN_CLICKED(DSCV_CK2, &Dlg_Open_DMCV::OnBnClickedCk2)
	ON_COMMAND(ID_TI40051, &Dlg_Open_DMCV::OnTi40051)
	ON_CBN_SELCHANGE(DSCV_CBBFL, &Dlg_Open_DMCV::OnCbnSelchangeCbbfl)
	ON_BN_CLICKED(DSCV_PRE, &Dlg_Open_DMCV::OnBnClickedPre)
	ON_BN_CLICKED(DSCV_NXT, &Dlg_Open_DMCV::OnBnClickedNxt)
	ON_NOTIFY(LVN_ENDSCROLL, DSCV_L1, &Dlg_Open_DMCV::OnLvnEndScrollL1)
	ON_BN_CLICKED(DSCV_CKDMCVCON, &Dlg_Open_DMCV::OnBnClickedCkdmcvcon)
END_MESSAGE_MAP()

BEGIN_EASYSIZE_MAP(Dlg_Open_DMCV)
	EASYSIZE(DSCV_S1,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(DSCV_T1,ES_BORDER,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(DSCV_PRE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(DSCV_NXT,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(DSCV_CBBFL,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(DSCV_S2,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(DSCV_L1,ES_BORDER,ES_BORDER,ES_BORDER,ES_BORDER,0)
	EASYSIZE(DSCV_B1,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(DSCV_B2,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(DSCV_B3,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(DSCV_B4,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(DSCV_B5,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(DSCV_B6,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)

	EASYSIZE(DSCV_G2,ES_BORDER,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(DSCV_L2,ES_BORDER,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(DSCV_L3,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)

	EASYSIZE(DSCV_CBS5,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)
	EASYSIZE(DSCV_CBB1,ES_BORDER,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	
	EASYSIZE(DSCV_B7,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(DSCV_B8,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(DSCV_B9,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(DSCV_B10,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(DSCV_B11,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)

	EASYSIZE(DSCV_S9,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)
	EASYSIZE(DSCV_CBS4,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)
	EASYSIZE(DSCV_B13,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)
	EASYSIZE(DSCV_S10,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)

	EASYSIZE(DSCV_CBS1,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)
	EASYSIZE(DSCV_CBS2,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)
	EASYSIZE(DSCV_CBS3,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)

	EASYSIZE(DSCV_CK2,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(DSCV_CK1,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	
	EASYSIZE(DSCV_CKDMCVCON, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(DSCV_CKLUU,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(DSCV_T2,ES_BORDER,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
END_EASYSIZE_MAP

BOOL Dlg_Open_DMCV::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	iTCclick=-1;
	sDownTC=L"",sFullTC=L"";

	m_Txt.SubclassDlgItem(DSCV_T1,this);
	m_Txt.SetBkColor(WHITE);
	m_Txt.SetTextColor(BLUE);

	m_chkDMCVCon.SetCheck(iLuuDMCVCon);
	m_chkLuuTC.SetCheck(iLuuTCCV);
	m_txtTChon.SubclassDlgItem(DSCV_T2,this);
	m_txtTChon.SetBkColor(WHITE);
	m_txtTChon.SetTextColor(BLUE);

	m_chkColor.SetCheck(chkClrCV);
	m_btnColor.EnableWindow(chkClrCV);
	m_txtColor.ShowWindow(chkClrCV);

	m_txtColor.SubclassDlgItem(DSCV_S10,this);
	m_txtColor.SetBkColor(fclrCV);
	m_txtColor.SetTextColor(BLACK);

	m_Group.SubclassDlgItem(DSCV_S2,this);
	m_Group.SetBkColor(YELLOW);
	m_Group.SetTextColor(BLACK);

	m_txtTChon.EnableWindow(iLuuTCCV);
	m_cbbox.EnableWindow(jNhomTC);
	m_chkBTC.SetCheck(jNhomTC);

	sLuuMaCT=L"",sLuuTenCT=L"",sLuuDVT=L"",sLuuLink=L"";
	sBefore=L"",sAfter=L"";

	iStopload=1;
	tslkq=0,lanshow=1,ibuocnhay=50;
	_iTab=0,nClick=-1;
	_isl1=0,_isl2=0;
	m_btn[1].SetCheck(1);

	f_Load_ToolTip();
	m_Txt.SetWindowText(xl_tukhoa);
	m_Txt.SetCueBanner(L"Nhập từ khóa tìm kiếm (Ví dụ: AK.8311, sơn cửa kính,... )");

	f_StyleList();
	f_StyleList2();
	f_StyleList3();	

	if(iCpy==0)
	{
		m_cbk1.SetCheck(1);
		m_cbk2.SetCheck(1);
	}
	else if(iCpy==1) m_cbk1.SetCheck(1);
	else if(iCpy==2) m_cbk2.SetCheck(1);

	if(jLuuDGKL==1) m_cbk3.SetCheck(jLuuDGKL);

	// Bổ sung 26.07.2017
	if(iTestnull==1)
	{
		f_Load_all_bieugia();
		
		// Khóa tất cả các nút
		m_Txt.EnableWindow(0);
		m_cbbFile.EnableWindow(0);
		m_btnPre.EnableWindow(0);
		m_btnNxt.EnableWindow(0);
		m_cbbFile.EnableWindow(0);
		for (int i = 0; i < 9; i++) bxx[i].EnableWindow(0);
	}
	else f_Load_DL();

	this->SetWindowTextW(L"Danh sách công việc | " + sFullPath);
	
	_GetDanhsachnhom();
	_LoadlistFile();

	INIT_EASYSIZE;

	// Set size dialog mfc
	if(irWCV>0 && irHCV>0)
	{
		this->SetWindowPos(NULL,0,0,irWCV,irHCV,SWP_NOMOVE | SWP_NOZORDER);

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
BOOL Dlg_Open_DMCV::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if (pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(DSCV_T1))
	{
		OnBnClickedB1();
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(DSCV_L1))
	{
		OnBnClickedB2();
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		if(ClickEsc==0) OnBnClickedB6();
		else GotoDlgCtrl(GetDlgItem(DSCV_L1));
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
		if(pWndWithFocus == GetDlgItem(DSCV_L2))
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

void Dlg_Open_DMCV::OnSize(UINT nType, int cx, int cy) 
{
	CDialog::OnSize(nType, cx, cy);
	UPDATE_EASYSIZE;
}

void Dlg_Open_DMCV::OnSizing(UINT fwSide, LPRECT pRect) 
{
	CDialog::OnSizing(fwSide, pRect);
	EASYSIZE_MINSIZE(320,240,fwSide,pRect);
}


void Dlg_Open_DMCV::f_Load_ToolTip()
{
	CButton * chk = (CButton *)GetDlgItem(DSCV_CK2);
	CButton * btn0 = (CButton *)GetDlgItem(DSCV_B7);
	CButton * btn2 = (CButton *)GetDlgItem(DSCV_B8);
	CButton * btn3 = (CButton *)GetDlgItem(DSCV_B9);
	CButton * btn4 = (CButton *)GetDlgItem(DSCV_B10);
	CButton * btn5 = (CButton *)GetDlgItem(DSCV_B11);
	CButton * btn6 = (CButton *)GetDlgItem(DSCV_CBS5);
	CButton * btn7 = (CButton *)GetDlgItem(DSCV_PRE);
	CButton * btn8 = (CButton *)GetDlgItem(DSCV_NXT);

	CListCtrl * pls1 = (CListCtrl *)GetDlgItem(DSCV_L1);
	CListCtrl * pls2 = (CListCtrl *)GetDlgItem(DSCV_L2);

	m_ToolTips.Create(this);
	m_ToolTips.SetMaxTipWidth(600);
	m_ToolTips.SetDelayTime(900);

	m_ToolTips.AddTool(chk, L"Mặc định tích chọn là trong hợp đồng. Bỏ tích là ngoài hợp đồng.");
	m_ToolTips.AddTool(btn0, L"Thêm mới dữ liệu tiêu chuẩn hoặc nội dung vào vị trí mong muốn."
		L"\nKích chọn dòng đầu tiên để chèn lên trên đầu."
		L"\nKích chọn trống (NULL) dưới cùng để chèn xuống dưới cùng.");

	m_ToolTips.AddTool(btn2, L"Xóa dữ liệu tiêu chuẩn hoặc nội dung.");
	m_ToolTips.AddTool(btn3, L"Cập nhật các dữ liệu chỉnh sửa.");
	m_ToolTips.AddTool(btn4, L"Kích chọn và di chuyển tiêu chuẩn/nội dung lên trên.");
	m_ToolTips.AddTool(btn5, L"Kích chọn và di chuyển tiêu chuẩn/nội dung xuống dưới.");
	m_ToolTips.AddTool(btn6, L"Tích chọn để đưa tiêu chuẩn sang 1 bảng riêng biệt.");
	m_ToolTips.AddTool(btn7, L"Kích để tìm kiếm dữ liệu từ tệp tin trước.");
	m_ToolTips.AddTool(btn8, L"Kích để tìm kiếm dữ liệu từ tệp tin tiếp theo.");

	m_ToolTips.AddTool(pls1, L"Kích để áp dụng công việc. Giữ Shift/Ctrl để chọn nhiều."
		L"\nKích đúp để chỉnh sửa dữ liệu.");

	m_ToolTips.AddTool(pls2, L"Kích đúp để nhập từ khóa tra cứu tiêu chuẩn.\nKích chuột phải để đọc tiêu chuẩn.");
}


void Dlg_Open_DMCV::f_Get_size()
{
	CRect rectCtrl;
	GetWindowRect(&rectCtrl);
	irHCV = rectCtrl.Height();
	irWCV = rectCtrl.Width();
}


void Dlg_Open_DMCV::f_StyleList()
{
	lit_dmcv.InsertColumn(0,L"",LVCFMT_LEFT,20);
	lit_dmcv.InsertColumn(1,L"Mã CV",LVCFMT_CENTER,_wL1[0]);
	lit_dmcv.InsertColumn(2,L"Tên công việc",LVCFMT_LEFT,_wL1[1]);
	lit_dmcv.InsertColumn(3,L"Đơn vị tính",LVCFMT_CENTER,100);
	lit_dmcv.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_CHECKBOXES | LVS_EX_LABELTIP);

	if(iTestnull==1) return;

	lit_dmcv.SetColumnReadOnly(0);
	lit_dmcv.SetColumnColors(0,WHITE,WHITE);
	lit_dmcv.SetDefaultEditor(NULL, NULL, &txtE_DMCV);
	txtE_DMCV.SubclassDlgItem(DSCV_ED1,this);txtE_DMCV.SetBkColor(YELLOW153);txtE_DMCV.SetTextColor(BLUE);
}


void Dlg_Open_DMCV::f_StyleList2()
{
	lit_tccv.InsertColumn(0,L"",LVCFMT_CENTER,0);
	lit_tccv.InsertColumn(1,L"Mã TC",LVCFMT_LEFT,_wL2[0]);
	lit_tccv.InsertColumn(2,L"Tiêu chuẩn áp dụng",LVCFMT_LEFT,_wL2[1]);
	lit_tccv.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);

	if(iTestnull==1) return;

	lit_tccv.SetColumnReadOnly(0);
	lit_tccv.SetColumnReadOnly(1);
	lit_tccv.SetDefaultEditor(NULL, NULL, &txtE_TCCV);
	txtE_TCCV.SubclassDlgItem(DSCV_ED2,this);txtE_TCCV.SetBkColor(YELLOW153);txtE_TCCV.SetTextColor(BLUE);
}


void Dlg_Open_DMCV::f_StyleList3()
{
	lit_ndkt.InsertColumn(0,L"",LVCFMT_CENTER,0);
	lit_ndkt.InsertColumn(1,L"Nội dung kiểm tra",LVCFMT_LEFT,_wL3[0]);
	lit_ndkt.InsertColumn(2,L"Phương pháp kiểm tra",LVCFMT_LEFT,_wL3[1]);
	lit_ndkt.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);

	if(iTestnull==1) return;

	lit_ndkt.SetColumnReadOnly(0);
	lit_ndkt.SetDefaultEditor(NULL, NULL, &txtE_NDCV);
	txtE_NDCV.SubclassDlgItem(DSCV_ED3,this);txtE_NDCV.SetBkColor(YELLOW153);txtE_NDCV.SetTextColor(BLUE);
}


void Dlg_Open_DMCV::f_SaveWidth()
{
	for (int i = 0; i < 2; i++) _wL1[i] = lit_dmcv.GetColumnWidth(i+1);
	for (int i = 0; i < 2; i++) _wL2[i] = lit_tccv.GetColumnWidth(i+1);
	for (int i = 0; i < 2; i++) _wL3[i] = lit_ndkt.GetColumnWidth(i+1);
}


void Dlg_Open_DMCV::OnBnClickedB1()
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
		lit_dmcv.DeleteAllItems();
		ASSERT(lit_dmcv.GetItemCount() == 0);

		lit_tccv.DeleteAllItems();
		ASSERT(lit_tccv.GetItemCount() == 0);

		lit_ndkt.DeleteAllItems();
		ASSERT(lit_ndkt.GetItemCount() == 0);

		// Truy vấn & load dữ liệu
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
				_TackTukhoaChuan(val,L"|",L" ",1,0);
				CString szval=L"";

				SqlString = L"SELECT * FROM tbDonGia WHERE (";
				for (int i = 1; i <= iKey; i++)
				{
					szval.Format(L"MHDG LIKE '%s%s%s' OR ",L"%",pKey[i],L"%");
					SqlString+=szval;
				}

				if(SqlString.Right(4)==L" OR ") SqlString = SqlString.Left(SqlString.GetLength()-4);
				SqlString+=L");";
			}
			else
			{
				f_Tack_tukhoa(val);
				f_Tao_truyvan();
			}
		}
		else SqlString = L"SELECT * FROM tbDonGia ORDER BY MHDG ASC;";

		f_Load_DL();
	}
	catch(...){}
}


void Dlg_Open_DMCV::f_Tack_tukhoa(CString sKey)
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


void Dlg_Open_DMCV::f_Tao_truyvan()
{
	try
	{
		CString val=L"";
		if(iKey==1)
		{
			SqlString.Format(L"SELECT * FROM tbDonGia "
			L"WHERE (MHDG LIKE '%s%s%s' OR TCV LIKE '%s%s%s' OR TK LIKE '%s%s%s');"
			,L"%",pKey[1],L"%",L"%",pKey[1],L"%",L"%",pKey[1],L"%");
		}
		else
		{
			SqlString.Format(L"SELECT * FROM tbDonGia " 
			L"WHERE ((MHDG LIKE '%s%s%s' OR TCV LIKE '%s%s%s' OR TK LIKE '%s%s%s')"
			,L"%",pKey[1],L"%",L"%",pKey[1],L"%",L"%",pKey[1],L"%");

			for (int i = 2; i <= iKey; i++)
			{
				val.Format(
					L" AND (MHDG LIKE '%s%s%s' OR TCV LIKE '%s%s%s' OR TK LIKE '%s%s%s')"
					,L"%",pKey[i],L"%",L"%",pKey[i],L"%",L"%",pKey[i],L"%");
				SqlString+=val;
			}
			SqlString+= L") ORDER BY MHDG ASC;";
		}
	}
	catch(...){}
}

void Dlg_Open_DMCV::f_Danh_sachfile()
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

void Dlg_Open_DMCV::f_Truy_van_DL()
{
	TRY
	{
		CString szval=L"";
		for (int i = 1; i <= slfileDL; i++)
		{
			//if(sFiledulieu[i].Find(L"\\")>0) szval = sFiledulieu[i];
			//else szval.Format(L"%s\\%s",sDDanfile,sFiledulieu[i]);

			szval = sFpathdulieu[i];
			ObjConn.OpenAccessDB(szval,L"",L"");
			CRecordset* Recset = ObjConn.Execute(SqlString);
			iKqua=0;
			while( !Recset->IsEOF() )
			{
				iKqua=1;
				Recset->GetFieldValue(L"MHDG",XLXD[1].CDG[0]);
				Recset->GetFieldValue(L"TCV",XLXD[1].CDG[1]);
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
		AfxMessageBox(_T("[1] Lỗi tìm kiếm dữ liệu công việc")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_Open_DMCV::f_Load_DL()
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
			Recset->GetFieldValue(L"MHDG",XLXD[tslkq].CDG[0]);
			Recset->GetFieldValue(L"TCV",XLXD[tslkq].CDG[1]);
			Recset->GetFieldValue(L"DVT",XLXD[tslkq].CDG[2]);

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
			lit_dmcv.InsertItem(i,L"",0);
			lit_dmcv.SetItemText(i,1,XLXD[i+1].CDG[0]);
			lit_dmcv.SetItemText(i,2,XLXD[i+1].CDG[1]);
			lit_dmcv.SetItemText(i,3,XLXD[i+1].CDG[2]);
		}

		// Kết quả tìm kiếm
		CString szval=L"";
		if (tslkq >= 4000) szval = L"Tìm thấy rất nhiều kết quả";
		else szval.Format(L"Tổng số: %d kết quả ",tslkq);
		m_Group.SetWindowText(szval);

		if(iz>0) GotoDlgCtrl(GetDlgItem(DSCV_L1));
		else GotoDlgCtrl(GetDlgItem(DSCV_T1));
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[2] Lỗi tìm kiếm dữ liệu công việc")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_Open_DMCV::xl_xdcot_DMCV(_WorksheetPtr pSh)
{
	_iCot[1] = getColumn(pSh,"DMBB_STT");
	_iCot[2] = getColumn(pSh,"DMBB_MACV");
	_iCot[3] = getColumn(pSh,"DMBB_MH");
	_iCot[4] = getColumn(pSh,"DMBB_MAHS");
	_iCot[5] = getColumn(pSh,"DMBB_ND");
	_iCot[7] = getColumn(pSh,"DMBB_DV");
	_iCot[6] = getColumn(pSh,"DMBB_KL");
	_iCot[8] = getColumn(pSh,"DMBB_TC");	

	_iCot[9] = getColumn(pSh,"DMBB_TN_NGAY");
	_iCot[10] = getColumn(pSh,"DMBB_NB_NGAY");
	_iCot[11] = getColumn(pSh,"DMBB_PHIEUYC");
	_iCot[12] = getColumn(pSh,"DMBB_AB_NGAY");

	_iCot[13] = getColumn(pSh,"DMBB_TC2");
	_iCot[21] = getColumn(pSh,"DMBB_PS");
	_iCot[26] = getColumn(pSh,"DMBB_TDDBT_GIO");

	_iCot[27] = getColumn(pSh,L"DMBB_DOT");
	_iCot[28] = getColumn(pSh,L"DMBB_DTLMAU");
	_iCot[29] = getColumn(pSh,L"DMBB_PHIEUYC_GIO");

	_iCot[40] = getColumn(pSh, L"DMBB_HK_TONGNGHI");
	_iCot[41] = getColumn(pSh, L"DMBB_HK_CTNGHI");

	_iCot[30] = getColumn(pSh,L"DMBB_HK_TONGNGAY");
	_iCot[31] = getColumn(pSh,L"DMBB_HK_NGAYBD");
	_iCot[32] = getColumn(pSh,L"DMBB_HK_NGAYKT");
	_iCot[33] = getColumn(pSh,L"DMBB_HK_LIENKET");
	
	_iCot[34] = getColumn(pSh,L"DMBB_HK_TNNC");
	_iCot[35] = getColumn(pSh,L"DMBB_HK_TNMAY");

	_iCot[36] = getColumn(pSh,L"DMBB_HK_DMNC");
	_iCot[37] = getColumn(pSh,L"DMBB_HK_HESO");
	_iCot[38] = getColumn(pSh,L"DMBB_HK_TONGCONG");
	_iCot[39] = getColumn(pSh,L"DMBB_HK_NCTN");	

	_iCotEND = getColumn(pSh, L"DMBB_MAUTDBT");
	_iCotMaDM = getColumn(pSh, L"DMBB_MADM");	
}


void Dlg_Open_DMCV::f_Tim_kiem_DMCV()
{
	try
	{
		iTestnull = 0;		
		CString szval = L"", szFormula = L"";

		psDMCV = pWb->ActiveSheet;
		prDMCV = psDMCV->Cells;
		iRow = psDMCV->Application->ActiveCell->Row;
		iColumn = psDMCV->Application->ActiveCell->Column;		

		// Xác định vị trí activate
		xl_xdcot_DMCV(psDMCV);

		// Xác định vị trí END
		xde = FindComment(psDMCV,_iCot[1],"END");
		if(iRow<8 || iRow>=xde) return;

/////// Kiểm tra vị trí activate = cột tra mã
		if(iColumn==_iCot[5])
		{
			// Đổi mã nếu đổi tên
			CString val= GIOR(iRow,1,prDMCV,L"");
			if(_wtoi(val)>0)
			{
				val = GIOR(iRow,_iCot[2],prDMCV,L"");
				if(val!=L"")
				{
					CString szMaL=val,szTenBF=L"",szTenAT=L"";
					if(szMaL.GetLength()>5 && (szMaL.Right(3)).Left(1)==L".") szMaL = val.Left(val.GetLength()-3);
					int vtL=0,demL=0,posL=0;

					while (true)
					{
						vtL = FindValue(psDMCV,_iCot[2],8,iRow-1,(_bstr_t)val);
						if(vtL==0) vtL = FindValue(psDMCV,_iCot[2],iRow+1,xde,(_bstr_t)val);
						if(vtL==0)
						{
							posL=1;
							break;
						}
						else
						{
							// Nếu đã tồn tại mã như vậy, tiếp tục kiểm tra xem tên có trùng không?
							// Nếu trùng bỏ qua, ngược lại tạo mã mới
							szTenBF = GIOR(vtL,_iCot[5],prDMCV,L"");
							szTenAT = GIOR(iRow,_iCot[5],prDMCV,L"");
							if(szTenBF==szTenAT) break;
						}

						demL++;
						if(demL<=9) val.Format(L"%s.0%d",szMaL,demL);
						else val.Format(L"%s.%d",szMaL,demL);
					}

					if (posL==1) prDMCV->PutItem(iRow,_iCot[2],(_bstr_t)val);
				}
			}
			else
			{
				// Kiểm tra vị trí tính diễn giải khối lượng
				CString s1=GIOR(iRow,_iCot[2],prDMCV,L"");
				CString s2=GIOR(iRow,_iCot[3],prDMCV,L"");
				if(s1!=L"" || s2!=L"") return;

				// Tính diễn giải khối lượng
				xl_calculator_dgkl(prDMCV,iRow,iColumn,_iCot[6]);

				// Hiển thị 2 cột đơn vị và khối lượng
				PRS = prDMCV->GetItem(1,_iCot[6]);
				int nCW = PRS->ColumnWidth;
				if(nCW==0)
				{
					psDMCV->GetRange(prDMCV->GetItem(1,_iCot[7]),prDMCV->GetItem(1,_iCot[7]))->EntireColumn->Hidden=false;
					psDMCV->GetRange(prDMCV->GetItem(1,_iCot[6]),prDMCV->GetItem(1,_iCot[6]))->EntireColumn->Hidden=false;
				}
					// Chèn thêm dòng (nếu cần)
					s1=GIOR(iRow+1,_iCot[1],prDMCV,L"");
					if(_wtoi(s1)>0)
					{
						PRS = prDMCV->GetItem(iRow+1,1);
						PRS->EntireRow->Insert(xlShiftDown);
						xde++;
					}

				// Tính lại hàm sum
				int ibd=0,ikt=0;
			
				// Xác định vị trí bắt đầu & kết thúc
				for (int i = iRow-1; i >= 8; i--)
				{
					s1=GIOR(i,_iCot[1],prDMCV,L"");
					s2=GIOR(i,_iCot[2],prDMCV,L"");
					if(_wtoi(s1)>0 || s2==L"HM" || s2==L"GD")
					{
						ibd=i;
						break;
					}
				}
				if(ibd==0) return;

				for (int i = iRow+1; i < xde; i++)
				{
					s1=GIOR(i,_iCot[1],prDMCV,L"");
					s2=GIOR(i,_iCot[2],prDMCV,L"");
					if(_wtoi(s1)>0 || s2==L"HM" || s2==L"GD")
					{
						ikt=i;
						break;
					}
				}
				if(ikt==0) ikt=xde;

				// Xác định vị trí cuối cùng có dữ liệu
				for (int i = ikt-1; i >= iRow; i--)
				{
					s1=GIOR(i,_iCot[5],prDMCV,L"");
					if(s1!=L"")
					{
						ikt=i+1;
						break;
					}
				}

				// Tính tổng
				s1.Format(L"=SUM(%s:%s)",GAR(ibd+1,_iCot[6],prDMCV,0),GAR(ikt-1,_iCot[6],prDMCV,0));
				prDMCV->PutItem(ibd,_iCot[6],(_bstr_t)s1);
			}

			return;
		}
		else if(iColumn==_iCot[9] || iColumn==_iCot[10] || iColumn==_iCot[11] || iColumn==_iCot[12])
		{
			// Leo edit 01.06.2018
			szval= GIOR(iRow,1,prDMCV,L"");
			if(_wtoi(szval)>0) CheckNhomKy(psDMCV,iRow,getColumn(psDMCV,L"DMBB_KBB"));
			return;
		}
		else if(iColumn==_iCot[8] || iColumn==_iCot[13])
		{
			Dlg_trauu_tieuchuan _dlg;
			_dlg.xl_timkiem_tieuchuan();
			return;
		}
		else if (iColumn == _iCot[2] || iColumn == _iCot[3] || iColumn == _iCot[21])
		{
			// Get Formula
			if (iColumn == _iCot[2])
			{
				PRS = prDMCV->GetItem(iRow, iColumn);
				szFormula = PRS->GetFormula();
			}

			// UPPER
			xl_tukhoa = GIOR(iRow, iColumn, prDMCV, L"");
			xl_tukhoa.MakeUpper();
			prDMCV->PutItem(iRow, iColumn, (_bstr_t)xl_tukhoa);
		}
		else
		{
			int icheckproject = 0;
			for (int i = 30; i <= 41; i++)
			{
				if (iColumn == _iCot[i])
				{
					icheckproject = 1;
					break;
				}
			}

			if (icheckproject == 1)
			{
				// Microsoft Project	<-- Leo 20.04.2019
				_EnterProject();
				return;
			}			
		}

		// Kiểm tra vị trí tra cứu dữ liệu
		if(iColumn!=_iCot[2]) return;

		CString sz5 = GIOR(iRow, _iCot[5], prDMCV, L""); sz5.Trim();
		CString sz8 = GIOR(iRow, _iCot[8], prDMCV, L""); sz8.Trim();

		// Kiểm tra các vị trí cố định Freeze (không thể tra cứu)
		// Tồn tại số tại cột STT
		szval = GIOR(iRow,_iCot[1],prDMCV,L"");
		szval.Trim();

		if (szval == L"")
		{
			szval = GIOR(iRow, _iCot[2], prDMCV, L"");
			szval = szval.Left(2);
			szval.MakeUpper();
			if (szval != L"HM" && szval != L"GD")
			{
				// Tồn tại giá trị cột tiêu chuẩn hoặc diễn giải
				if ((sz5 != L"" || sz8 != L"") && szFormula.Left(1) != L"=") return;
			}
		}

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
				szval = GIOR(iRow,_iCot[2],prDMCV,L"");
				szval.MakeUpper();
				if(szval==L"")
				{
					szval = L"'" + szFormula;
					prDMCV->PutItem(iRow,_iCot[2],(_bstr_t)szval);
				}
				else if(szval==L"0") ichecknull=1;

				int vtkt = iRow;
				if(iRowCopy>iRow) vtkt = xde;

				int iRowEndCopy=vtkt;
				if(ichecknull==0)
				{
					if(szval==L"HM")
					{
						for (int i = iRowCopy+1; i < vtkt; i++)
						{
							szval = GIOR(i,_iCot[2],prDMCV,L"");
							if(szval==L"HM")
							{
								iRowEndCopy=i;
								break;
							}
						}
					}
					else if(szval==L"GD")
					{
						for (int i = iRowCopy+1; i < vtkt; i++)
						{
							szval = GIOR(i,_iCot[2],prDMCV,L"");
							if(szval==L"HM" || szval==L"GD")
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
							szval = GIOR(i,_iCot[2],prDMCV,L"");
							if(szval!=L"")
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
						sz5 = GIOR(i,_iCot[5],prDMCV,L""); sz5.Trim();
						sz8 = GIOR(i,_iCot[8],prDMCV,L""); sz8.Trim();
						if(sz5!=L"" || sz8!=L"")
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

				PRS = psDMCV->GetRange(prDMCV->GetItem(iRow+1,1),prDMCV->GetItem(iRow+slcp,1));
				PRS->EntireRow->Insert(xlShiftDown);
				
				
				PRS = psDMCV->GetRange(prDMCV->GetItem(iRowStart, 1), prDMCV->GetItem(iRowEndCopy - 1, 1));

				bool blCopyAll = true;	// Tạm thời để copy tất cả (giá trị, công thức,...)
				if (blCopyAll)
				{
					PRS->EntireRow->Copy(prDMCV->GetItem(iRow, 1));
				}
				else
				{
					// Chỉ copy value
					PRS->EntireRow->Copy();

					PRS2 = prDMCV->GetItem(iRow, 1);
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
				f_Auto_stt_dmuc(prDMCV,ntype,0,8,xde,_iCot[1],_iCot[2],_iCot[4]);

				if(slcp>5) PRS = prDMCV->GetItem(iRow,_iCot[2]);
				else
				{
					if(posR>0) PRS = prDMCV->GetItem(iRow+slcp,_iCot[2]);
					else PRS = prDMCV->GetItem(iRow+slcp-1,_iCot[2]);
				}
				PRS->Select();

				return;			
			}
		}

/*********************************************************************/

		getPathQLCL();
		f_Danh_sachfile();

		jNhomTC = _wtoi(getVCell(psTS,L"TS_GRPTC"));
		iCpy = _wtoi(getVCell(psTS,L"CF_DMCOPY"));	// TS_CDATE

		shBTC = name_sheet("shNhomTC");
		psBTC = xl->Sheets->GetItem(shBTC);
		prBTC = psBTC->GetCells();
		int iVisible = psBTC->GetVisible();
		if (iVisible!=-1) psBTC->PutVisible(XlSheetVisibility::xlSheetVisible);

		// Xác định từ khóa tìm kiếm
		xl_tukhoa = GIOR(iRow,iColumn,prDMCV,L"");
		xl_tukhoa.Replace(L"'",L"");
		xl_tukhoa.Trim();
		if(xl_tukhoa==L"")
		{
			SqlString = L"SELECT * FROM tbDonGia ORDER BY MHDG ASC;";

			//if(sFiledulieu[1].Find(L"\\")>0) sFullPath = sFiledulieu[1];
			//else sFullPath.Format(L"%s\\%s",sDDanfile,sFiledulieu[1]);

			sFullPath = sFpathdulieu[1];
			
			AFX_MANAGE_STATE(AfxGetStaticModuleState());
			xl->PutScreenUpdating(VARIANT_TRUE);
			DoModal();
			xl->PutScreenUpdating(VARIANT_FALSE);

			return;
		}

		// Trường hợp là ghi chú (21.03.2019)
		if(xl_tukhoa.Find(L"*")>=0)
		{
			PRS = prDMCV->GetItem(iRow,1);
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
				prDMCV->PutItem(iRow,iColumn,(_bstr_t)sDbiet);
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
				prDMCV->PutItem(iRow,iColumn,(_bstr_t)sDbiet);
				return;
			}
		}

		// Các trường hợp đặc biệt khác (DG,DS)
		sDbiet = xl_tukhoa.Left(2);
		sDbiet.MakeUpper();
		if (sDbiet == L"HM" || sDbiet == L"GD")
		{
			PRS = psDMCV->GetRange(prDMCV->GetItem(iRow, 1), prDMCV->GetItem(iRow, _iCotEND));
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
					szval = GIOR(j, _iCot[2], prDMCV, L"");
					if (szval == L"HM") demhm++;
				}

				prDMCV->PutItem(iRow, _iCot[2], (_bstr_t)L"HM");

				szval.Format(L"HM%d",demhm);
				prDMCV->PutItem(iRow, _iCot[4], (_bstr_t)szval);

				szval = GIOR(iRow, _iCot[5], prDMCV, L"");
				if (szval != L"")
				{
					szval.Format(L"Hạng mục %d", demhm);
					prDMCV->PutItem(iRow, _iCot[5], (_bstr_t)szval);
				}
			}
			else
			{
				PRS->EntireRow->Interior->PutColor(RGB(255, 255, 255));

				PRS->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->PutWeight(xlThin);
				PRS->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->PutLineStyle(xlDot);
				PRS->Borders->GetItem(XlBordersIndex::xlInsideVertical)->Weight = xlThin;

				szval = GIOR(iRow-1, _iCot[2], prDMCV, L"");
				if (szval == L"HM")
				{
					PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->Weight = xlThin;
				}
				else
				{
					PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->PutWeight(xlThin);
					PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->PutLineStyle(xlDot);
				}

				int demgd = 1;
				for (int j = iRow-1; j >= 8; j--)
				{
					szval = GIOR(j, _iCot[2], prDMCV, L"");

					if (szval == L"HM") break;
					if (szval == L"GD") demgd++;
				}

				prDMCV->PutItem(iRow, _iCot[2], (_bstr_t)L"GD");

				szval.Format(L"GD%d", demgd);
				prDMCV->PutItem(iRow, _iCot[4], (_bstr_t)szval);

				szval = GIOR(iRow, _iCot[5], prDMCV, L"");
				if (szval != L"")
				{
					szval.Format(L"Giai đoạn %d", demgd);
					prDMCV->PutItem(iRow, _iCot[5], (_bstr_t)szval);
				}
			}

			if (iRow + 1 == xde)
			{
				PRS = psDMCV->GetRange(prDMCV->GetItem(xde, 1), prDMCV->GetItem(xde+4, 1));
				PRS->EntireRow->Insert(xlShiftDown);

				PRS = psDMCV->GetRange(prDMCV->GetItem(xde, 1), prDMCV->GetItem(xde + 4, _iCotEND));
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
				PRS = psDMCV->GetRange(prDMCV->GetItem(xde, 1), prDMCV->GetItem(xde, _iCotEND));
				PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->Weight = xlThin;
			}

			return;
		}

		// Bổ sung 20.08.2018
		// Tìm mã tương tự ở trên bảng tính
		int irTrung=0;
		CString sktrU = xl_tukhoa;
		sktrU.MakeUpper();
		for (int i = 8; i < iRow-1; i++)
		{
			szval = GIOR(i,iColumn,prDMCV,L"");
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
				szval = GIOR(i,iColumn,prDMCV,L"");
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
				PRS = prDMCV->GetItem(irTrung,1);
				PRS->EntireRow->Copy();

				PRS2 = prDMCV->GetItem(iRow,1);
				PRS2->PasteSpecial(xlPasteValues,Excel::xlPasteSpecialOperationNone);
				xl->PutCutCopyMode(XlCutCopyMode(false));

				PRS = prDMCV->GetItem(iRow,1);
				PRS->EntireRow->Font->PutColor(RGB(0,0,0));
				PRS->EntireRow->Font->PutBold(0);
				PRS->EntireRow->Font->PutItalic(0);
				PRS->EntireRow->Rows->AutoFit();
				if((int)PRS->GetRowHeight()>80) PRS->PutWrapText(FALSE);

				PRS = psDMCV->GetRange(prDMCV->GetItem(iRow,1),prDMCV->GetItem(iRow, _iCotEND));
				if(chkClrCV==1) PRS->Interior->PutColor(fclrCV);
				else PRS->EntireRow->Interior->PutColor(xlNone);

				// Thay đổi mã công việc và mã hồ sơ nghiệm thu
				/*CString bsMaCV = GIOR(irTrung,_iCot[2],prDMCV,L"");
				if(bsMaCV.GetLength()>5 && (bsMaCV.Right(3)).Left(1)==L".") bsMaCV = bsMaCV.Left(bsMaCV.GetLength()-3);

				for (int i = 1; i < 100; i++)
				{
					if(i<=9) szval.Format(L"%s.0%d",bsMaCV,i);
					else szval.Format(L"%s.%d",bsMaCV,i);

					 if((int)FindValue(psDMCV,_iCot[2],1,0,(_bstr_t)szval)==0)
					 {
						 bsMaCV = L"'" + szval;
						 prDMCV->PutItem(iRow,_iCot[2],(_bstr_t)bsMaCV);
						 break;
					 }
				}*/

				CString bsMaHSNT = GIOR(irTrung,_iCot[4],prDMCV,L"");
				for (int i = 1; i < 100; i++)
				{
					if(i<=9) szval.Format(L"%s.0%d",bsMaHSNT,i);
					else szval.Format(L"%s.%d",bsMaHSNT,i);

					 if((int)FindValue(psDMCV,_iCot[4],1,0,(_bstr_t)szval)==0)
					 {
						 bsMaHSNT = L"'" + szval;
						 prDMCV->PutItem(iRow,_iCot[4],(_bstr_t)bsMaHSNT);
						 break;
					 }
				}

				// Sắp xếp lại stt
				int ntype = _wtoi(getVCell(psTS,L"TS_ATOSTT"));
				if(ntype!=0 && ntype!=1 && ntype!=2) ntype=0;
				f_Auto_stt_dmuc(prDMCV,ntype,0,8,xde,_iCot[1],_iCot[2],_iCot[4]);

				PRS = prDMCV->GetItem(iRow, iColumn);
				PRS->Select();
				return;
			}
		}

		// Bổ sung trường hợp đặc biệt (26.07.2017)		
		if(sDbiet==L"BG" || sDbiet==L"DG" || sDbiet==L"đG")
		{
			iTestnull=1;
			SqlString = L"SELECT * FROM tbDonGia ORDER BY MHDG ASC;";

			//if(sFiledulieu[1].Find(L"\\")>0) sFullPath = sFiledulieu[1];
			//else sFullPath.Format(L"%s\\%s",sDDanfile,sFiledulieu[1]);

			sFullPath = sFpathdulieu[1];
			
			AFX_MANAGE_STATE(AfxGetStaticModuleState());
			xl->PutScreenUpdating(VARIANT_TRUE);
			DoModal();
			xl->PutScreenUpdating(VARIANT_FALSE);

			return;
		}

		// Bổ sung 18.03.2019
		if((int)xl_tukhoa.Find(L" ")==-1 && (int)xl_tukhoa.Find(L"+")==-1)
		{
			if(xl_tukhoa.GetLength()>5 && (xl_tukhoa.Right(3)).Left(1)==L".") xl_tukhoa = xl_tukhoa.Left(xl_tukhoa.GetLength()-3);
		}

		// Bổ sung 16.12.2017
		if((int)xl_tukhoa.Find(L" ")==-1 && (int)xl_tukhoa.Find(L"+")==-1 && (int)xl_tukhoa.Find(L"|")>0)
		{
			_TackTukhoaChuan(xl_tukhoa,L"|",L" ",1,0);
			SqlString = L"SELECT * FROM tbDonGia WHERE (";
			for (int i = 1; i <= iKey; i++)
			{
				szval.Format(L"MHDG LIKE '%s%s%s' OR ",L"%",pKey[i],L"%");
				SqlString+=szval;
			}

			if(SqlString.Right(4)==L" OR ") SqlString = SqlString.Left(SqlString.GetLength()-4);
			SqlString+=L") ORDER BY MHDG ASC;";
		}
		else
		{
			f_Tack_tukhoa(xl_tukhoa);
			f_Tao_truyvan();
		}

///////// Truy vấn dữ liệu
		f_Truy_van_DL();

		if(iKqua>0)
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
			PRS = prDMCV->GetItem(iRow,iColumn);
			PRS->Font->PutColor(RGB(255,0,0));
			PRS->Activate();
		}
	}
	catch(...){_s(L"Lỗi tra cứu dữ liệu công việc.");}
}

void Dlg_Open_DMCV::OnBnClickedB3()
{
	try
	{
		int count = lit_dmcv.GetItemCount();
		if(count==0) return;

		int ktrchk=-1;
		for (int i = 0; i < count; i++)
		{
			if((int)lit_dmcv.GetCheck(i)==1)
			{
				ktrchk=i;
				break;
			}
		}

		if(ktrchk==-1)
		{
			_s(L"Bạn chưa tích chọn công việc cần cập nhật dữ liệu.");
			return;
		}

		if(_yn(L"Bạn chắc chắn cập nhật tiêu chuẩn và nội dung "
			L"cho các công việc được tích chọn?",0)==6)
		{
			ObjConn.OpenAccessDB(sFullPath,L"",L"");
			ObjConn2.OpenAccessDB(pathMDB,L"",L"");

			for (int i = ktrchk; i < count; i++)
			{
				if((int)lit_dmcv.GetCheck(i)==1) _UpdateDL(i,0);				
			}

			ObjConn2.CloseDatabase();
			ObjConn.CloseDatabase();

			_s(L"Cập nhật dữ liệu thành công!");
		}
	}
	catch(...){}
}


// 09.06.2017: Kiểm tra tiêu chuẩn trùng
int Dlg_Open_DMCV::f_Check_tieuchuan(CString sKtraTC)
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


void Dlg_Open_DMCV::f_Load_all_bieugia()
{
	try
	{
		shBieugia = name_sheet("shDTXD");
		psBieugia=xl->Sheets->GetItem(shBieugia);
		prBieugia=psBieugia->GetCells();

		// Xác định cột
		int icot1 = getColumn(psBieugia,L"BIEUGIA_STT");
		int icot5 = getColumn(psBieugia,L"BIEUGIA_MHDG");
		int icot6 = getColumn(psBieugia,L"BIEUGIA_TEN");
		int icot7 = getColumn(psBieugia,L"BIEUGIA_DV");
		int icot8 = getColumn(psBieugia,L"BIEUGIA_KL");

		tslkq=0;
		CString szv[2]={L"",L""};
		int iEnd = FindComment(psBieugia,1,"END");
		for (int i = 7; i < iEnd; i++)
		{
			szv[0] = GIOR(i,icot1,prBieugia,L"");
			szv[1] = GIOR(i,icot5,prBieugia,L"");

			if(_wtoi(szv[0])>0 && szv[1]!=L"")
			{
				tslkq++;
				XLXD[tslkq].CDG[0] = szv[1];
				XLXD[tslkq].CDG[1] = GIOR(i,icot6,prBieugia,L"");
				XLXD[tslkq].CDG[2] = GIOR(i,icot7,prBieugia,L"");
			}
		}

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
			lit_dmcv.InsertItem(i,L"",0);
			lit_dmcv.SetItemText(i,1,XLXD[i+1].CDG[0]);
			lit_dmcv.SetItemText(i,2,XLXD[i+1].CDG[1]);
			lit_dmcv.SetItemText(i,3,XLXD[i+1].CDG[2]);
		}

		CString szval=L"";
		if (tslkq >= 4000) szval = L"Tìm thấy rất nhiều kết quả";
		else szval.Format(L"Tổng số: %d kết quả ",tslkq);
		m_Group.SetWindowText(szval);

		if(iz>0) GotoDlgCtrl(GetDlgItem(DSCV_L1));
		else GotoDlgCtrl(GetDlgItem(DSCV_T1));
	}
	catch(...){}
}


void Dlg_Open_DMCV::_GetDanhsachnhom()
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

void Dlg_Open_DMCV::_LoadlistFile()
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

void Dlg_Open_DMCV::OnBnClickedB2()
{
  try
  {
	ClickEsc=0;

	// Kiểm tra có tích chọn hay không?
	int iTCCVluutam = iLuuTCCV;
	if(iTCCVluutam==1 && sLuuMaCT==L"") iTCCVluutam=0;

	// Kiểm tra có chia nhỏ công tác không?
	int slctcon = 0;
	int ChiaCongtac = iLuuDMCVCon;
	if (ChiaCongtac == 1 && _CheckMultiItem() == 0) ChiaCongtac = 0;

	// Kiểm tra có checkbox hay không?
	int tonglkq = lit_dmcv.GetItemCount();
	int ktrchk=0;
	for (int i = 0; i < tonglkq; i++)
	{
		if((int)lit_dmcv.GetCheck(i)==1)
		{
			ktrchk=1;
			break;
		}
	}

	POSITION pos=lit_dmcv.GetFirstSelectedItemPosition();
	if(ktrchk==0 && pos==NULL)
	{
		_s(L"Bạn chưa chọn dữ liệu. Vui lòng kiểm tra lại.");
		return;
	}

	xl->PutScreenUpdating(VARIANT_FALSE);

	// Kiểm tra tên tiêu chuẩn đã tồn tại chưa? (vtnhom>0)
	int icTen = getColumn(psBTC,L"NHTC_TEN");

	vtnhom=0;
	jNhomTC = m_chkBTC.GetCheck();
	if(jNhomTC==1)
	{
		m_cbbox.GetWindowTextW(sAfter);
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

	int iEnd = FindComment(psDMCV,_iCot[1],"END");	

////////////////////////////////////////////////////////////////

	ObjConn2.OpenAccessDB(pathMDB,L"",L"");
	CRecordset *RecsetTC,*RecsetND;
	CString sMSCV=L"",sNA=L"",sDVT=L"",sLLink=L"",val=L"",stemp=L"",valtmp=L"";
	CString szval=L"";
	int fla=0;
	slMaTC=0;

	// iTCCVluutam=1 tương ứng tích chọn vào checkbox "tra theo tích chọn"
	if(iTCCVluutam==1)
	{
		// Bổ sung 25.06.2019
		_TackTukhoaChuan(sLuuLink, L"|", L";", 0, 1);
		for (int i = 1; i <= iKey; i++) MSCVC[i].szPath = pKey[i];

		_TackTukhoaChuan(sLuuDVT, L"|", L";", 0, 1);
		for (int i = 1; i <= iKey; i++) MSCVC[i].szDVT = pKey[i];

		_TackTukhoaChuan(sLuuTenCT, L"|", L";", 0, 1);
		for (int i = 1; i <= iKey; i++) MSCVC[i].szNoidung = pKey[i];

		_TackTukhoaChuan(sLuuMaCT, L"|", L" ", 0, 1);
		for (int i = 1; i <= iKey; i++) MSCVC[i].szMahieu = pKey[i];
		
		slctcon = iKey;

		// Bổ sung 16.12.2017
		sMSCV = sLuuMaCT;
		sNA = sLuuTenCT;
		sDVT = sLuuDVT;
		sLLink = sLuuLink;

		if(sMSCV.Right(1)==L"|") sMSCV = sMSCV.Left(sMSCV.GetLength()-1);
		sMSCV.Replace(L"|",L"|\n");

		if(sNA.Right(1)==L"|") sNA = sNA.Left(sNA.GetLength()-1);
		sNA.Replace(L"|",L";\n");

		if(sDVT.Right(1)==L"|") sDVT = sDVT.Left(sDVT.GetLength()-1);
		sDVT.Replace(L"|",L"|\n");

		if(sLLink.Right(1)==L"|") sLLink = sLLink.Left(sLLink.GetLength()-1);
		sLLink.Replace(L"|",L"|\n");

		slctcon = iKey;
		for(int i=1; i<=iKey; i++)
		{
			fla=1;

			//Get TC
			SqlString.Format(L"SELECT * FROM tbMaCV_Tieuchuan WHERE MaCV LIKE '%s';",pKey[i]);
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
		// ktrchk = 1 tương ứng là có tích chọn checkbox trong list
		// ktrchk = 0 là chỉ lựa chọn bằng cách SHIFT + click chuột (LVIS_SELECTED)
		if(ktrchk==1)
		{
			slctcon = 0;
			for (int i = 0; i < lit_dmcv.GetItemCount(); i++)
			{
				if ((int)lit_dmcv.GetCheck(i) == 1)
				{
					slctcon++;

					fla = 1;
					val = lit_dmcv.GetItemText(i, 1);
					MSCVC[slctcon].szMahieu = val;

					//Get MACV & NA--------------
					if (sMSCV == L"") sMSCV.Format(L"%s", val);
					else sMSCV.Format(L"%s|\n%s", sMSCV, val);

					valtmp = lit_dmcv.GetItemText(i, 2);
					MSCVC[slctcon].szNoidung = valtmp;

					if (sNA == L"") sNA.Format(L"%s", valtmp);
					else sNA.Format(L"%s;\n%s", sNA, valtmp);

					valtmp = lit_dmcv.GetItemText(i, 3);
					MSCVC[slctcon].szDVT = valtmp;

					if (sDVT == L"") sDVT.Format(L"%s", valtmp);
					else sDVT.Format(L"%s|\n%s", sDVT, valtmp);

					MSCVC[slctcon].szPath = sFullPath;
					if (sLLink == L"") sLLink.Format(L"%s", sFullPath);
					else sLLink.Format(L"%s|\n%s", sLLink, sFullPath);

					//Get MACV & NA--------------/

					//Get TC
					SqlString.Format(L"SELECT * FROM tbMaCV_Tieuchuan WHERE MaCV LIKE '%s';", val);
					RecsetTC = ObjConn2.Execute(SqlString);
					while (!RecsetTC->IsEOF())
					{
						RecsetTC->GetFieldValue(L"MaTC", stemp);
						if (f_Check_tieuchuan(stemp) == 0)
						{
							slMaTC++;
							sKTMaTC[slMaTC] = stemp;
							sKTTenTC[slMaTC] = stemp + L": Không tìm thấy tên tiêu chuẩn";

							SqlString.Format(L"SELECT * FROM tbTentieuchuan WHERE MaTC LIKE '%s';", stemp);
							RecsetND = ObjConn2.Execute(SqlString);
							while (!RecsetND->IsEOF())
							{
								RecsetND->GetFieldValue(L"TenTC", sKTTenTC[slMaTC]);
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
		else
		{
			for (int i = 0; i < lit_dmcv.GetItemCount(); i++)
			{
				if (lit_dmcv.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					slctcon++;

					fla = 1;
					val = lit_dmcv.GetItemText(i, 1);
					MSCVC[slctcon].szMahieu = val;

					//Get MACV & NA--------------
					if (sMSCV == L"") sMSCV.Format(L"%s", val);
					else sMSCV.Format(L"%s|\n%s", sMSCV, val);

					valtmp = lit_dmcv.GetItemText(i, 2);
					MSCVC[slctcon].szNoidung = valtmp;

					if (sNA == L"") sNA.Format(L"%s", valtmp);
					else sNA.Format(L"%s;\n%s", sNA, valtmp);

					valtmp = lit_dmcv.GetItemText(i, 3);
					MSCVC[slctcon].szDVT = valtmp;

					if (sDVT == L"") sDVT.Format(L"%s", valtmp);
					else sDVT.Format(L"%s|\n%s", sDVT, valtmp);

					MSCVC[slctcon].szPath = sFullPath;
					if (sLLink == L"") sLLink.Format(L"%s", sFullPath);
					else sLLink.Format(L"%s|\n%s", sLLink, sFullPath);

					//Get MACV & NA--------------/

					//Get TC
					SqlString.Format(L"SELECT * FROM tbMaCV_Tieuchuan WHERE MaCV LIKE '%s';", val);
					RecsetTC = ObjConn2.Execute(SqlString);
					//Fill TC

					while (!RecsetTC->IsEOF())
					{
						RecsetTC->GetFieldValue(L"MaTC", stemp);
						if (f_Check_tieuchuan(stemp) == 0)
						{
							slMaTC++;
							sKTMaTC[slMaTC] = stemp;
							sKTTenTC[slMaTC] = stemp + L": Không tìm thấy tên tiêu chuẩn";

							SqlString.Format(L"SELECT * FROM tbTentieuchuan WHERE MaTC LIKE '%s';", stemp);
							RecsetND = ObjConn2.Execute(SqlString);
							while (!RecsetND->IsEOF())
							{
								RecsetND->GetFieldValue(L"TenTC", sKTTenTC[slMaTC]);
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
	}	

	ObjConn2.CloseDatabase();

	// Sửa 12.06.2017
	val = GIOR(iRow+1,_iCot[2],prDMCV,L"");
	int cong=0;
	if(iRow+1>=iEnd || val==L"HM" || val==L"GD")
	{
		cong=1;
		PRS = psDMCV->GetRange(prDMCV->GetItem(iRow+1,1),prDMCV->GetItem(iRow+cong+1,1));
		PRS->EntireRow->Insert(xlShiftDown);

		PRS = psDMCV->GetRange(prDMCV->GetItem(iRow+1,1),prDMCV->GetItem(iRow+cong+1,1));
		PRS->EntireRow->Interior->PutColor(xlNone);

		iEnd+=cong;
	}

////////////////////////////////////////////////////////////////////////////

	// Sử dụng bảng tiêu chuẩn hay không? <-- Leo 08.02.2018
	// Nếu sử dụng (jNhomTC=1) --> đưa tất cả tiêu chuẩn sang sheet Bang tieu chuan
	// Tiếp đó đặt tên nhóm và link sang DMCV/VL

	// Lưu tên công việc
	CString sluuTenCV = GIOR(iRow,_iCot[5],prDMCV,L"");
	CString sluuTenDTLM = GIOR(iRow,_iCot[28],prDMCV,L"");

	// Lưu đơn vị tính
	CString sluuDVTx = GIOR(iRow,_iCot[7],prDMCV,L"");

	// Lưu mã HSNK (nếu có)
	CString sluuMHSNT = GIOR(iRow,_iCot[4],prDMCV,L"");

	// Bổ sung lưu lại diễn giải (12.09.2017)
	jLuuDGKL = m_cbk3.GetCheck();
	int sldgiai=0, iXoa=0;;
	CString sMax=L"",sXoa=L"",sDGiai=L"";

	// Xóa công tác (17.06.2017)	
	sXoa= GIOR(iRow,_iCot[1],prDMCV,L"");
	if(_wtoi(sXoa)>0)
	{
		for (int i = iRow+1; i < iEnd; i++)
		{
			sXoa = GIOR(i,_iCot[1],prDMCV,L"");
			sMax = GIOR(i,_iCot[2],prDMCV,L"");
			
			if(_wtoi(sXoa)>0 || sMax==L"HM" || sMax==L"GD")
			{
				iXoa=i;
				break;
			}
		}

		if(iXoa>0)
		{
			// jLuuDGKL = 1 cho phép lưu diễn giải
			if(jLuuDGKL==1)
			{
				for (int i = iXoa-1; i >= iRow+1; i--)
				{
					sDGiai = GIOR(i,_iCot[5],prDMCV,L"");
					if(sDGiai!=L"")
					{
						sldgiai = i-iRow;
						break;
					}
				}						
			}

			// Không lưu diễn giải hoặc số lượng diễn giải = 0 --> chỉ xét đến số lượng tiêu chuẩn
			if(jLuuDGKL==0 || sldgiai==0)
			{
				if(iRow+1<iXoa)
				{
					PRS = psDMCV->GetRange(prDMCV->GetItem(iRow+1,1),prDMCV->GetItem(iXoa-1,1));
					PRS->EntireRow->Delete(xlShiftUp);
				}

				PRS = prDMCV->GetItem(iRow,1);
				PRS->EntireRow->ClearContents();

				// Chèn dòng
				if ((slMaTC > 1 && jNhomTC == 0) || (ChiaCongtac == 1 && slctcon > 0))
				{
					int icongvt = 0;
					if (slMaTC > 1 && jNhomTC == 0) icongvt = slMaTC - 1;
					if (ChiaCongtac == 1 && slctcon > icongvt) icongvt = slctcon;

					PRS = psDMCV->GetRange(prDMCV->GetItem(iRow + 1, 1), prDMCV->GetItem(iRow + icongvt, 1));
					PRS->EntireRow->Insert(xlShiftDown);

					PRS = psDMCV->GetRange(prDMCV->GetItem(iRow + 1, 1), prDMCV->GetItem(iRow + icongvt, 1));
					PRS->EntireRow->Interior->PutColor(xlNone);
				}				
			}
			else
			{
				// Bổ sung 12.09.2017
				PRS = psDMCV->GetRange(prDMCV->GetItem(iRow,_iCot[8]),prDMCV->GetItem(iXoa-1,_iCot[8]));
				PRS->ClearContents();

				PRS = psDMCV->GetRange(prDMCV->GetItem(iRow,_iCot[13]),prDMCV->GetItem(iXoa-1,_iCot[13]));
				PRS->ClearContents();

				int sltc = slMaTC;
				if(jNhomTC==1 || sltc==0) sltc=1;

				// Kiểm tra xem số lượng diễn giải có nhiều hơn tiêu chuẩn không?
				int imaxx = sltc-1;
				if(sldgiai>=sltc) imaxx = sldgiai;

				int jThem = iXoa-iRow-1;
				if(imaxx>jThem)
				{
					// Chèn dòng
					PRS = psDMCV->GetRange(prDMCV->GetItem(iXoa,1),prDMCV->GetItem(iXoa+(imaxx-jThem),1));
					PRS->EntireRow->Insert(xlShiftDown);

					PRS = psDMCV->GetRange(prDMCV->GetItem(iXoa,1),prDMCV->GetItem(iXoa+(imaxx-jThem),1));
					PRS->EntireRow->Interior->PutColor(xlNone);
				}
				else if(jThem>imaxx)
				{
					PRS = psDMCV->GetRange(prDMCV->GetItem(iRow+imaxx+1,1),prDMCV->GetItem(iXoa-1,1));
					PRS->EntireRow->Delete(xlShiftUp);
				}
			}
		}
	}
	else
	{
		// Chèn dòng
		if((slMaTC>1 && jNhomTC==0) || (ChiaCongtac==1 && slctcon>0))
		{
			int icongvt = 0;
			if (slMaTC > 1 && jNhomTC == 0) icongvt = slMaTC-1;
			if (ChiaCongtac == 1 && slctcon > icongvt) icongvt = slctcon;

			PRS = psDMCV->GetRange(prDMCV->GetItem(iRow+1,1),prDMCV->GetItem(iRow+ icongvt,1));
			PRS->EntireRow->Insert(xlShiftDown);

			PRS = psDMCV->GetRange(prDMCV->GetItem(iRow+1,1),prDMCV->GetItem(iRow+ icongvt,1));
			PRS->EntireRow->Interior->PutColor(xlNone);
		}
	}

////////////////////////////////////////////////////////////////////////////

	// Chèn danh mục công việc con (nếu có)
	if (ChiaCongtac == 1 && slctcon > 0)
	{
		for (int i = 1; i <= slctcon; i++)
		{
			prDMCV->PutItem(iRow + i, _iCotMaDM, (_bstr_t)MSCVC[i].szMahieu);
			prDMCV->PutItem(iRow + i, _iCot[3], (_bstr_t)L"CV");
			prDMCV->PutItem(iRow + i, _iCot[5], (_bstr_t)MSCVC[i].szNoidung);
			prDMCV->PutItem(iRow + i, _iCot[7], (_bstr_t)MSCVC[i].szDVT);
		}
	}

////////////////////////////////////////////////////////////////////////////

	// Chèn nội dung tiêu chuẩn
	if(slMaTC>0)
	{
		if(jNhomTC==0)
		{
			for (int i = 0; i < slMaTC; i++)
			{
				prDMCV->PutItem(iRow+i,_iCot[13],(_bstr_t)sKTMaTC[i+1]);
				prDMCV->PutItem(iRow+i,_iCot[8],(_bstr_t)sKTTenTC[i+1]);
				val = prDMCV->GetItem(iRow+i,_iCot[8]+1);
				if(val==L"") prDMCV->PutItem(iRow+i,_iCot[8]+1,(_bstr_t)L" ");
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

							// Số lượng tiêu chuẩn bằng nhau							
							if(icheck==slMaTC)
							{
								// Kiểm tra tiếp dòng dưới có phải tiêu chuẩn không và nó có nằm trong nhóm khác không?
								szval = GIOR(i+slMaTC,icTC,prBTC,L"");
								if(szval!=L"")
								{
									szval = GIOR(i+slMaTC,icTen,prBTC,L"");
									if(szval!=L"")
									{
										// Đã tồn tại tên nhóm tiêu chuẩn
										iOk=1;
										break;
									}
								}
								else
								{
									// Đã tồn tại tên nhóm tiêu chuẩn
									iOk=1;
									break;
								}
							}
						}
					}
				}

				if(iOk==1)
				{
					szval.Format(L"='%s'!%s",(LPCTSTR)shBTC,GAR(ivtr,icTen,prBTC,3));
					prDMCV->PutItem(iRow,_iCot[8],(_bstr_t)szval);
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
					prDMCV->PutItem(iRow,_iCot[8],(_bstr_t)szval);

					for (int i = 1; i <= slMaTC; i++)
					{
						prBTC->PutItem(ivtr+i,icMa,(_bstr_t)sKTMaTC[i]);
						prBTC->PutItem(ivtr+i,icTC,(_bstr_t)sKTTenTC[i]);
					}
				}
			}
			else
			{
				szval.Format(L"='%s'!%s",(LPCTSTR)shBTC,GAR(vtnhom,icTen,prBTC,3));
				prDMCV->PutItem(iRow,_iCot[8],(_bstr_t)szval);
			}		
		}		
	}
	else
	{
		if(vtnhom>0)
		{
			szval.Format(L"='%s'!%s",(LPCTSTR)shBTC,GAR(vtnhom,icTen,prBTC,3));
			prDMCV->PutItem(iRow,_iCot[8],(_bstr_t)szval);
		}
	}

// ********************** Đổ dữ liệu ************************************************
	if(fla==1)
	{
		prDMCV->PutItem(iRow,_iCot[2],(_bstr_t)sMSCV);

		// Bổ sung 22.04.2019
		szfDMNhancong=L"", szfTNMay=L"";
		_getDinhmucNhancong(sMSCV,sLLink);

		if (ChiaCongtac == 1 && slctcon > 0)
		{
			for (int z = 1; z <= slctcon; z++)
			{
				szconDMNC[z].Trim();
				if (szconDMNC[z] == L"") szconDMNC[z] = L"0";
				prDMCV->PutItem(iRow+z, _iCot[36], (_bstr_t)szconDMNC[z]);		// Định mức nhân công  <-- Lấy trong CSDL
				prDMCV->PutItem(iRow+z, _iCot[35], (_bstr_t)szconTNMay[z]);		// Tài nguyên máy

				prDMCV->PutItem(iRow+z, _iCot[37], 1);							// Hệ số năng suất
				prDMCV->PutItem(iRow+z, _iCot[30], 1);							// Tổng thời gian thực hiện

				CString szTongcong = L"";
				szTongcong.Format(L"=%s*%s*%s",
					GAR(iRow+z, _iCot[6], prDMCV, 0), GAR(iRow+z, _iCot[36], prDMCV, 0), GAR(iRow+z, _iCot[37], prDMCV, 0));
				prDMCV->PutItem(iRow+z, _iCot[38], (_bstr_t)szTongcong);		// Tổng công

				CString szNCNgay = L"";
				szNCNgay.Format(L"=IFERROR(ROUNDUP(%s/%s,0),0)", GAR(iRow+z, _iCot[38], prDMCV, 0), GAR(iRow+z, _iCot[30], prDMCV, 0));
				prDMCV->PutItem(iRow+z, _iCot[39], (_bstr_t)szNCNgay);			// Nhân công trong ngày
			}
		}
		else
		{
			szfDMNhancong.Trim();
			if (slctcon > 1 || szfDMNhancong == L"") szfDMNhancong = L"0";
			prDMCV->PutItem(iRow, _iCot[36], (_bstr_t)szfDMNhancong);		// Định mức nhân công  <-- Lấy trong CSDL
			prDMCV->PutItem(iRow, _iCot[35], (_bstr_t)szfTNMay);			// Tài nguyên máy

			prDMCV->PutItem(iRow, _iCot[37], 1);							// Hệ số năng suất
			prDMCV->PutItem(iRow, _iCot[30], 1);							// Tổng thời gian thực hiện

			CString szTongcong = L"";
			szTongcong.Format(L"=%s*%s*%s",
				GAR(iRow, _iCot[6], prDMCV, 0), GAR(iRow, _iCot[36], prDMCV, 0), GAR(iRow, _iCot[37], prDMCV, 0));
			prDMCV->PutItem(iRow, _iCot[38], (_bstr_t)szTongcong);			// Tổng công

			CString szNCNgay = L"";
			szNCNgay.Format(L"=IFERROR(ROUNDUP(%s/%s,0),0)", GAR(iRow, _iCot[38], prDMCV, 0), GAR(iRow, _iCot[30], prDMCV, 0));
			prDMCV->PutItem(iRow, _iCot[39], (_bstr_t)szNCNgay);			// Nhân công trong ngày
		}		

		if(sluuTenCV!=L"")
		{
			int iktten = _wtoi(getVCell(psTS,L"CF_TRATEN"));
			if(iktten==1) prDMCV->PutItem(iRow,_iCot[5],(_bstr_t)sNA);
			else
			{
				prDMCV->PutItem(iRow,_iCot[5],(_bstr_t)sluuTenCV);
				prDMCV->PutItem(iRow,_iCot[28],(_bstr_t)sluuTenDTLM);
			}
		}
		else prDMCV->PutItem(iRow,_iCot[5],(_bstr_t)sNA);

		if (ChiaCongtac != 1 || slctcon == 0)
		{
			if (sluuDVTx != L"") prDMCV->PutItem(iRow, _iCot[7], (_bstr_t)sluuDVTx);
			else prDMCV->PutItem(iRow, _iCot[7], (_bstr_t)sDVT);
		}		

		// Copy ngày tháng
		int iktrNT=0;
		for (int j = iRow-1; j >= 7; j--)
		{
			val = GIOR(j,1,prDMCV,L"");
			if(_wtoi(val)>0)
			{
				if(iCpy==0 || iCpy==1)
				{
					// ngày
					int gtr[10];
					gtr[0] = getColumn(psDMCV,"DMBB_TN_NGAY");
					gtr[1] = getColumn(psDMCV,"DMBB_TN_KQ");
					gtr[2] = getColumn(psDMCV,"DMBB_NB_NGAY");
					gtr[3] = getColumn(psDMCV,"DMBB_PHIEUYC");
					gtr[4] = getColumn(psDMCV,"DMBB_AB_NGAY");
					gtr[5] = getColumn(psDMCV,"DMBB_HK_NGAYBD");
					gtr[6] = getColumn(psDMCV,"DMBB_HK_NGAYKT");
					gtr[7] = getColumn(psDMCV,"DMBB_DKDBT_NGAY");
					gtr[8] = getColumn(psDMCV,"DMBB_TDDBT_NGAY");

					for (int k = 0; k < 9; k++)
					{
						szval = GIOR(j,gtr[k],prDMCV,L"");
						if (ChiaCongtac == 1 && slctcon > 0)
						{
							for (int z = 1; z <= slctcon; z++)
							{
								f_ktra_date(szval, prDMCV, iRow + z, gtr[k]);
							}
						}
						else
						{
							f_ktra_date(szval, prDMCV, iRow, gtr[k]);
						}
					}
				}

				if(iCpy==0 || iCpy==2)
				{
					// giờ
					int gtr[6];
					gtr[0] = getColumn(psDMCV,"DMBB_TN_GIO");
					gtr[1] = getColumn(psDMCV,"DMBB_NB_GIO");
					gtr[2] = getColumn(psDMCV,"DMBB_AB_GIO");
					gtr[3] = getColumn(psDMCV,"DMBB_DKDBT_GIO");
					gtr[4] = getColumn(psDMCV,"DMBB_TDDBT_GIO");
					gtr[5] = getColumn(psDMCV,"DMBB_PHIEUYC_GIO");

					for (int k = 0; k < 6; k++)
					{
						szval = GIOR(j,gtr[k],prDMCV,L"");						

						if (ChiaCongtac == 1 && slctcon > 0)
						{
							for (int z = 1; z <= slctcon; z++)
							{
								prDMCV->PutItem(iRow+z, gtr[k], (_bstr_t)szval);
							}
						}
						else
						{
							prDMCV->PutItem(iRow, gtr[k], (_bstr_t)szval);
						}
					}
				}

				iktrNT=1;
				break;
			}
		}

		// iktrNT=0 --> Trường hợp không tìm thấy ngày sao chép
		if (iktrNT == 0 || (iCpy != 0 && iCpy != 1))
		{
			// ngày
			int gtr[10];
			gtr[0] = getColumn(psDMCV,"DMBB_TN_NGAY");
			gtr[1] = getColumn(psDMCV,"DMBB_TN_KQ");
			gtr[2] = getColumn(psDMCV,"DMBB_NB_NGAY");
			gtr[3] = getColumn(psDMCV,"DMBB_PHIEUYC");
			gtr[4] = getColumn(psDMCV,"DMBB_AB_NGAY");
			gtr[5] = getColumn(psDMCV,"DMBB_HK_NGAYBD");
			gtr[6] = getColumn(psDMCV,"DMBB_HK_NGAYKT");
			gtr[7] = getColumn(psDMCV,"DMBB_DKDBT_NGAY");
			gtr[8] = getColumn(psDMCV,"DMBB_TDDBT_NGAY");

			// Ngày bắt đầu
			prDMCV->PutItem(iRow,gtr[5],(_bstr_t)L"=NOW()");
			val = GIOR(iRow,gtr[5],prDMCV,L"");
			int pos2= val.Find(L" ");
			if(pos2>0) val = val.Left(pos2);
			val.Trim();			

			if (ChiaCongtac == 1 && slctcon > 0)
			{
				for (int z = 1; z <= slctcon; z++)
				{
					f_ktra_date(val, prDMCV, iRow+z, gtr[5]);
				}
			}
			else
			{
				f_ktra_date(val, prDMCV, iRow, gtr[5]);
			}

			// Ngày kết thúc
			val.Format(L"=%s",GAR(iRow,gtr[5],prDMCV,0));
			prDMCV->PutItem(iRow,gtr[6],(_bstr_t)val);

			if (ChiaCongtac == 1 && slctcon > 0)
			{
				for (int z = 1; z <= slctcon; z++)
				{
					prDMCV->PutItem(iRow+z, gtr[6], (_bstr_t)val);
				}
			}
			else
			{
				prDMCV->PutItem(iRow, gtr[6], (_bstr_t)val);
			}

			// NT Nội bộ + Yêu cầu
			val.Format(L"=%s",GAR(iRow,gtr[6],prDMCV,0));
			prDMCV->PutItem(iRow,gtr[2],(_bstr_t)val);
			prDMCV->PutItem(iRow,gtr[3],(_bstr_t)val);

			if (ChiaCongtac == 1 && slctcon > 0)
			{
				for (int z = 1; z <= slctcon; z++)
				{
					prDMCV->PutItem(iRow+z, gtr[2], (_bstr_t)val);
					prDMCV->PutItem(iRow+z, gtr[3], (_bstr_t)val);
				}
			}
			else
			{
				prDMCV->PutItem(iRow, gtr[2], (_bstr_t)val);
				prDMCV->PutItem(iRow, gtr[3], (_bstr_t)val);
			}

			// Nghiệm thu
			val.Format(L"=%s+1",GAR(iRow,gtr[3],prDMCV,0));
			prDMCV->PutItem(iRow,gtr[4],(_bstr_t)val);

			if (ChiaCongtac == 1 && slctcon > 0)
			{
				for (int z = 1; z <= slctcon; z++)
				{
					prDMCV->PutItem(iRow+z, gtr[4], (_bstr_t)val);
				}
			}
			else
			{
				prDMCV->PutItem(iRow, gtr[4], (_bstr_t)val);
			}
		}

		val = prDMCV->GetItem(iRow,_iCot[8]+1);
		if(val==L"") prDMCV->PutItem(iRow,_iCot[8]+1,(_bstr_t)L" ");		

		PRS = prDMCV->GetItem(iRow,_iCot[5]);
		PRS->PutWrapText(TRUE);

		if (ChiaCongtac == 1 && slctcon > 0)
		{
			PRS = psDMCV->GetRange(prDMCV->GetItem(iRow+1, _iCot[5]), prDMCV->GetItem(iRow+ slctcon, _iCot[5]));
			PRS->PutWrapText(TRUE);
		}

		PRS = psDMCV->GetRange(prDMCV->GetItem(iRow,1),prDMCV->GetItem(iRow, _iCotEND));
		if(chkClrCV==1) PRS->Interior->PutColor(fclrCV);
		else PRS->EntireRow->Interior->PutColor(xlNone);
	}
	
	// Định dạng khung
	// Kiểm tra công tác kế trên có phải hạng mục hay giai đoạn không?
	val = GIOR(iRow-1,_iCot[2],prDMCV,L"");
	val.MakeUpper();
	val.Trim();

	int iktx=slMaTC;
	if(jNhomTC==1 || slMaTC==0) iktx=1; 
	PRS = psDMCV->GetRange(prDMCV->GetItem(iRow,_iCot[1]),prDMCV->GetItem(iRow+iktx+cong, _iCotEND));
	
	if(val==L"HM" || val==L"GD") PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->Weight=xlThin;
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

	if(!m_btn[0].GetCheck())
	{
		if (ChiaCongtac == 1 && slctcon > 0)
		{
			for (int z = 1; z <= slctcon; z++)
			{
				//bChkPS = 0;
				if (m_btn[1].GetCheck()) prDMCV->PutItem(iRow+z, _iCot[21], (_bstr_t)"T");	//nRaPS = 1;
				else prDMCV->PutItem(iRow+z, _iCot[21], (_bstr_t)"N");						//nRaPS = 2;
			}
		}
		else
		{
			//bChkPS = 0;
			if (m_btn[1].GetCheck()) prDMCV->PutItem(iRow, _iCot[21], (_bstr_t)"T");	//nRaPS = 1;
			else prDMCV->PutItem(iRow, _iCot[21], (_bstr_t)"N");						//nRaPS = 2;
		}		
	}
	else
	{
		prDMCV->PutItem(iRow, _iCot[21], (_bstr_t)L"");
		if (ChiaCongtac == 1 && slctcon > 0)
		{
			for (int z = 1; z <= slctcon; z++)
			{
				prDMCV->PutItem(iRow+z, _iCot[21], (_bstr_t)L"");
			}
		}
	} 

	// Bổ sung nhóm ký biên bản (31.05.2018)
	CheckNhomKy(psDMCV,iRow,getColumn(psDMCV,L"DMBB_KBB"));

	// Autofit
	PRS = prDMCV->GetItem(iRow,1);
	PRS->EntireRow->Rows->AutoFit();

	if (ChiaCongtac == 1 && slctcon > 0)
	{
		PRS = psDMCV->GetRange(prDMCV->GetItem(iRow + 1, _iCot[5]), prDMCV->GetItem(iRow + slctcon, _iCot[5]));
		PRS->EntireRow->Rows->AutoFit();
	}

	// Sắp xếp lại stt
	iEnd = FindComment(psDMCV,_iCot[1],"END");
	int ntype = _wtoi(getVCell(psTS,L"TS_ATOSTT"));
	if(ntype!=0 && ntype!=1 && ntype!=2) ntype=0;
	f_Auto_stt_dmuc(prDMCV,ntype,0,8,iEnd,_iCot[1],_iCot[2],_iCot[4]);

	if(sluuMHSNT!=L"") prDMCV->PutItem(iRow,_iCot[4],(_bstr_t)sluuMHSNT);
	else
	{
		CString mm = GIOR(iRow,_iCot[1],prDMCV,L"");
		CString qq = mm;
		if(_wtoi(mm)<10) mm.Format(L"'0%s",qq);
		else mm.Format(L"'%s",qq);
		prDMCV->PutItem(iRow,_iCot[4],(_bstr_t)mm);
	}

	// Leo add 07.02.2018
	if(slMaTC>1 && jNhomTC==0)
	{
		xl->ActiveWindow->PutScrollRow(iRow-5);
		PRS = prDMCV->GetItem(iRow+slMaTC-1,_iCot[2]);
		PRS->Select();
	}

	xl->PutScreenUpdating(VARIANT_TRUE);
	CDialog::OnOK();
	


	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	Dlg_open_progress p;
	p.DoModal();

	}
	catch(...){_s(L"Lỗi tra cứu dữ liệu công việc.");}
}


int Dlg_Open_DMCV::f_Get_change()
{
	try
	{
		CString val=L"";
		int ktr=0,nCount = lit_tccv.GetItemCount();
		if(_isl1!=nCount) return 1;
		if(nCount>0)
		{
			for (int i = 0; i < nCount; i++)
			{
				val = lit_tccv.GetItemText(i,0);
				if(val!=L"") return 1;
			}
		}

		nCount = lit_ndkt.GetItemCount();
		if(_isl2!=nCount) return 1;
		if(nCount>0)
		{
			for (int i = 0; i < nCount; i++)
			{
				val = lit_ndkt.GetItemText(i,0);
				if(val!=L"") return 1;
			}
		}

		return 0;
	}
	catch(...){}

	return 0;
}


void Dlg_Open_DMCV::OnNMClickL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	ClickEsc=0;
	iTCclick=-1;
	sDownTC=L"",sFullTC=L"";

	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	int nItem = pNMListView->iItem;
	int nSubItem = pNMListView->iSubItem;

	iLuuTCCV = m_chkLuuTC.GetCheck();
	CString szval=L"",szval2=L"",szval3=L"";
	if(nItem>=0)
	{
		if(nSubItem==0)
		{
			int _pos=-1;
			szval = lit_dmcv.GetItemText(nItem,1);
			CString szcong = szval+L"|";

			szval2 = lit_dmcv.GetItemText(nItem,2);
			szval2.Replace(L"|",L" ");
			if(szval2==L"") szval2=L"<Công tác chưa đặt tên>";

			szval3 = lit_dmcv.GetItemText(nItem,3);
			if(szval3==L"") szval3=L"<đvt>";

			if(sLuuMaCT!=L"")
			{
				 if(sLuuMaCT.Right(1)!=L"|") sLuuMaCT+=L"|";
				 if(sLuuTenCT.Right(1)!=L"|") sLuuTenCT+=L"|";
				 if(sLuuDVT.Right(1)!=L"|") sLuuDVT+=L"|";
				 if(sLuuLink.Right(1)!=L"|") sLuuLink+=L"|";
				_pos = sLuuMaCT.Find(szcong);
			}

			if((int)lit_dmcv.GetCheck(nItem)==0)
			{
				// Tích chọn
				lit_dmcv.SetRowColors(nItem,RGB(255,255,150),BLACK);

				if(_pos==-1 && iLuuTCCV==1)
				{
					sLuuMaCT+=szval;
					m_txtTChon.SetWindowText(sLuuMaCT);

					sLuuTenCT+=szval2;
					sLuuDVT+=szval3;
					sLuuLink+=sFullPath;
				}								
			}
			else
			{
				// Bỏ tích
				lit_dmcv.SetRowColors(nItem,WHITE,BLACK);

				if(_pos>=0 && iLuuTCCV==1)
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
						CString szgg[5]={L"",sLuuMaCT,sLuuTenCT,sLuuDVT,sLuuLink};
						for (int a = 1; a <= 4; a++)
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
						sLuuDVT = szgg[3];
						sLuuLink = szgg[4];

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
						sLuuDVT= L"";
						sLuuLink= L"";
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
			szval = L"Dữ liệu chỉnh sửa nội dung & phương pháp nghiệm thu chưa được lưu."
				L"\nBạn vẫn muốn xem công việc khác mà không lưu?";
			int n0 = _yn(szval,0);
			if(n0==7)
			{
				lit_dmcv.SetItemState(nItem,0,LVIS_SELECTED);
				lit_dmcv.SetItemState(nClick,LVIS_SELECTED,LVIS_SELECTED);
				return;
			}
		}
		
		// Xóa toàn bộ dữ liệu trên List Control
		lit_tccv.DeleteAllItems();
		ASSERT(lit_tccv.GetItemCount() == 0);

		lit_ndkt.DeleteAllItems();
		ASSERT(lit_ndkt.GetItemCount() == 0);

		nClick = nItem;
		CString sMacv = lit_dmcv.GetItemText(nClick,1);
		
		ObjConn2.OpenAccessDB(pathMDB,L"",L"");

		// Tìm kiếm tiêu chuẩn
		f_tk_tchuan(sMacv);

		// Tìm kiếm nội dung công việc
		f_tk_ndung(sMacv);

		ObjConn2.CloseDatabase();
	}

	*pResult = 0;
}


void Dlg_Open_DMCV::f_tk_tchuan(CString sVal)
{
	TRY
	{
		SqlString.Format(L"SELECT * FROM tbMaCV_Tieuchuan WHERE MaCV LIKE '%s' ORDER BY ID ASC;",sVal);		
		CRecordset* Recset = ObjConn2.Execute(SqlString);
		
		int dem=0;
		CString sGan=L"";
		
		while( !Recset->IsEOF() )
		{
			Recset->GetFieldValue(L"MaTC",sGan);
			sGan.Trim();
			if(sGan!=L"")
			{
				lit_tccv.InsertItem(dem,L"",0);
				lit_tccv.SetItemText(dem,1,sGan);
				lit_tccv.SetItemText(dem,2,f_ten_tchuan(sGan));
				dem++;
			}

			Recset->MoveNext();
		}

		Recset->Close();
		_isl1 = dem;
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[8] Lỗi tìm kiếm dữ liệu tiêu chuẩn công việc")+e->m_strError);
	}
	END_CATCH;
}

void Dlg_Open_DMCV::_getDinhmucNhancong(CString szMaCV, CString szLLk)
{
	try
	{
		CRecordset* Recset;
		CRecordset* RCS2;

		CString szMaDM=L"",szMSVT=L"",szHPhi=L"",szTenmay=L"";
		CString szTruyvan=L"";
		CString szMa[50],szLink[50];

		szMaCV.Replace(L"\n",L"");
		_TackTukhoaChuan(szMaCV,L"|",L" ",0,0);
		for (int i = 1; i <= iKey; i++) szMa[i] = pKey[i];

		szLLk.Replace(L"\n",L"");
		_TackTukhoaChuan(szLLk,L"|",L";",0,0);
		for (int i = 1; i <= iKey; i++) szLink[i] = pKey[i];

		for (int i = 1; i <= iKey; i++)
		{
			szconDMNC[i] = L"";
			szconTNMay[i] = L"";
			
			// Xác định định mức nhân công
			// Xác định mã định mức
			if(_FileExists(szLink[i],0)==1)
			{
				szMaDM=L"";
				ObjConn.OpenAccessDB(szLink[i],L"",L"");
				SqlString.Format(L"SELECT * FROM tbDonGia WHERE MHDG LIKE '%s';",szMa[i]);
				Recset = ObjConn.Execute(SqlString);
				while( !Recset->IsEOF() )
				{
					Recset->GetFieldValue(L"MHDM",szMaDM);
					break;
				}
				Recset->Close();

				if(szMaDM!=L"")
				{
					szMSVT = L"", szHPhi = L"", szTenmay = L"";
					SqlString.Format(L"SELECT * FROM tbDinhMuc WHERE MHDM LIKE '%s';",szMaDM);
					Recset = ObjConn.Execute(SqlString);
					while( !Recset->IsEOF() )
					{
						Recset->GetFieldValue(L"MSVT",szMSVT);						
						if(szMSVT.Left(1)==L"N")
						{
							Recset->GetFieldValue(L"HPVT",szHPhi);
							szHPhi.Replace(L",",L".");
							szHPhi.Trim();
							if(szHPhi==L"") szHPhi=L"0";

							szconDMNC[i] = szHPhi;
							szconDMNC[i] += L"+";

							szfDMNhancong+=szHPhi;
							szfDMNhancong+=L"+";
						}
						else if(szMSVT.Left(1)==L"M")
						{
							// Tìm máy trong từ điển và giá ca máy
							SqlString.Format(L"SELECT * FROM tbTuDienVatTu WHERE MSVT LIKE '%s';",szMSVT);
							RCS2 = ObjConn.Execute(SqlString);
							while( !RCS2->IsEOF() )
							{
								RCS2->GetFieldValue(L"TVT",szTenmay);
								break;
							}
							RCS2->Close();

							if(szTenmay==L"")
							{
								SqlString.Format(L"SELECT * FROM tbGiaCaMay WHERE MH LIKE '%s';",szMSVT);
								RCS2 = ObjConn.Execute(SqlString);
								while( !RCS2->IsEOF() )
								{
									RCS2->GetFieldValue(L"LM_TB",szTenmay);
									break;
								}
								RCS2->Close();
							}

							if(szTenmay!=L"" && szTenmay!=L"Máy khác")
							{
								szTenmay.Replace(L";",L",");
								szTenmay+=L" [1];";

								szconTNMay[i] += szTenmay;
								szfTNMay+=szTenmay;
							}
						}

						Recset->MoveNext();
					}
					Recset->Close();
				}

				ObjConn.CloseDatabase();
			}
		}

		CString szval=L"";
		if(szfDMNhancong!=L"")
		{
			szval.Format(L"=%s",szfDMNhancong.Left(szfDMNhancong.GetLength()-1));
			szfDMNhancong = szval;
			if(szfDMNhancong.Find(L"+")==-1) szfDMNhancong.Replace(L"=",L"");
		}

		if(szfTNMay!=L"")
		{
			szval.Format(L"'%s",szfTNMay);
			szfTNMay = szval;
		}

		for (int i = 1; i <= iKey; i++)
		{
			if (szconDMNC[i] != L"")
			{
				szval.Format(L"=%s", szconDMNC[i].Left(szconDMNC[i].GetLength() - 1));
				szconDMNC[i] = szval;
				if (szconDMNC[i].Find(L"+") == -1) szconDMNC[i].Replace(L"=", L"");
			}

			if (szconTNMay[i] != L"")
			{
				szval.Format(L"'%s", szconTNMay[i]);
				szconTNMay[i] = szval;
			}
		}
	}
	catch(...){}
}


CString Dlg_Open_DMCV::f_ten_tchuan(CString sVal)
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


void Dlg_Open_DMCV::f_tk_ndung(CString sVal)
{
	TRY
	{
		SqlString.Format(L"SELECT * FROM tbMaCV_NDnghiemthu WHERE MaCV LIKE '%s' ORDER BY ID ASC;",sVal);
		CRecordset* Recset = ObjConn2.Execute(SqlString);
	
		int dem=0;
		CString sGan[3];
		while( !Recset->IsEOF() )
		{
			Recset->GetFieldValue(L"NoidungKiemtra",sGan[1]);
			sGan[1].Trim();
			
			if(sGan[1]!=L"")
			{
				Recset->GetFieldValue(L"PhuongphapKiemtra",sGan[2]);
				sGan[2].Trim();

				lit_ndkt.InsertItem(dem,L"",0);
				lit_ndkt.SetItemText(dem,1,sGan[1]);
				lit_ndkt.SetItemText(dem,2,sGan[2]);
				dem++;
			}

			Recset->MoveNext();
		}
		Recset->Close();
		_isl2 = dem;
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[10] Lỗi cập nhật dữ liệu nội dung công việc")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_Open_DMCV::OnEnKillfocusEd1()
{
	try
	{
		sBefore = lit_dmcv.GetItemText(CLRow,CLCol);
		sBefore.Replace(L"'",L"");

		txtE_DMCV.GetWindowTextW(sAfter);
		sAfter.Replace(L"'",L"");
		if(sAfter==sBefore) return;

		ObjConn.OpenAccessDB(sFullPath,L"",L"");

		if(CLCol==2 || CLCol==3)
		{
			// Kiểm tra mã đó có trong CSDL chưa? 
			// Chưa có thì thêm mới vào
			CString sdem=L"";
			CString sId = lit_dmcv.GetItemText(CLRow,1);		
			SqlString.Format(L"SELECT COUNT(*) AS iDem FROM tbDonGia WHERE MHDG LIKE '%s';",sId);
			CRecordset*	Recset = ObjConn.Execute(SqlString);
			Recset->GetFieldValue(L"iDem",sdem);
			Recset->Close();

			if(_wtoi(sdem)>0)
			{
				if(CLCol==2) sdem = L"TCV";
				else sdem=L"DVT";

				SqlString.Format(L"UPDATE tbDonGia SET %s = '%s' WHERE MHDG LIKE '%s';",sdem,sAfter,sId);
				ObjConn.ExecuteRB(SqlString);
			}
		}
		else if(CLCol==1)
		{
			SqlString.Format(L"UPDATE tbDonGia SET MHDG = '%s' WHERE MHDG LIKE '%s';",sAfter,sBefore);
			ObjConn.ExecuteRB(SqlString);
		}

		ObjConn.CloseDatabase();
	}
	catch(...){}
}


void Dlg_Open_DMCV::OnBnClickedB4()
{
	TRY
	{
		ClickEsc=0;
		int nItem=-1;
		int nCount = lit_dmcv.GetItemCount();

		POSITION pos = lit_dmcv.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			for (int i = 0; i < nCount; i++)
			if (lit_dmcv.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
			{
				nItem=i;
				nClick++;
				break;
			}
		}
		if(nItem==-1) nItem = nCount;

		// Số ngẫu nhiên theo thời gian
		CString val=L"";
		time_t now = time(0);
		tm *ltm = localtime(&now);
		val.Format(L"%d%d",ltm->tm_min,ltm->tm_sec);
		int _nRand = _wtoi(val);

		// Kiểm tra dữ liệu có trong CSDL không?
		CString sTvan=L"",sDem=L"",jNew=L"";
		jNew.Format(L"NW.%d%d%d",_nRand,rand()%10,rand()%10);
		sTvan.Format(L"SELECT COUNT(*) AS iDem FROM tbDonGia WHERE MHDG LIKE '%s';",jNew);
		
		// Kết nối SQL
		ObjConn.OpenAccessDB(sFullPath,L"",L"");
		CRecordset* Recset = ObjConn.Execute(sTvan);
		Recset->GetFieldValue(L"iDem",sDem);
		Recset->Close();

		while (_wtoi(sDem)>0)
		{
			sDem=L"";
			jNew.Format(L"NW.%d%d%d",_nRand,rand()%10,rand()%10);
			sTvan.Format(L"SELECT COUNT(*) AS iDem FROM tbDonGia WHERE MHDG LIKE '%s';",jNew);
			Recset = ObjConn.Execute(sTvan);
			Recset->GetFieldValue(L"iDem",sDem);
			Recset->Close();
		}

		// Đổ dữ liệu vào list
		sDem = L"Nhập tên công việc mới";
		lit_dmcv.InsertItem(nItem,L"",0);
		lit_dmcv.SetItemText(nItem,1,jNew);
		lit_dmcv.SetItemText(nItem,2,sDem);
		lit_dmcv.SetRowColors(nItem,RGB(0,128,0),RGB(255,255,255));

		// Lưu vào trong CSDL
		SqlString.Format(L"INSERT INTO tbDonGia (MHDG,TCV,TK) VALUES ('%s','%s','NEW');",jNew,sDem);
		ObjConn.ExecuteRB(SqlString);

		ObjConn.CloseDatabase();
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[11] Lỗi thêm mới dữ liệu nội dung công việc")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_Open_DMCV::OnBnClickedB5()
{
	TRY
	{
		ClickEsc=0;
		int nCount = lit_dmcv.GetItemCount();
		if(nCount<=0) return;

		POSITION pos=lit_dmcv.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			int dem=0;
			for (int i = 0; i < nCount; i++)
				if (lit_dmcv.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED) dem++;

			CString sTvan,val = L"";
			if(dem==1) val = L"Bạn có chắc chắn sẽ xóa công việc này trong dữ liệu?";
			else val = L"Bạn có chắc chắn sẽ xóa những công việc này trong dữ liệu?";
			
			if(_yn(val,0) == 6)
			{
				// Kết nối SQL
				ObjConn.OpenAccessDB(sFullPath,L"",L"");

				for (int i = nCount-1; i >=0; i--)
				{
					if (lit_dmcv.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
					{
						// Xóa trong CSDL
						val = lit_dmcv.GetItemText(i,1);
						sTvan.Format(L"DELETE FROM tbDonGia WHERE MHDG LIKE '%s';",val);
						ObjConn.ExecuteRB(sTvan);

						// Xóa trên list control
						lit_dmcv.DeleteItem(i);
					}
				}

				ObjConn.CloseDatabase();
			}
		}
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[28] Lỗi xóa dữ liệu công việc")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_Open_DMCV::OnBnClickedB6()
{
	ClickEsc=0;
	f_SaveWidth();
	f_Get_size();
	CDialog::OnCancel();
}


void Dlg_Open_DMCV::OnNMClickL2(NMHDR *pNMHDR, LRESULT *pResult)
{
	ClickEsc=0;
	_iTab=1;

	sDownTC=L"",sFullTC=L"";
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	iTCclick = pNMListView->iItem;
	if(iTCclick>=0)
	{
		sDownTC = lit_tccv.GetItemText(iTCclick,1);
		sFullTC = lit_tccv.GetItemText(iTCclick,2);
	}
}


void Dlg_Open_DMCV::OnNMClickL3(NMHDR *pNMHDR, LRESULT *pResult)
{
	ClickEsc=0;
	_iTab=2;
}


void Dlg_Open_DMCV::OnBnClickedB7()
{
	try
	{
		ClickEsc=0;
		int nCount=0,nItem=-1;
		if(_iTab==0) return;
		else if(_iTab==1)
		{
			// Thêm mới tiêu chuẩn
			nCount = lit_tccv.GetItemCount();
			POSITION pos=lit_tccv.GetFirstSelectedItemPosition();
			if(pos!=NULL)
			{
				for (int i = 0; i < nCount; i++)
				if (lit_tccv.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					nItem=i;
					break;
				}
			}

			if(nItem==-1) nItem = nCount;

			// Đổ dữ liệu vào list
			lit_tccv.InsertItem(nItem,L"",0);
			lit_tccv.SetItemText(nItem,2,L"Gõ tìm kiếm tiêu chuẩn");
			lit_tccv.SetRowColors(nItem,RGB(0,128,0),RGB(255,255,255));


		}
		else if(_iTab==2)
		{
			// Thêm mới phương pháp
			nCount = lit_ndkt.GetItemCount();
			POSITION pos=lit_ndkt.GetFirstSelectedItemPosition();
			if(pos!=NULL)
			{
				for (int i = 0; i < nCount; i++)
				if (lit_ndkt.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					nItem=i;
					break;
				}
			}

			if(nItem==-1) nItem = nCount;

			// Đổ dữ liệu vào list
			lit_ndkt.InsertItem(nItem,L"",0);
			lit_ndkt.SetItemText(nItem,1,L"Nhập nội dung kiểm tra");
			lit_ndkt.SetItemText(nItem,2,L"Nhập phương pháp kiểm tra");
			lit_ndkt.SetRowColors(nItem,RGB(0,128,0),RGB(255,255,255));
		}
	}
	catch(...){_s(L"Lỗi thêm mới nội dung & phương pháp nghiệm thu.");}
}


void Dlg_Open_DMCV::OnBnClickedB8()
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
			nCount = lit_tccv.GetItemCount();
			POSITION pos=lit_tccv.GetFirstSelectedItemPosition();
			if(pos!=NULL)
			{
				if(_yn(sMsg,0)==6)
				{
					for (int i = nCount-1; i >= 0; i--)
						if (lit_tccv.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED) lit_tccv.DeleteItem(i);
				}				
			}
		}
		else if(_iTab==2)
		{
			// Xóa phương pháp
			nCount = lit_ndkt.GetItemCount();
			POSITION pos=lit_ndkt.GetFirstSelectedItemPosition();
			if(pos!=NULL)
			{
				if(_yn(sMsg,0)==6)
				{
					for (int i = nCount-1; i >= 0; i--)
						if (lit_ndkt.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED) lit_ndkt.DeleteItem(i);
				}
			}
		}
	}
	catch(...){_s(L"Lỗi xóa dữ liệu nội dung & phương pháp nghiệm thu.");}
}


void Dlg_Open_DMCV::OnBnClickedB9()
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


void Dlg_Open_DMCV::_UpdateDL(int iItemR, int itype)
{
	try
	{
		CString sKtra=L"",sTvan=L"",sDem=L"",sGan[3];
		CString val = lit_dmcv.GetItemText(iItemR,1);

		// Tạo mã công việc mới
		int iRes = 0;
		if(itype==1)
		{
			sKtra=L"Mã công việc đã tồn tại. Bạn có muốn tự động mã mới?"
			L"\nChọn 'Yes' để tạo mã mới hoặc 'No' để thay thế dữ liệu cũ. \nHoặc nhấn Cancel để hủy cập nhật.";
			iRes = _yn(sKtra,1);
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
			if(val.GetLength()>8) sGn = val.Left(8);

			// Kiểm tra dữ liệu có trong CSDL không?
			jNew.Format(L"%s.%d%d",sGn,_nRand,rand()%10);
			sTvan.Format(L"SELECT COUNT(*) AS iDem FROM tbDonGia WHERE MHDG LIKE '%s';",jNew);
			CRecordset* Recset = ObjConn.Execute(sTvan);
			Recset->GetFieldValue(L"iDem",sDem);
			Recset->Close();

			while (_wtoi(sDem)>0)
			{
				sDem=L"";
				jNew.Format(L"%s.%d%d",sGn,_nRand,rand()%10);
				sTvan.Format(L"SELECT COUNT(*) AS iDem FROM tbDonGia WHERE MHDG LIKE '%s';",jNew);
				Recset = ObjConn.Execute(sTvan);
				Recset->GetFieldValue(L"iDem",sDem);
				Recset->Close();
			}

			val = jNew;

			// Gán lại mã công việc
			sGan[1] = lit_dmcv.GetItemText(iItemR,2);
			sGan[2] = lit_dmcv.GetItemText(iItemR,3);
			sTvan.Format(L"INSERT INTO tbDonGia (MHDG,TCV,DVT,TK) "
				L"VALUES ('%s','%s','%s','Công tác vận dụng');",val,sGan[1],sGan[2]);
			ObjConn.ExecuteRB(sTvan);

			// Cập nhật trạng thái trong list
			lit_dmcv.InsertItem(iItemR+1,L"",0);
			lit_dmcv.SetItemText(iItemR+1,1,val);
			lit_dmcv.SetItemText(iItemR+1,2,sGan[1]);
			lit_dmcv.SetItemText(iItemR+1,3,sGan[2]);
			lit_dmcv.SetRowColors(iItemR+1,RGB(0,128,0),RGB(255,255,255));

			//lit_dmcv.SetItemState(iItemR,0,LVIS_SELECTED);
			//lit_dmcv.SetItemState(iItemR+1,LVIS_SELECTED,LVIS_SELECTED);
		}
		else if(iRes==7)
		{
			// Đã tồn tại --> Xóa các dữ liệu cũ (tiêu chuẩn & nội dung)
			// Xóa các dữ liệu tiêu chuẩn cũ
			sTvan.Format(L"DELETE FROM tbMaCV_Tieuchuan WHERE MaCV LIKE '%s';",val);
			ObjConn2.ExecuteRB(sTvan);

			// Xóa các dữ liệu phương pháp nghiệm thu cũ
			sTvan.Format(L"DELETE FROM tbMaCV_NDnghiemthu WHERE MaCV LIKE '%s';",val);
			ObjConn2.ExecuteRB(sTvan);
		}
		else return;

		// Thêm mới dữ liệu tiêu chuẩn
		_isl1 = lit_tccv.GetItemCount();
		for (int i = 0; i < _isl1; i++)
		{
			// Lưu dữ liệu
			if(i<9) sDem.Format(L"0%d",i+1);
			else sDem.Format(L"%d",i+1);
			sGan[1] = lit_tccv.GetItemText(i,1);
			sTvan.Format(L"INSERT INTO tbMaCV_Tieuchuan (ID,MaCV,MaTC) VALUES ('%s','%s','%s');",sDem,val,sGan[1]);
			ObjConn2.ExecuteRB(sTvan);

			// Cập nhật trạng thái trong list
			lit_tccv.SetItemText(i,0,L"");
		}

		// Thêm mới dữ liệu nội dung
		_isl2 = lit_ndkt.GetItemCount();
		for (int i = 0; i < _isl2; i++)
		{
			// Lưu dữ liệu
			if(i<9) sDem.Format(L"0%d",i+1);
			else sDem.Format(L"%d",i+1);
			sGan[1] = lit_ndkt.GetItemText(i,1);
			sGan[2] = lit_ndkt.GetItemText(i,2);
			sTvan.Format(L"INSERT INTO tbMaCV_NDnghiemthu (ID,MaCV,NoidungKiemtra,PhuongphapKiemtra) VALUES ('%s','%s','%s','%s');",sDem,val,sGan[1],sGan[2]);
			ObjConn2.ExecuteRB(sTvan);

			// Cập nhật trạng thái trong list
			lit_ndkt.SetItemText(i,0,L"");
		}		
	}
	catch(...){}
}

void Dlg_Open_DMCV::f_UpDown_listTC(int iStyle)
{
	try
	{
		int nCount=0,nItem=-1,nCol=4;
		nCount = lit_tccv.GetItemCount();
		POSITION pos=lit_tccv.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			for (int i = 0; i < nCount; i++)
				if (lit_tccv.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
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
				sGan = lit_tccv.GetItemText(nItem,j);
				lit_tccv.SetItemText(nItem,j,lit_tccv.GetItemText(nItem+iStyle,j));
				lit_tccv.SetItemText(nItem+iStyle,j,sGan);
			}

			lit_tccv.SetItemText(nItem+iStyle,0,L"1");
			lit_tccv.SetRowColors(nItem,RGB(255,255,255),RGB(0,0,0));
			lit_tccv.SetRowColors(nItem+iStyle,RGB(255,255,255),RGB(0,0,0));
			lit_tccv.SetItemState(nItem,0,LVIS_SELECTED);
			lit_tccv.SetItemState(nItem+iStyle,LVIS_SELECTED,LVIS_SELECTED);
		}
	}
	catch(...){_s(L"Lỗi hoán đổi vị trí tiêu chuẩn.");}
}


void Dlg_Open_DMCV::f_UpDown_listND(int iStyle)
{
	try
	{
		int nCount=0,nItem=-1,nCol=3;
		nCount = lit_ndkt.GetItemCount();
		POSITION pos=lit_ndkt.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			for (int i = 0; i < nCount; i++)
				if (lit_ndkt.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
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
				sGan = lit_ndkt.GetItemText(nItem,j);
				lit_ndkt.SetItemText(nItem,j,lit_ndkt.GetItemText(nItem+iStyle,j));
				lit_ndkt.SetItemText(nItem+iStyle,j,sGan);
			}

			lit_ndkt.SetItemText(nItem+iStyle,0,L"1");
			lit_ndkt.SetRowColors(nItem,RGB(255,255,255),RGB(0,0,0));
			lit_ndkt.SetRowColors(nItem+iStyle,RGB(255,255,255),RGB(0,0,0));
			lit_ndkt.SetItemState(nItem,0,LVIS_SELECTED);
			lit_ndkt.SetItemState(nItem+iStyle,LVIS_SELECTED,LVIS_SELECTED);
		}
	}
	catch(...){_s(L"Lỗi hoán đổi vị trí nội dung.");}
}


void Dlg_Open_DMCV::OnBnClickedB10()
{
	ClickEsc=0;
	if(_iTab==0) return;
	else if(_iTab==1) f_UpDown_listTC(-1);
	else f_UpDown_listND(-1);
}


void Dlg_Open_DMCV::OnBnClickedB11()
{
	ClickEsc=0;
	if(_iTab==0) return;
	else if(_iTab==1) f_UpDown_listTC(1);
	else f_UpDown_listND(1);
}


void Dlg_Open_DMCV::OnBnClickedCk1()
{
	ClickEsc=0;
}


void Dlg_Open_DMCV::OnEnKillfocusEd2()
{
	CString sEdit = L"";
	txtE_TCCV.GetWindowTextW(sEdit);

	CString sKtra = lit_tccv.GetItemText(CLRow,CLCol);
	if(sEdit != sKtra)
	{
		// Thông tin đã được chỉnh sửa
		lit_tccv.SetItemText(CLRow,0,L"1");
	}
}


void Dlg_Open_DMCV::OnEnKillfocusEd3()
{
	CString sEdit = L"";
	txtE_NDCV.GetWindowTextW(sEdit);

	CString sKtra = lit_ndkt.GetItemText(CLRow,CLCol);
	if(sEdit != sKtra)
	{
		// Thông tin đã được chỉnh sửa
		lit_ndkt.SetItemText(CLRow,0,L"1");
	}
}


void Dlg_Open_DMCV::OnLvnKeydownL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);
	if(pLVKeyDow->wVKey == VK_SPACE) OnBnClickedB2();
	else if(pLVKeyDow->wVKey == VK_DELETE) OnBnClickedB5();

	*pResult = 0;
}


void Dlg_Open_DMCV::OnLvnKeydownL2(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);
	if(pLVKeyDow->wVKey == VK_DELETE) OnBnClickedB8();

	*pResult = 0;
}


void Dlg_Open_DMCV::OnLvnKeydownL3(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);
	if(pLVKeyDow->wVKey == VK_DELETE) OnBnClickedB8();

	*pResult = 0;
}

void Dlg_Open_DMCV::OnEnSetfocusEd1()
{
	ClickEsc=1;
}


void Dlg_Open_DMCV::OnEnSetfocusT1()
{
	ClickEsc=0;
}


void Dlg_Open_DMCV::OnEnSetfocusEd2()
{
	ClickEsc=1;
}


void Dlg_Open_DMCV::OnEnSetfocusEd3()
{
	ClickEsc=1;
}


void Dlg_Open_DMCV::OnBnClickedCbs1()
{
	ClickEsc=0;
}


void Dlg_Open_DMCV::OnBnClickedCbs2()
{
	ClickEsc=0;
}


void Dlg_Open_DMCV::OnBnClickedCbs3()
{
	ClickEsc=0;
}


void Dlg_Open_DMCV::OnBnClickedCkluu()
{
	iLuuTCCV = m_chkLuuTC.GetCheck();
	m_txtTChon.EnableWindow(iLuuTCCV);
	sLuuMaCT=L"",sLuuTenCT=L"",sLuuDVT=L"",sLuuLink=L"";	

	if(iLuuTCCV==1)
	{
		int ncount = lit_dmcv.GetItemCount();
		if(ncount>0)
		{
			CString szval=L"";
			for (int i = 0; i < ncount; i++)
			{
				if((int)lit_dmcv.GetCheck(i)==1)
				{
					szval = lit_dmcv.GetItemText(i,1);
					sLuuMaCT+=szval;
					sLuuMaCT+=L"|";

					szval = lit_dmcv.GetItemText(i,2);
					szval.Replace(L"|",L" ");
					sLuuTenCT+=szval;
					sLuuTenCT+=L"|";

					szval = lit_dmcv.GetItemText(i,3);
					sLuuDVT+=szval;
					sLuuDVT+=L"|";

					sLuuLink+=sFullPath;
					sLuuLink+=L"|";
				}				
			}

			if(sLuuMaCT.Right(1)==L"|") sLuuMaCT = sLuuMaCT.Left(sLuuMaCT.GetLength()-1);
			if(sLuuTenCT.Right(1)==L"|") sLuuTenCT = sLuuTenCT.Left(sLuuTenCT.GetLength()-1);
			if(sLuuDVT.Right(1)==L"|") sLuuDVT = sLuuDVT.Left(sLuuDVT.GetLength()-1);			
			if(sLuuLink.Right(1)==L"|") sLuuLink = sLuuLink.Left(sLuuLink.GetLength()-1);

			m_txtTChon.SetWindowText(sLuuMaCT);
		}			
	}
	else m_txtTChon.SetWindowText(sLuuMaCT);
}


void Dlg_Open_DMCV::OnBnClickedCbs4()
{
	chkClrCV = m_chkColor.GetCheck();
	m_btnColor.EnableWindow(chkClrCV);
	m_txtColor.ShowWindow(chkClrCV);
}


void Dlg_Open_DMCV::OnBnClickedB13()
{
	CColorDialog dlg; 
	if (dlg.DoModal() == IDOK) 
	{ 
		fclrCV = dlg.GetColor(); 
		m_txtColor.SetBkColor(fclrCV);
		m_txtColor.SetTextColor(BLACK);
	}
}


void Dlg_Open_DMCV::OnSysCommand(UINT nID, LPARAM lParam)
{
	if(nID == SC_CLOSE) OnBnClickedB6();
	else CDialog::OnSysCommand(nID, lParam);
}


void Dlg_Open_DMCV::OnBnClickedCbs5()
{
	jNhomTC = m_chkBTC.GetCheck();
	m_cbbox.EnableWindow(jNhomTC);
	if(jNhomTC==1) GotoDlgCtrl(GetDlgItem(DSCV_CBB1));
}

void Dlg_Open_DMCV::OnNMDblclkL3(NMHDR *pNMHDR, LRESULT *pResult)
{
	ClickEsc=0;
	_iTab=2;

	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	int nItem = pNMListView->iItem;
	int nSubItem = pNMListView->iSubItem;

	if(nItem==-1 || nSubItem==-1) OnBnClickedB7();
}


void Dlg_Open_DMCV::OnNMDblclkL2(NMHDR *pNMHDR, LRESULT *pResult)
{
	ClickEsc=0;
	_iTab=1;

	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	int nItem = pNMListView->iItem;
	int nSubItem = pNMListView->iSubItem;

	if(nItem==-1 || nSubItem==-1) OnBnClickedB7();
}


void Dlg_Open_DMCV::OnContextMenu(CWnd* /*pWnd*/, CPoint point)
{
	try
	{
		if(iTCclick==-1) return;

		CMenu mnu2;
		mnu2.LoadMenu(IDR_MENU5);
		
		CListCtrlEx *pClist;
		pClist = reinterpret_cast<CListCtrlEx *>(GetDlgItem(DSCV_L2));

		CRect rectSubmitButton;
		pClist->GetWindowRect(&rectSubmitButton);

		CMenu *mnuP2 = mnu2.GetSubMenu(0);
		ASSERT(mnuP2);

		if( rectSubmitButton.PtInRect(point) )
			mnuP2->TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON, point.x, point.y, this);
	}
	catch(...){}
}


void Dlg_Open_DMCV::OnNMRClickL2(NMHDR *pNMHDR, LRESULT *pResult)
{
	ClickEsc=0;
	sDownTC=L"",sFullTC=L"";
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	iTCclick = pNMListView->iItem;
	if(iTCclick>=0)
	{
		sDownTC = lit_tccv.GetItemText(iTCclick,1);
		sFullTC = lit_tccv.GetItemText(iTCclick,2);
	}
}

void Dlg_Open_DMCV::OnNMRClickL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	ClickEsc=0;
	iTCclick=-1;
	sDownTC=L"",sFullTC=L"";
}


CString Dlg_Open_DMCV::_GetMaTieuchuan()
{
	try
	{
		iTCclick=-1;
		CString szval=L"";
		POSITION pos = lit_tccv.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			for (int i = 0; i < lit_tccv.GetItemCount(); i++)
			if (lit_tccv.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
			{
				iTCclick=i;
				szval = lit_tccv.GetItemText(i,1);
				return szval;
			}
		}		
	}
	catch(...){return L"";}
	return L"";
}


// Đọc tiêu chuẩn
void Dlg_Open_DMCV::OnTi40039()
{
	sDownTC = _GetMaTieuchuan();
	f_DocFileTC(sDownTC, 1);
}

// Tìm kiếm tiêu chuẩn
void Dlg_Open_DMCV::OnTi40040()
{
	sDownTC = _GetMaTieuchuan();
	f_SearchGoogle(L"",sDownTC,0);
}


void Dlg_Open_DMCV::OnBnClickedCk2()
{
	ClickEsc=0;
}


void Dlg_Open_DMCV::OnTi40051()
{
	try
	{
		POSITION pos = lit_tccv.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			for (int i = 0; i < lit_tccv.GetItemCount(); i++)
			if (lit_tccv.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
			{
				CLRow = i;
				qThemTC=0;
				jTTTC=1;
				_pcaltc=0;
				xl_tukhoa = L"";
				Dlg_all_tieuchuan _dlg;
				_dlg.f_Tim_kiem_TC();
				return;
			}
		}
	}
	catch(...){}
}

void Dlg_Open_DMCV::OnCbnSelchangeCbbfl()
{
	int num = m_cbbFile.GetCurSel();
	if(num==-1) return;	
	
	//if(sFiledulieu[num+1].Find(L"\\")>0) sFullPath = sFiledulieu[num+1];
	//else sFullPath.Format(L"%s\\%s",sDDanfile,sFiledulieu[num+1]);

	sFullPath = sFpathdulieu[num+1];
	this->SetWindowTextW(L"Danh sách công việc | " + sFullPath);

	OnBnClickedB1();
}

void Dlg_Open_DMCV::OnBnClickedPre()
{
	int num = m_cbbFile.GetCurSel();
	num--;

	if(num<0) num = slfileDL-1;
	m_cbbFile.SetCurSel(num);

	//if(sFiledulieu[num+1].Find(L"\\")>0) sFullPath = sFiledulieu[num+1];
	//else sFullPath.Format(L"%s\\%s",sDDanfile,sFiledulieu[num+1]);

	sFullPath = sFpathdulieu[num+1];
	this->SetWindowTextW(L"Danh sách công việc | " + sFullPath);

	OnBnClickedB1();
}

void Dlg_Open_DMCV::OnBnClickedNxt()
{
	int num = m_cbbFile.GetCurSel();
	num++;

	if(num>=slfileDL) num=0;
	m_cbbFile.SetCurSel(num);
	
	//if(sFiledulieu[num+1].Find(L"\\")>0) sFullPath = sFiledulieu[num+1];
	//else sFullPath.Format(L"%s\\%s",sDDanfile,sFiledulieu[num+1]);

	sFullPath = sFpathdulieu[num+1];
	this->SetWindowTextW(L"Danh sách công việc | " + sFullPath);

	OnBnClickedB1();
}

void Dlg_Open_DMCV::OnLvnEndScrollL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	try
	{
		if(iStopload==0) return;

		RECT r;
		CRect rectCtrl;
		lit_dmcv.GetItemRect(0, &r, LVIR_BOUNDS);
		lit_dmcv.GetWindowRect(&rectCtrl);
		int a = r.bottom - r.top;		
		int b = rectCtrl.Height();
		int topIndex = lit_dmcv.GetTopIndex();
		int ncount = lit_dmcv.GetItemCount();
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
				lit_dmcv.InsertItem(i,XLXD[i+1].CDG[0],0);
				lit_dmcv.SetItemText(i,1,XLXD[i+1].CDG[0]);
				lit_dmcv.SetItemText(i,2,XLXD[i+1].CDG[1]);
				lit_dmcv.SetItemText(i,3,XLXD[i+1].CDG[2]);
			}

			//lit_dmcv.EnsureVisible(ncount+(int)(b/a)-5, FALSE);
		}
	}
	catch(...){}
}


int Dlg_Open_DMCV::_CheckTongnghaynghi(CString szDSNgay)
{
	szDSNgay.Trim();
	if (szDSNgay == L"") return 0;

	int sl=0, tong=0;
	CString szngay[3000];
	CString szbd = L"", szkt = L"";
	_TackTukhoaChuan(szDSNgay, L"|", L";", 1, 0);

	int pos = -1;
	for (int i = 1; i <= iKey; i++)
	{
		pos = pKey[i].Find(L"-");
		if (pos > 0)
		{
			szbd = pKey[i].Left(pos);
			szkt = pKey[i].Right(pKey[i].GetLength() - pos - 1);
			tong = _NumDay2(szkt, szbd, 0);
			if (tong > 0) sl += tong;
		}
		else sl++;
	}

	return sl;
}


void Dlg_Open_DMCV::_EnterProject()
{
	try
	{
		int jColumnEnd = (int)psDMCV->Cells->SpecialCells(xlCellTypeLastCell)->GetColumn();
		CString szval = GIOR(iRow, jColumnEnd, prDMCV, L"");
		if (szval != L"") jColumnEnd++;

		RangePtr PRS = prDMCV->GetItem(iRow, jColumnEnd);
		PRS->PutNumberFormat(L"General");

		int num=0;
		CString szTong = L"", szBD = L"", szKT = L"", szLienhe = L"";
		int jTime = _iCot[30], jStart = _iCot[31], jFinish = _iCot[32];
		int jLienhe = _iCot[33], jTongnghi = _iCot[40], jChitiet = _iCot[41];
		
		PRS = prDMCV->GetItem(iRow,jTime);
		PRS->PutNumberFormat(L"General");

		szLienhe = GIOR(iRow, jLienhe, prDMCV, L"");
		szLienhe.MakeUpper();

		if(iColumn==jTime || iColumn == jTongnghi || iColumn == jChitiet)
		{
			if (iColumn == jChitiet)
			{

				szval = GIOR(iRow, jChitiet, prDMCV, L"");
				int itg = _CheckTongnghaynghi(szval);
				if (itg > 0) prDMCV->PutItem(iRow, jTongnghi, itg);
				else prDMCV->PutItem(iRow, jTongnghi, (_bstr_t)_T(""));
			}

			// Nhập giá trị vào thời gian thực hiện
			szTong = GIOR(iRow,jTime,prDMCV,L"");
			num = _wtoi(szTong);
			if(num<=0) num = 1;
			prDMCV->PutItem(iRow,jTime,num);

			szBD = GIOR(iRow,jStart,prDMCV,L"");
			if(szBD!=L"")
			{
				szKT.Format(L"=IF(%s>0,%s+%s-1,%s)+%s", GAR(iRow, jTime, prDMCV, 0),
					GAR(iRow,jStart,prDMCV,0),GAR(iRow,jTime,prDMCV,0), 
					GAR(iRow, jStart, prDMCV, 0), GAR(iRow, jTongnghi, prDMCV, 0));
				prDMCV->PutItem(iRow,jFinish,(_bstr_t)szKT);

				if (szLienhe.Find(L"FF") >= 0 || szLienhe.Find(L"SF") >= 0)
				{
					PRS = prDMCV->GetItem(iRow, jLienhe);
					PRS->ClearContents();
				}
			}
			else
			{
				szKT = GIOR(iRow,jFinish,prDMCV,L"");
				if(szKT!=L"")
				{
					szBD.Format(L"=%s-%s-%s+1",GAR(iRow,jFinish,prDMCV,0),
						GAR(iRow,jTime,prDMCV,0), GAR(iRow, jTongnghi, prDMCV, 0));
					prDMCV->PutItem(iRow, jColumnEnd, (_bstr_t)szBD);
					szBD = GIOR(iRow, jColumnEnd, prDMCV, L"");
					f_ktra_date(szBD, prDMCV, iRow, jStart);

					PRS = prDMCV->GetItem(iRow, jColumnEnd);
					PRS->ClearContents();

					szval.Format(L"%d", _wtoi(szLienhe));
					if ((_wtoi(szval)>0 && szval==szLienhe) || szLienhe.Find(L"FS") >= 0 || szLienhe.Find(L"SS") >= 0)
					{
						PRS = prDMCV->GetItem(iRow, jLienhe);
						PRS->ClearContents();
					}
				}
			}
		}
		else if(iColumn==jStart)
		{
			_Change_Start(prDMCV, iRow, jTime, jStart, jFinish, jLienhe, jTongnghi, jColumnEnd);
		}
		else if(iColumn==jFinish)
		{
			_Change_Finish(prDMCV, iRow, jTime, jStart, jFinish, jLienhe, jTongnghi, jColumnEnd);
		}
		else if(iColumn==jLienhe)
		{
			// Liên kết công việc
			szLienhe = GIOR(iRow,jLienhe,prDMCV,L"");
			szLienhe.MakeUpper();
			CString szDau = _T("+");
			
			// Kiểm tra liên hệ có phải cộng/trừ với 0 không: Ví dụ: 3SF+0 --> 3SF
			int pos = -1;
			if (szLienhe.Find(L"+") > 0 || szLienhe.Find(L"-") > 0)
			{
				pos = szLienhe.Find(L"+");
				if (pos == -1)
				{
					pos = szLienhe.Find(L"-");
					if (pos > 0) szDau = _T("-");
				}
				szval = szLienhe.Right(szLienhe.GetLength() - pos - 1);
				if (_wtoi(szval.Left(1)) <= 0) szLienhe = szLienhe.Left(pos);
			}

			// Kiểm tra nếu liên hệ có đuôi là FS --> Rút gọn cách biểu diễn: Ví dụ: 3FS -> 3;
			if (szLienhe.Right(2) == L"FS") szLienhe = szLienhe.Left(szLienhe.GetLength() - 2);

			
			// Kiểm tra có phải là số hay là công thức liên hệ?
			pos = -1;
			for (int i = 1; i < szLienhe.GetLength(); i++)
			{
				if (szLienhe.Mid(i, 1) == L"F" || szLienhe.Mid(i, 1) == L"S")
				{
					pos = i;
					break;
				}
			}

			CString szTT = GIOR(iRow, _iCot[1], prDMCV, L"");
			if (pos> 0) szval = szLienhe.Left(pos);
			else szval = szLienhe;

			if (_wtoi(szTT) == _wtoi(szval))
			{
				szLienhe = _T("");
				_s(L"Bạn không thể liên kết với chính công tác được chọn. \nVui lòng kiểm tra lại.",1);
			}
			else
			{
				if (pos > 0)
				{
					num = 0;
					int result = szLienhe.Find(L"+");
					if (result == -1)
					{
						result = szLienhe.Find(L"-");
						if (result > 0) szDau = _T("-");
					}
					if (result > 0) num = _wtoi(szLienhe.Right(szLienhe.GetLength()-result-1));

					int iRowLK = FindValue(psDMCV, _iCot[1], 8, 0, (_bstr_t)szLienhe.Left(pos));
					if (iRowLK > 0)
					{
						result = 0;
						if (szLienhe.Find(L"FS") > 0)
						{
							// Đổ vào cột bắt đầu
							if(num>0) szval.Format(L"=%s%s%d+1", GAR(iRowLK, jFinish, prDMCV, 3), szDau, num);
							else szval.Format(L"=%s+1", GAR(iRowLK, jFinish, prDMCV, 3));
							prDMCV->PutItem(iRow, jStart, (_bstr_t)szval);
							result = 1;
						}
						else if (szLienhe.Find(L"FF") > 0)
						{
							// Đổ vào cột kết thúc
							if (num > 0) szval.Format(L"=%s%s%d", GAR(iRowLK, jFinish, prDMCV, 3), szDau, num);
							else szval.Format(L"=%s", GAR(iRowLK, jFinish, prDMCV, 3));
							prDMCV->PutItem(iRow, jFinish, (_bstr_t)szval);
							result = 2;
						}
						else if (szLienhe.Find(L"SS") > 0)
						{
							// Đổ vào cột bắt đầu
							if (num > 0) szval.Format(L"=%s%s%d", GAR(iRowLK, jStart, prDMCV, 3), szDau, num);
							else szval.Format(L"=%s", GAR(iRowLK, jStart, prDMCV, 3));
							prDMCV->PutItem(iRow, jStart, (_bstr_t)szval);
							result = 1;
						}
						else if (szLienhe.Find(L"SF") > 0)
						{
							// Đổ vào cột kết thúc
							if (num > 0) szval.Format(L"=%s%s%d-1", GAR(iRowLK, jStart, prDMCV, 3), szDau, num);
							else szval.Format(L"=%s-1", GAR(iRowLK, jStart, prDMCV, 3));
							prDMCV->PutItem(iRow, jFinish, (_bstr_t)szval);
							result = 2;
						}
						else
						{
							szLienhe = _T("");
							
							_s(L"Bạn cần nhập giá trị liên hệ đúng theo quy ước trong Project. "
								L"\nVí dụ: 3FS+2, 2FF-3, 6SF+0, 9SS+1, ...",1);
						}

						if (result == 1)
						{
							// Thay đổi giá trị ở cột bắt đầu
							_Change_Start(prDMCV, iRow, jTime, jStart, jFinish, jLienhe, jTongnghi, jColumnEnd);
							
						}
						else if (result == 2)
						{
							// Thay đổi giá trị ở cột kết thúc
							_Change_Finish(prDMCV, iRow, jTime, jStart, jFinish, jLienhe, jTongnghi, jColumnEnd);
						}
					}
					else
					{
						szLienhe = _T("");
						_s(L"Không tồn tại công tác cần liên kết. \nVui lòng kiểm tra lại.",1);
					}					
				}
				else
				{
					// Đây là số
					int iRowLket = FindValue(psDMCV, _iCot[1], 8, 0, (_bstr_t)szLienhe);
					if (iRowLket > 0)
					{
						szBD.Format(L"=%s+1", GAR(iRowLket, jFinish, prDMCV, 0));
						prDMCV->PutItem(iRow, jStart, (_bstr_t)szBD);

						// Thay đổi giá trị ở cột bắt đầu
						_Change_Start(prDMCV, iRow, jTime, jStart, jFinish, jLienhe, jTongnghi, jColumnEnd);
					}
					else
					{
						szLienhe = _T("");
						_s(L"Không tồn tại công tác cần liên kết. \nVui lòng kiểm tra lại.",1);
					}
				}
			}			
			
			prDMCV->PutItem(iRow, jLienhe, (_bstr_t)szLienhe);
		}

		_Check_nghile(prDMCV, iRow, jStart);
		_Check_nghile(prDMCV, iRow, jFinish);
	}
	catch(...){}
}


void Dlg_Open_DMCV::_Change_Start(RangePtr prDMCV, int iRow, int jTime, int jStart, int jFinish, int jLienhe, int jTongnghi, int jColumnEnd)
{
	try
	{
		int num = 0;
		CString szBD = L"", szKT = L"", szLienhe = L"";
		RangePtr PRS;

		CString szTong = GIOR(iRow, jTime, prDMCV, L"");
		if (szTong != L"")
		{
			num = _wtoi(szTong);
			if (num <= 0) num = 1;
			prDMCV->PutItem(iRow, jTime, num);

			szKT.Format(L"=IF(%s>0,%s+%s-1,%s)+%s", GAR(iRow, jTime, prDMCV, 0),
				GAR(iRow, jStart, prDMCV, 0), GAR(iRow, jTime, prDMCV, 0), 
				GAR(iRow, jStart, prDMCV, 0), GAR(iRow, jTongnghi, prDMCV, 0));
			prDMCV->PutItem(iRow, jFinish, (_bstr_t)szKT);

			if (szLienhe.Find(L"FF") >= 0 || szLienhe.Find(L"SF") >= 0)
			{
				PRS = prDMCV->GetItem(iRow, jLienhe);
				PRS->ClearContents();
			}
		}
		else
		{
			szBD = GIOR(iRow, jStart, prDMCV, L"");
			szKT = GIOR(iRow, jFinish, prDMCV, L"");
			if (szKT != L"")
			{
				// return =0 --> szBD > szKT
				int ilen = szKT.GetLength();
				if (compare_date(ilen, szBD, szKT) == 0)
				{
					prDMCV->PutItem(iRow, jTime, 1);

					szKT.Format(L"=IF(%s>0,%s+%s-1,%s)+%s", GAR(iRow, jTime, prDMCV, 0),
						GAR(iRow, jStart, prDMCV, 0), GAR(iRow, jTime, prDMCV, 0), 
						GAR(iRow, jStart, prDMCV, 0), GAR(iRow, jTongnghi, prDMCV, 0));
					prDMCV->PutItem(iRow, jFinish, (_bstr_t)szKT);

					if (szLienhe.Find(L"FF") >= 0 || szLienhe.Find(L"SF") >= 0)
					{
						PRS = prDMCV->GetItem(iRow, jLienhe);
						PRS->ClearContents();
					}
				}
				else
				{
					szTong.Format(L"=%s-%s-%s+1", GAR(iRow, jFinish, prDMCV, 0), 
						GAR(iRow, jStart, prDMCV, 0), GAR(iRow, jTongnghi, prDMCV, 0));
					prDMCV->PutItem(iRow, jColumnEnd, (_bstr_t)szTong);
					szTong = GIOR(iRow, jColumnEnd, prDMCV, L"");
					prDMCV->PutItem(iRow, jTime, (_bstr_t)szTong);

					PRS = prDMCV->GetItem(iRow, jColumnEnd);
					PRS->ClearContents();
				}
			}
		}

		CString szval = L"";
		szval.Format(L"%d", _wtoi(szLienhe));
		if ((_wtoi(szval) > 0 && szval == szLienhe) || szLienhe.Find(L"FS") >= 0 || szLienhe.Find(L"SS") >= 0)
		{
			PRS = prDMCV->GetItem(iRow, jLienhe);
			PRS->ClearContents();
		}
	}
	catch(...){}
}


void Dlg_Open_DMCV::_Change_Finish(RangePtr prDMCV, int iRow, int jTime, int jStart, int jFinish, int jLienhe, int jTongnghi, int jColumnEnd)
{
	try
	{
		int num = 0;
		CString szBD = L"", szKT = L"", szLienhe = L"", szTong = L"";
		RangePtr PRS = prDMCV->GetItem(iRow, jFinish);
		szKT = PRS->GetFormula();
		szKT.Replace(L"$", L"");
		CString szval = GAR(iRow, jStart, prDMCV, 0);
		if (szKT.Find(szval) >= 0)
		{
			szBD = GIOR(iRow, jStart, prDMCV, L"");
			szKT = GIOR(iRow, jFinish, prDMCV, L"");
			if (szBD != L"")
			{
				// return =0 --> szBD > szKT
				int ilen = szKT.GetLength();
				if (compare_date(ilen, szBD, szKT) == 0)
				{
					prDMCV->PutItem(iRow, jTime, 1);
					prDMCV->PutItem(iRow, jStart, (_bstr_t)szKT);

					szval.Format(L"%d", _wtoi(szLienhe));
					if ((_wtoi(szval) > 0 && szval == szLienhe) || szLienhe.Find(L"FS") >= 0 || szLienhe.Find(L"SS") >= 0)
					{
						PRS = prDMCV->GetItem(iRow, jLienhe);
						PRS->ClearContents();
					}
				}
				else
				{
					szTong.Format(L"=%s-%s-%s+1", GAR(iRow, jFinish, prDMCV, 0),
						GAR(iRow, jStart, prDMCV, 0), GAR(iRow, jTongnghi, prDMCV, 0));
					prDMCV->PutItem(iRow, jColumnEnd, (_bstr_t)szTong);
					szTong = GIOR(iRow, jColumnEnd, prDMCV, L"");
					prDMCV->PutItem(iRow, jTime, (_bstr_t)szTong);

					PRS = prDMCV->GetItem(iRow, jColumnEnd);
					PRS->ClearContents();
				}
			}
			else
			{
				num = _wtoi(szTong);
				if (num <= 0) num = 1;
				prDMCV->PutItem(iRow, jTime, num);

				szBD.Format(L"=%s-%s-%s+1", GAR(iRow, jFinish, prDMCV, 0), 
					GAR(iRow, jTime, prDMCV, 0), GAR(iRow, jTongnghi, prDMCV, 0));
				prDMCV->PutItem(iRow, jColumnEnd, (_bstr_t)szBD);
				szBD = GIOR(iRow, jColumnEnd, prDMCV, L"");
				f_ktra_date(szBD, prDMCV, iRow, jStart);

				PRS = prDMCV->GetItem(iRow, jColumnEnd);
				PRS->ClearContents();

				szval.Format(L"%d", _wtoi(szLienhe));
				if ((_wtoi(szval) > 0 && szval == szLienhe) || szLienhe.Find(L"FS") >= 0 || szLienhe.Find(L"SS") >= 0)
				{
					PRS = prDMCV->GetItem(iRow, jLienhe);
					PRS->ClearContents();
				}
			}
		}
		else
		{
			szTong = GIOR(iRow, jTime, prDMCV, L"");
			if (szTong != L"")
			{
				num = _wtoi(szTong);
				if (num <= 0) num = 1;
				prDMCV->PutItem(iRow, jTime, num);

				szBD.Format(L"=%s-%s-%s+1", GAR(iRow, jFinish, prDMCV, 0), 
					GAR(iRow, jTime, prDMCV, 0), GAR(iRow, jTongnghi, prDMCV, 0));
				prDMCV->PutItem(iRow, jColumnEnd, (_bstr_t)szBD);
				szBD = GIOR(iRow, jColumnEnd, prDMCV, L"");
				f_ktra_date(szBD, prDMCV, iRow, jStart);

				PRS = prDMCV->GetItem(iRow, jColumnEnd);
				PRS->ClearContents();

				szval.Format(L"%d", _wtoi(szLienhe));
				if ((_wtoi(szval) > 0 && szval == szLienhe) || szLienhe.Find(L"FS") >= 0 || szLienhe.Find(L"SS") >= 0)
				{
					PRS = prDMCV->GetItem(iRow, jLienhe);
					PRS->ClearContents();
				}
			}
			else
			{
				szBD = GIOR(iRow, jStart, prDMCV, L"");
				szKT = GIOR(iRow, jFinish, prDMCV, L"");
				if (szBD != L"")
				{
					// return =0 --> szBD > szKT
					int ilen = szBD.GetLength();
					if (compare_date(ilen, szBD, szKT) == 0)
					{
						prDMCV->PutItem(iRow, jTime, 1);
						prDMCV->PutItem(iRow, jStart, (_bstr_t)szKT);

						szval.Format(L"%d", _wtoi(szLienhe));
						if ((_wtoi(szval) > 0 && szval == szLienhe) || szLienhe.Find(L"FS") >= 0 || szLienhe.Find(L"SS") >= 0)
						{
							PRS = prDMCV->GetItem(iRow, jLienhe);
							PRS->ClearContents();
						}
					}
					else
					{
						szTong.Format(L"=%s-%s-%s+1", GAR(iRow, jFinish, prDMCV, 0), 
							GAR(iRow, jStart, prDMCV, 0), GAR(iRow, jTongnghi, prDMCV, 0));
						prDMCV->PutItem(iRow, jColumnEnd, (_bstr_t)szTong);
						szTong = GIOR(iRow, jColumnEnd, prDMCV, L"");
						prDMCV->PutItem(iRow, jTime, (_bstr_t)szTong);

						PRS = prDMCV->GetItem(iRow, jColumnEnd);
						PRS->ClearContents();
					}
				}
			}
		}

		if (szLienhe.Find(L"FF") >= 0 || szLienhe.Find(L"SF") >= 0)
		{
			PRS = prDMCV->GetItem(iRow, jLienhe);
			PRS->ClearContents();
		}
	}
	catch(...){}
}


void Dlg_Open_DMCV::_Check_nghile(RangePtr pr1, int ir, int ic)
{
	try
	{
		long gtri = 0;
		if (ichkLe == 1)
		{
			// Đọc từ file MDB
			gtri = f_Read_Nhatky();

			// Đọc file dữ liệu csv
			CString spath = L"";
			spath.Format(L"%s\\%s", _fGPathFolder(L"ToolQLCL.dll"), _dFileNle);
			spath.Replace(L"\\\\", L"\\");
			gtri += f_Read_file_CSV(spath, gtri);
		}

		RangePtr PRS = pr1->GetItem(ir, ic);
		PRS->ClearComments();

		int _iktra = 0;
		CString sngay = GIOR(ir, ic, pr1, L"");
		if (sngay != L"")
		{
			if (ichk7 == 1 || ichkCN == 1)
			{
				/*// Nghỉ cuối tuần
				if (ichk7 == 0 && ichkCN == 1)
					szval.Format(L"=IFERROR(CHOOSE(WEEKDAY(%s),\"CN\"),\"\")", GAR(ir, ic, pr1, 0));
				else if (ichk7 == 1 && ichkCN == 0)
					szval.Format(
						L"=CHOOSE(WEEKDAY(%s),\"\",\"\",\"\",\"\",\"\",\"\",\"T7\")",
						GAR(ir, ic, pr1, 0));
				else if (ichk7 == 1 && ichkCN == 1)
					szval.Format(
						L"=CHOOSE(WEEKDAY(%s),\"CN\",\"\",\"\",\"\",\"\",\"\",\"T7\")",
						GAR(ir, ic, pr1, 0));

				pr1->PutItem(1, 99, (_bstr_t)szval);
				szval = GIOR(1, 99, pr1, L"");*/

				sngay = fThuMay(sngay);
				if ((ichk7 == 1 && sngay == L"T7") || (ichkCN == 1 && sngay == L"CN"))
				{
					_iktra = 1;
					PRS->Interior->PutColor(clrKT_Bg);
					PRS->Font->PutColor(clrKT_Txt);
				}
			}

			if (_iktra == 0)
			{
				// Kiểm tra tiếp xem có phải ngày nghỉ lễ tết không?
				if (gtri > 0)
				{
					for (int k = 1; k <= gtri; k++)
					{
						if (sngay == sNgayle[k])
						{
							_iktra = 1;
							PRS->Interior->PutColor(clrKT_Bg);
							PRS->Font->PutColor(clrKT_Txt);
							if (ichkCmt == 1 && sGchu[k] != L"") PRS->AddComment((_bstr_t)sGchu[k]);
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

		//pr1->PutItem(1, 99, (_bstr_t)L"");
	}
	catch (...) { _s(L"Lỗi kiểm tra ngày nghỉ lễ."); }
}

long Dlg_Open_DMCV::f_Read_Nhatky()
{
	try
	{
		szCheckNgay = L"";
		if (getPathNHKY() == L"") return 0;
		if (connectDsn(pathNK) == -1) return 0;

		CConnection ObjConn;
		ObjConn.OpenAccessDB(pathNK, L"", L"");
		CRecordset* Recset = ObjConn.Execute(L"SELECT * FROM tbBoqua;");

		int dem = 0;
		int icheck = 0;
		CString szval = L"";
		while (!Recset->IsEOF())
		{
			icheck = 0;
			Recset->GetFieldValue(L"ngayghi", szval);

			if (szCheckNgay != L"")
			{
				if (szCheckNgay.Find(szval) >= 0) icheck = 1;
			}

			if (icheck == 0)
			{
				dem++;
				sNgayle[dem] = szval;
				Recset->GetFieldValue(L"ghichu", sGchu[dem]);

				szCheckNgay += sNgayle[dem];
				szCheckNgay += L"|";
			}

			Recset->MoveNext();
		}
		Recset->Close();
		ObjConn.CloseDatabase();

		return dem;
	}
	catch (...) {}
	return 0;
}


long Dlg_Open_DMCV::f_Read_file_CSV(CString pathCSV, long num)
{
	int icheck = 0;
	long cong = num;

	try
	{
		wchar_t **dArray;
		int numberOfLine = 0;
		int i;

		long _size;
		_size = f_GetLineFile(pathCSV);
		_size++;
		dArray = (wchar_t **)calloc(_size, sizeof(wchar_t *));
		wchar_t **p;
		p = (wchar_t **)calloc(_size, sizeof(wchar_t *));
		for (i = 0; i < _size; i++)
		{
			*(dArray + i) = (wchar_t *)malloc(SIZE_LINE * sizeof(wchar_t));
			*(p + i) = (wchar_t *)malloc(SIZE_LINE * sizeof(wchar_t));
		}

		FILE *pReadFile;
		if (_wfopen_s(&pReadFile, pathCSV, L"r, ccs=UTF-16LE") != 0)
		{
			_s(L"Không thể mở được file dữ liệu. Vui lòng chọn lại file dữ liệu.");
			for (i = 0; i < _size; i++)
			{
				free(*(dArray + i));
			}
			free(dArray);
			return 0;
		}
		else
		{
			i = 0;
			CString szVarString = L"";
			while (fgetws(*(p + i), ROW_LINE, pReadFile)) i++;

			numberOfLine = i;
			fclose(pReadFile);

			int fdem = 0;
			CString ft[2] = { L"",L"" };

			for (int i = 1; i < numberOfLine; i++)
			{
				wcscpy_s(*(dArray + i), wcslen(*(p + i)) + 1, *(p + i));
				wchar_t *textString = new wchar_t[SIZE_LINE];
				wcscpy_s(textString, wcslen(*(dArray + i)) + 1, *(dArray + i));
				szVarString = textString;

				long length;
				CString szLeft;
				int nToken = szVarString.Find(L'\t');

				fdem = 0;
				while (nToken >= 0)
				{
					length = szVarString.GetLength();
					szLeft = szVarString.Left(nToken);
					szLeft.TrimLeft(); szLeft.TrimRight();
					szVarString = szVarString.Right(length - nToken - 1);
					nToken = szVarString.Find(L'\t');

					fdem++;
					if (fdem == 1) ft[0] = szLeft;

					// Thêm vào để lấy cột cuối cùng (by Leo 23.06.15)
					ft[1] = L"";
					if (nToken < 0) ft[1] = szVarString;
				}

				ft[0].Trim();
				ft[0].Replace(L" ", L"");
				ft[0].Replace(L"N", L"");
				if (ft[0] != L"")
				{
					icheck = 0;
					if (szCheckNgay != L"")
					{
						if (szCheckNgay.Find(ft[0]) >= 0) icheck = 1;
					}

					if (icheck == 0)
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
			free(*(dArray + i));
			free(*(p + i));
		}
		free(dArray);
		free(p);
	}
	catch (_com_error & error) { _s(L"#QL4.126: Lỗi đọc dữ liệu ngày tháng nghỉ lễ."); }

	return cong;
}


long Dlg_Open_DMCV::f_GetLineFile(CString szPath)
{
	FILE * pFile;
	wchar_t *c = (wchar_t *)malloc(2000 * sizeof(wchar_t));
	long line = 0;
	pFile = _wfopen(szPath, L"r, ccs=UTF-16LE");
	if (pFile == NULL) perror("Error opening file");
	else
	{
		while (fgetws(c, 2000, pFile))
		{
			line++;
		}
		fclose(pFile);
		printf("File contains %d.\n", line);
	}
	free(c);
	return line;
}

int Dlg_Open_DMCV::_CheckMultiItem()
{
	if (m_chkLuuTC.GetCheck() == 1)
	{
		CString szval = _T("");
		m_txtTChon.GetWindowTextW(szval);
		int pos = szval.Find(L"|");

		if (pos > 0 && pos < szval.GetLength() - 1) return 1;
	}

	int num = 0, icheck = -1;
	int ncount = lit_dmcv.GetItemCount();
	if (ncount > 0)
	{
		for (int i = 0; i < ncount; i++)
		{
			if ((int)lit_dmcv.GetCheck(i) == 1)
			{
				icheck = i;
				break;
			}
		}

		if (icheck >= 0)
		{
			for (int i = icheck + 1; i < ncount; i++)
			{
				if (lit_dmcv.GetCheck(i) == true)
				{
					num++;
					if (num >= 2) return 1;
				}				
			}
		}
		else
		{
			for (int i = 0; i < ncount; i++)
			{
				if (lit_dmcv.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					num++;
					if (num >= 2) return 1;
				}
			}
		}
	}

	return 0;
}

void Dlg_Open_DMCV::OnBnClickedCkdmcvcon()
{
	iLuuDMCVCon = m_chkDMCVCon.GetCheck();
}
