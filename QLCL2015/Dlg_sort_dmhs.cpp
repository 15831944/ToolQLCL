#include "stdafx.h"
#include "Dlg_sort_dmhs.h"

Dlg_sort_dmhs::Dlg_sort_dmhs() : CDialog(Dlg_sort_dmhs::IDD)
{
	iCheckTenTT = 0;
}

void Dlg_sort_dmhs::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, SRT_CK1, Crd1);
	DDX_Control(pDX, SRT_CK2, Crd2);
	DDX_Control(pDX, SRT_CB1, Cbb1);
	DDX_Control(pDX, SRT_CHKB1, Ckhb1);	
}

BEGIN_MESSAGE_MAP(Dlg_sort_dmhs, CDialog)
	ON_BN_CLICKED(IDC_BUTTON2, &Dlg_sort_dmhs::OnBnClickedButton2)
	ON_BN_CLICKED(SRT_B1, &Dlg_sort_dmhs::OnBnClickedB1)
	ON_BN_CLICKED(SRT_CK1, &Dlg_sort_dmhs::OnBnClickedCk1)
	ON_BN_CLICKED(SRT_CK2, &Dlg_sort_dmhs::OnBnClickedCk2)
	ON_CBN_SELCHANGE(SRT_CB1, &Dlg_sort_dmhs::OnCbnSelchangeCb1)
END_MESSAGE_MAP()

BOOL Dlg_sort_dmhs::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	Cbb1.AddString(L"Thứ tự mặc định");
	Cbb1.AddString(L"Mã hiệu mặc định");
	Cbb1.AddString(L"Ngày lấy mẫu");
	Cbb1.AddString(L"Ngày nghiệm thu VL/CV");
	Cbb1.AddString(L"Đợt thanh toán");
	Cbb1.SetCurSel(iscbb);

	// istang=1 tăng dần
	if(istang==1) Crd1.SetCheck(1);
	else Crd2.SetCheck(1);
	Ckhb1.SetCheck(ichkpnhom);
	//Ckhb1.EnableWindow(istang);

	return TRUE; 
}


// Bắt sự kiện click Enter
BOOL Dlg_sort_dmhs::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if (pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(SRT_CB1))
	{
		OnBnClickedB1();
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		CDialog::OnCancel();
		return TRUE; 
	}

	return FALSE;
}


void Dlg_sort_dmhs::OnBnClickedButton2()
{
	CDialog::OnCancel();
}


void Dlg_sort_dmhs::OnBnClickedB1()
{
	isktrHS=1;
	ichkpnhom = Ckhb1.GetCheck();
	istang = Crd1.GetCheck();
	iscbb = Cbb1.GetCurSel();

	CButton* m_check = (CButton*)GetDlgItem(SRT_CHKB2);
	iCheckTenTT = (int)m_check->GetCheck();

	CDialog::OnOK();
}


void Dlg_sort_dmhs::OnBnClickedCk1()
{
	//Ckhb1.EnableWindow(1);
}


void Dlg_sort_dmhs::OnBnClickedCk2()
{
	//Ckhb1.EnableWindow(0);
}


void Dlg_sort_dmhs::OnCbnSelchangeCb1()
{
	int ncbb = Cbb1.GetCurSel();
	if(ncbb==4) Ckhb1.SetWindowText(L"Phân chia chi tiết theo đợt");
	else Ckhb1.SetWindowText(L"Phân chia chi tiết hạng mục, giai đoạn");
}
