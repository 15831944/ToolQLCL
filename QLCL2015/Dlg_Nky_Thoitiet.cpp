#include "stdafx.h"
#include "Dlg_Nky_Thoitiet.h"

// Dlg_Nky_Thoitiet dialog
IMPLEMENT_DYNAMIC(Dlg_Nky_Thoitiet, CDialogEx)

Dlg_Nky_Thoitiet::Dlg_Nky_Thoitiet(CWnd* pParent /*=nullptr*/)
	: CDialogEx(DLG_NKY_THOITIET, pParent)
{

}

Dlg_Nky_Thoitiet::~Dlg_Nky_Thoitiet()
{
}

void Dlg_Nky_Thoitiet::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(Dlg_Nky_Thoitiet, CDialogEx)
	ON_BN_CLICKED(IDC_BUTTON3, &Dlg_Nky_Thoitiet::OnBnClickedButton3)
END_MESSAGE_MAP()

// Dlg_Nky_Thoitiet message handlers
void Dlg_Nky_Thoitiet::OnBnClickedButton3()
{
	CDialogEx::OnCancel();
}
