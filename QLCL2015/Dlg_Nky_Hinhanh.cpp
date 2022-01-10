#include "stdafx.h"
#include "Dlg_Nky_Hinhanh.h"

// Dlg_Nky_Hinhanh dialog
IMPLEMENT_DYNAMIC(Dlg_Nky_Hinhanh, CDialogEx)

Dlg_Nky_Hinhanh::Dlg_Nky_Hinhanh(CWnd* pParent /*=nullptr*/)
	: CDialogEx(DLG_NKY_HINHANH, pParent)
{

}

Dlg_Nky_Hinhanh::~Dlg_Nky_Hinhanh()
{
}

void Dlg_Nky_Hinhanh::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(Dlg_Nky_Hinhanh, CDialogEx)
	ON_BN_CLICKED(IDC_BUTTON4, &Dlg_Nky_Hinhanh::OnBnClickedButton4)
END_MESSAGE_MAP()

// Dlg_Nky_Hinhanh message handlers
void Dlg_Nky_Hinhanh::OnBnClickedButton4()
{
	CDialogEx::OnCancel();
}
