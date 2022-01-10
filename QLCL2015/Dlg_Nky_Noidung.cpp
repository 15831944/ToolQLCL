#include "stdafx.h"
#include "Dlg_Nky_Noidung.h"

// Dlg_Nky_Noidung dialog
IMPLEMENT_DYNAMIC(Dlg_Nky_Noidung, CDialogEx)

Dlg_Nky_Noidung::Dlg_Nky_Noidung(CWnd* pParent /*=nullptr*/)
	: CDialogEx(DLG_NKY_NOIDUNG, pParent)
{

}

Dlg_Nky_Noidung::~Dlg_Nky_Noidung()
{
}

void Dlg_Nky_Noidung::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(Dlg_Nky_Noidung, CDialogEx)
	ON_BN_CLICKED(IDC_BUTTON3, &Dlg_Nky_Noidung::OnBnClickedButton3)
END_MESSAGE_MAP()

// Dlg_Nky_Noidung message handlers
void Dlg_Nky_Noidung::OnBnClickedButton3()
{
	CDialogEx::OnCancel();
}
