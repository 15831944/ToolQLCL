#include "stdafx.h"
#include "Dlg_Nky_BangNLTB.h"

// Dlg_Nky_BangNLTB dialog
IMPLEMENT_DYNAMIC(Dlg_Nky_BangNLTB, CDialogEx)

Dlg_Nky_BangNLTB::Dlg_Nky_BangNLTB(CWnd* pParent /*=nullptr*/)
	: CDialogEx(DLG_NKY_BANGNLTB, pParent)
{

}

Dlg_Nky_BangNLTB::~Dlg_Nky_BangNLTB()
{
}

void Dlg_Nky_BangNLTB::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(Dlg_Nky_BangNLTB, CDialogEx)
	ON_BN_CLICKED(IDC_BUTTON3, &Dlg_Nky_BangNLTB::OnBnClickedButton3)
END_MESSAGE_MAP()

// Dlg_Nky_BangNLTB message handlers
void Dlg_Nky_BangNLTB::OnBnClickedButton3()
{
	CDialogEx::OnCancel();
}
