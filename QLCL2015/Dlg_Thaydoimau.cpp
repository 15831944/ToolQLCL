#include "stdafx.h"
#include "Dlg_Thaydoimau.h"


IMPLEMENT_DYNAMIC(Dlg_Thaydoimau, CDialogEx)

Dlg_Thaydoimau::Dlg_Thaydoimau(CWnd* pParent /*=NULL*/)
	: CDialogEx(Dlg_Thaydoimau::IDD, pParent)
{

}

Dlg_Thaydoimau::~Dlg_Thaydoimau()
{
}

void Dlg_Thaydoimau::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(Dlg_Thaydoimau, CDialogEx)
	ON_BN_CLICKED(CHGMAU_B1, &Dlg_Thaydoimau::OnBnClickedB1)
	ON_BN_CLICKED(CHGMAU_B2, &Dlg_Thaydoimau::OnBnClickedB2)
END_MESSAGE_MAP()


// Dlg_Thaydoimau message handlers
BOOL Dlg_Thaydoimau::OnInitDialog()
{
	CDialog::OnInitDialog();

	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);
	
	return TRUE;
}


BOOL Dlg_Thaydoimau::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	
	if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		CDialog::OnCancel();
		return TRUE; 
	}

	return FALSE;
}


void Dlg_Thaydoimau::OnBnClickedB1()
{
	CDialog::OnOK();	
}


void Dlg_Thaydoimau::OnBnClickedB2()
{
	CDialog::OnCancel();
}
