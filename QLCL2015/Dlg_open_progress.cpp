#include "Dlg_open_progress.h"

// Dlg_open_progress dialog
IMPLEMENT_DYNAMIC(Dlg_open_progress, CDialogEx)

Dlg_open_progress::Dlg_open_progress(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_DIALOG_PROGRESS, pParent)
{
	
}

Dlg_open_progress::~Dlg_open_progress()
{
}

void Dlg_open_progress::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(Dlg_open_progress, CDialogEx)
	ON_WM_TIMER()
END_MESSAGE_MAP()


// Dlg_open_progress message handlers
BOOL Dlg_open_progress::OnInitDialog()
{
	CDialog::OnInitDialog();
	
	demTimer = 0;
	SetTimer(1, 10, NULL);

	return TRUE;
}

void Dlg_open_progress::OnTimer(UINT_PTR nIDEvent)
{
	demTimer++;
	if (demTimer >= 2) CDialog::OnCancel();
	CDialogEx::OnTimer(nIDEvent);
}
