#include "stdafx.h"
#include "Dlg_NhomkyBB.h"

// Dlg_NhomkyBB dialog
IMPLEMENT_DYNAMIC(Dlg_NhomkyBB, CDialogEx)

Dlg_NhomkyBB::Dlg_NhomkyBB(CWnd* pParent /*=NULL*/)
	: CDialogEx(Dlg_NhomkyBB::IDD, pParent)
{

}

Dlg_NhomkyBB::~Dlg_NhomkyBB()
{
}

void Dlg_NhomkyBB::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, NKBB_L1, m_list);
}


BEGIN_MESSAGE_MAP(Dlg_NhomkyBB, CDialogEx)
	ON_BN_CLICKED(NKBB_B1, &Dlg_NhomkyBB::OnBnClickedB1)
	ON_BN_CLICKED(NKBB_B2, &Dlg_NhomkyBB::OnBnClickedB2)
	ON_NOTIFY(NM_DBLCLK, NKBB_L1, &Dlg_NhomkyBB::OnNMDblclkL1)
	ON_NOTIFY(LVN_KEYDOWN, NKBB_L1, &Dlg_NhomkyBB::OnLvnKeydownL1)
END_MESSAGE_MAP()


// Dlg_NhomkyBB message handlers
BOOL Dlg_NhomkyBB::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	_LoadDS();
	return TRUE; 
}

// Bắt sự kiện click Enter
BOOL Dlg_NhomkyBB::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if (pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(NKBB_L1))
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


void Dlg_NhomkyBB::_LoadDS()
{
	m_list.InsertColumn(0, L"", LVCFMT_CENTER, 0);
	m_list.InsertColumn(1, L"STT", LVCFMT_CENTER, 40);
	m_list.InsertColumn(2, L"Nhóm ký", LVCFMT_LEFT, 200);
	m_list.InsertColumn(3, L"Ngày bắt đầu", LVCFMT_CENTER, 120);
	m_list.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);

	m_list.InsertItem(0, L"", 0);
	m_list.SetItemText(0, 1, L"");
	m_list.SetItemText(0, 2, L"NHÓM MẶC ĐỊNH");

	_bstr_t shKyBB = name_sheet((_bstr_t)"Sheet1");
	_WorksheetPtr psKyBB = xl->Sheets->GetItem(shKyBB);
	RangePtr prKyBB = psKyBB->Cells;

	int dem = 0;
	int slnhom = 0;
	CString szval = L"", slg = L"";
	for (int i = 4; i <= 100; i++)
	{
		szval = GIOR(3, i, prKyBB, L"");
		szval.Trim();
		if (szval != L"")
		{
			slnhom++;
			slg.Format(L"%d", slnhom);
			m_list.InsertItem(slnhom, L"", 0);
			m_list.SetItemText(slnhom, 1, slg);
			m_list.SetItemText(slnhom, 2, szval);

			szval = GIOR(1, i, prKyBB, L"");
			m_list.SetItemText(slnhom, 3, szval);

			dem = 0;
		}
		else dem++;
		if (dem == 10) break;
	}
}


void Dlg_NhomkyBB::OnBnClickedB1()
{
	POSITION pos=m_list.GetFirstSelectedItemPosition();
	if(pos!=NULL)
	{
		for (int i = 0; i < m_list.GetItemCount(); i++)
		{
			if (m_list.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
			{
				sTukhoa.Format(L"%d", i);
				CDialog::OnOK();
				return;
			}
		}		
	}
}


void Dlg_NhomkyBB::OnBnClickedB2()
{
	CDialog::OnCancel();
}


void Dlg_NhomkyBB::OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	OnBnClickedB1();
	*pResult = 0;
}


void Dlg_NhomkyBB::OnLvnKeydownL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);
	if(pLVKeyDow->wVKey == VK_SPACE) OnBnClickedB1();

	*pResult = 0;
}
