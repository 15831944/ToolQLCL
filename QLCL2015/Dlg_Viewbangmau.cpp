#include "stdafx.h"
#include "Dlg_Viewbangmau.h"
#include "HookMouse.h"

// Dlg_Viewbangmau dialog
IMPLEMENT_DYNAMIC(Dlg_Viewbangmau, CDialogEx)

Dlg_Viewbangmau::Dlg_Viewbangmau(CWnd* pParent /*=nullptr*/)
	: CDialogEx(DLG_VIEWBMAU, pParent)
{
	szNumRow = L"";
	szTenbangmau = L"";
	MyHook::Instance().InstallHook();
}

Dlg_Viewbangmau::~Dlg_Viewbangmau()
{
	MyHook::Instance().UninstallHook();
}

void Dlg_Viewbangmau::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST_BMAU, m_list);	
}


BEGIN_MESSAGE_MAP(Dlg_Viewbangmau, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_BN_CLICKED(IDC_BUTTON1, &Dlg_Viewbangmau::OnBnClickedButton1)
END_MESSAGE_MAP()


// Dlg_Viewbangmau message handlers
BOOL Dlg_Viewbangmau::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	CStatic* m_stt = (CStatic*)GetDlgItem(IDC_STATIC_TENMAU);
	m_stt->SetWindowText(szTenbangmau);

	LoadList();

	SetLocation();
	hWndViewBmau = m_hWnd;

	return TRUE;
}

BOOL Dlg_Viewbangmau::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		CPoint ptMouse;
		GetCursorPos(&ptMouse);
		int X = (int)ptMouse.x;
		int Y = (int)ptMouse.y;

		if (iWPar == VK_ESCAPE || (X < dPosX || Y < dPosY || X > dPosX + dWidthF || Y > dPosY + dHeightF))
		{
			OnBnClickedButton1();
			return TRUE;
		}
	}

	return FALSE;
}

void Dlg_Viewbangmau::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) OnBnClickedButton1();
	else CDialog::OnSysCommand(nID, lParam);
}

void Dlg_Viewbangmau::SetLocation()
{
	try
	{
		CPoint point;
		GetCursorPos(&point);

		CRect crt;
		GetWindowRect(&crt);
		ScreenToClient(&crt);

		dWidthF = crt.Width();
		dHeightF = crt.Height();

		dPosX = (int)point.x + 50;
		dPosY = (int)point.y - 50;

		this->MoveWindow(dPosX, dPosY, 
			crt.right - crt.left, crt.bottom - crt.top, TRUE);
	}
	catch (...) {}
}

void Dlg_Viewbangmau::LoadList()
{
	try
	{
		// Xác định số lượng
		int num = _wtoi(szNumRow);
		if (num <= 0) return;

		_bstr_t shBMau = name_sheet((_bstr_t)"shBangmauNT");
		_WorksheetPtr psBMau = xl->Sheets->GetItem(shBMau);
		RangePtr prBMau = psBMau->Cells;
		int pRowEnd = prBMau->SpecialCells(xlCellTypeLastCell)->GetRow();

		CString szval = L"";
		int numEnd = pRowEnd;
		for (int i = num + 1; i <= pRowEnd; i++)
		{
			szval = GIOR(i, 1, prBMau, L"");
			szval.Trim();
			if (szval == L"END")
			{
				numEnd = i;
				break;
			}
		}

		if (num + 20 < numEnd) numEnd = num + 20;

		// Xác định các cột được sử dụng
		int slcot=0, icheck = 0;
		
		CString szTitle[27] = { L"C",L"D", L"E", L"F", L"G", L"H", L"I", 
			L"J", L"K", L"L", L"M",	L"N", L"O", L"P", L"Q", L"R", L"S", L"T", 
			L"U", L"V", L"W", L"X", L"Y", L"Z",	L"AA", L"AB", L"AC" };

		for (int i = 3; i <= 29; i++)
		{
			icheck = 0;
			for (int j = num; j <= numEnd; j++)
			{
				szval = GIOR(j, i , prBMau, L"");
				if (szval != L"")
				{
					icheck = 1;
					break;
				}
			}

			if (icheck == 1)
			{
				slcot++;
				MSDSC[slcot].icol = i;
				MSDSC[slcot].szten = szTitle[i-3];
			}
		}

		if (slcot == 0) return;
		MSDSC[slcot+1].icol = MSDSC[slcot].icol+1;

		// Xác định độ rộng mỗi cột	
		for (int i = 1; i <= slcot; i++)
		{
			if (MSDSC[i+1].icol > MSDSC[i].icol + 1)
			{
				MSDSC[i].iwidth = 30 * (MSDSC[i+1].icol - MSDSC[i].icol);
			}
			else MSDSC[i].iwidth = 50;
		}

		// Xác định số lượng dòng chứa tiêu đề
		int icong = 0;
		for (int i = num+1; i <= numEnd; i++)
		{
			szval = GIOR(i, MSDSC[1].icol, prBMau, L"");
			if (szval == L"") icong++;
			else break;
		}		

		for (int i = 0; i < slcot; i++)
		{
			m_list.InsertColumn(i, MSDSC[i+1].szten, LVCFMT_LEFT, MSDSC[i + 1].iwidth);
		}

		m_list.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);

		int dem = 0;
		for (int i = num; i <= numEnd; i++)
		{
			szval = GIOR(i, MSDSC[1].icol, prBMau, L"");
			m_list.InsertItem(dem, szval, 0);
			dem++;
		}

		if (slcot > 1)
		{
			for (int i = 1; i <= slcot; i++)
			{
				dem = 0;
				for (int j = num; j <= numEnd; j++)
				{
					szval = GIOR(j, MSDSC[i+1].icol, prBMau, L"");
					m_list.SetItemText(dem, i, szval);
					dem++;
				}
			}
		}		

		for (int i = 0; i <= icong; i++)
		{
			m_list.SetRowColors(i, RGB(255, 255, 153), RGB(0, 0, 0));
		}		
	}
	catch (...){}
}

void Dlg_Viewbangmau::OnBnClickedButton1()
{
	CDialogEx::OnCancel();
}
