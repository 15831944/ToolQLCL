#include "stdafx.h"
#include "Dlg_Create_Tieude.h"

Dlg_Create_Tieude::Dlg_Create_Tieude(CWnd* pParent /*=NULL*/)
	: CDialog(Dlg_Create_Tieude::IDD, pParent)
{
	sldong = -1;
	nItem = -1, nSubItem = -1;	
	szTitle = L"", szAfter = L"", szBefore = L"";	
}

void Dlg_Create_Tieude::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, CRTD_LIST, litCRETD);
	DDX_Control(pDX, CRTD_EDIT, edCreateTxt);
}


BEGIN_MESSAGE_MAP(Dlg_Create_Tieude, CDialog)
	ON_WM_SIZE()
	ON_WM_SIZING()
	ON_WM_SYSCOMMAND()
	ON_WM_CONTEXTMENU()
	ON_BN_CLICKED(CRTD_BTN_CANCEL, &Dlg_Create_Tieude::OnBnClickedBtnCancel)
	ON_BN_CLICKED(CRTD_BTN_OK, &Dlg_Create_Tieude::OnBnClickedBtnOk)
	ON_COMMAND(ID_T40065, &Dlg_Create_Tieude::OnT40065)
	ON_COMMAND(ID_T40066, &Dlg_Create_Tieude::OnT40066)
	ON_NOTIFY(NM_CLICK, CRTD_LIST, &Dlg_Create_Tieude::OnNMClickList)
	ON_NOTIFY(NM_RCLICK, CRTD_LIST, &Dlg_Create_Tieude::OnNMRClickList)
	ON_EN_SETFOCUS(CRTD_EDIT, &Dlg_Create_Tieude::OnEnSetfocusEdit)
	ON_EN_KILLFOCUS(CRTD_EDIT, &Dlg_Create_Tieude::OnEnKillfocusEdit)
END_MESSAGE_MAP()

BEGIN_EASYSIZE_MAP(Dlg_Create_Tieude)
	EASYSIZE(CRTD_STT_1, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(CRTD_STT_2, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, 0)
	EASYSIZE(CRTD_LIST, ES_BORDER, ES_BORDER, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(CRTD_BTN_OK, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(CRTD_BTN_CANCEL, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
END_EASYSIZE_MAP

// Dlg_Create_Tieude message handlers
BOOL Dlg_Create_Tieude::OnInitDialog()
{
	CDialog::OnInitDialog();

	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	CreateList();

	INIT_EASYSIZE;

	return TRUE;
}

BOOL Dlg_Create_Tieude::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if (pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		if (ClickEsc == 0) OnBnClickedBtnCancel();
		else GotoDlgCtrl(GetDlgItem(CRTD_LIST));
		ClickEsc = 0;

		return TRUE;
	}	

	return FALSE;
}

void Dlg_Create_Tieude::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	UPDATE_EASYSIZE;
}

void Dlg_Create_Tieude::OnSizing(UINT fwSide, LPRECT pRect)
{
	CDialog::OnSizing(fwSide, pRect);
	EASYSIZE_MINSIZE(100, 100, fwSide, pRect);
}

void Dlg_Create_Tieude::OnContextMenu(CWnd* /*pWnd*/, CPoint point)
{
	try
	{
		CMenu _menu;
		CMenu *_pMenu;
		CListCtrlEx *pClist;
		pClist = reinterpret_cast<CListCtrlEx *>(GetDlgItem(CRTD_LIST));
		CRect rectSubmitButton;
		pClist->GetWindowRect(&rectSubmitButton);

		_menu.LoadMenu(IDR_MENU9);
		_pMenu = _menu.GetSubMenu(0);

		if (nItem == -1 || nSubItem == -1)
		{
			_pMenu->EnableMenuItem(ID_T40066, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
		}

		ASSERT(_pMenu);
		if (rectSubmitButton.PtInRect(point))
			_pMenu->TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON, point.x, point.y, this);
	}
	catch (...) {}
}

void Dlg_Create_Tieude::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) OnBnClickedBtnCancel();
	else CDialog::OnSysCommand(nID, lParam);
}

void Dlg_Create_Tieude::OnBnClickedBtnCancel()
{
	sldong = -1;
	CDialog::OnCancel();
}

void Dlg_Create_Tieude::OnBnClickedBtnOk()
{
	sldong = 0;
	szTitle = L"";

	CDialog::OnOK();

	int ncount = litCRETD.GetItemCount();
	int ncolumn = litCRETD.GetColumnCount();

	// Xác định số lượng dòng tiêu đề
	int icheck = 0;
	CString szval = L"";	
	for (int i = 0; i < ncount; i++)
	{
		icheck = 0;
		for (int j = 1; j < ncolumn; j++)
		{
			szval = litCRETD.GetItemText(i, j);
			szval.Trim();
			if (szval != L"")
			{
				icheck = 1;
				break;
			}
		}

		if (icheck == 1) sldong++;
	}

	CString str = L"";
	for (int i = 1; i < ncolumn; i++)
	{
		str = L"";
		for (int j = 0; j < ncount; j++)
		{
			szval = litCRETD.GetItemText(j, i);
			szval.Trim();

			str += szval;
			str += L";";
		}

		szTitle += str;
		szTitle += L"|";
	}
}

void Dlg_Create_Tieude::CreateList()
{
	try
	{
		CString szval = L"";
		litCRETD.InsertColumn(0, L"", LVCFMT_LEFT, 0);

		if (sl >= 1)
		{
			for (int i = 1; i < sl; i++)
			{
				szval.Format(L"Cột %d", i);
				litCRETD.InsertColumn(i, szval, LVCFMT_LEFT, 80);
			}

			for (int i = 0; i < 2; i++)
			{
				litCRETD.InsertItem(i, L"", 0);
				litCRETD.SetItemText(i, 1, L"(Tên)");
			}
		}

		litCRETD.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);
		litCRETD.SetDefaultEditor(NULL, NULL, &edCreateTxt);
		edCreateTxt.SubclassDlgItem(CRTD_EDIT, this); edCreateTxt.SetBkColor(YELLOW153); edCreateTxt.SetTextColor(BLUE);
	}
	catch (...) {}
}

void Dlg_Create_Tieude::OnT40065()
{
	int ncount = litCRETD.GetItemCount();
	litCRETD.InsertItem(ncount, L"", 0);
	litCRETD.SetItemText(ncount, 1, L"(Tên)");
}

void Dlg_Create_Tieude::OnT40066()
{
	POSITION pos = litCRETD.GetFirstSelectedItemPosition();
	if (pos != NULL)
	{
		int ncount = litCRETD.GetItemCount();
		for (int i = ncount - 1; i >=0 ; i--)
		{
			if (litCRETD.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
			{
				litCRETD.DeleteItem(i);
			}
		}
	}
}


void Dlg_Create_Tieude::OnNMClickList(NMHDR *pNMHDR, LRESULT *pResult)
{
	ClickEsc = 0;
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	nItem = pNMListView->iItem;
	nSubItem = pNMListView->iSubItem;
	*pResult = 0;
}


void Dlg_Create_Tieude::OnNMRClickList(NMHDR *pNMHDR, LRESULT *pResult)
{
	ClickEsc = 0;
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	nItem = pNMListView->iItem;
	nSubItem = pNMListView->iSubItem;
	*pResult = 0;
}


void Dlg_Create_Tieude::OnEnSetfocusEdit()
{
	ClickEsc = 1;
}


void Dlg_Create_Tieude::OnEnKillfocusEdit()
{
	ClickEsc = 0;
	szBefore = litCRETD.GetItemText(CLRow, CLCol);
	edCreateTxt.GetWindowTextW(szAfter);

	if (szAfter == szBefore) return;

	int pos = -1;
	CString szval = L"";
	CString strL = L"", strR = L"";
	CString szcl[2] = { szBefore, szAfter };
	COLORREF clrBR[2] = { RGB(255, 255, 255), RGB(250, 191, 143) };
	COLORREF clrBC[2] = { RGB(255, 255, 255), RGB(60, 205, 220) };
	COLORREF clrT[2] = { RGB(0, 0, 0), RGB(0, 0, 0) };

	int ncount = litCRETD.GetItemCount();
	int ncolumn = litCRETD.GetColumnCount();

	for (int k = 0; k < 2; k++)
	{
		pos = szcl[k].Find(L"[");
		if (pos > 0)
		{
			szval = szcl[k].Right(szcl[k].GetLength() - pos - 1);
			pos = szval.Find(L"]");
			if (pos > 0)
			{
				szval = szval.Left(pos);
				szval.Trim();
				szval.MakeUpper();
				if (szval.Find(L"R") ==-1 && szval.Find(L"C") == -1) szval += L"R";

				strL = szval.Left(1);
				strR = szval.Right(1);
				if (strL == L"R" || strL == L"C" || strR == L"R" || strR == L"C")
				{
					if(strL == L"R" || strL == L"C") pos = _wtoi(szval.Right(szval.GetLength() - 1));
					else pos = _wtoi(szval.Left(szval.GetLength() - 1));
					
					if (pos > 0)
					{
						for (int i = 0; i < pos; i++)
						{
							if (strL == L"R" || strR == L"R")
							{
								if (CLRow + i < ncount)
								{
									litCRETD.SetCellColors(CLRow + i, CLCol, clrBR[k], clrT[k]);
									if (i >= 1) litCRETD.SetItemText(CLRow + i, CLCol, L"");
								}
							}
							else
							{
								if (CLCol + i < ncolumn)
								{
									litCRETD.SetCellColors(CLRow, CLCol + i, clrBC[k], clrT[k]);
									if (i >= 1) litCRETD.SetItemText(CLRow, CLCol + i, L"");
								}
							}
						}
					}
				}
			}
		}
	}	
}
