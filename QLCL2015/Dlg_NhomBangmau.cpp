#include "stdafx.h"
#include "Dlg_NhomBangmau.h"
#include "Dlg_Viewbangmau.h"

#define allcviec		L"▶ Tất cả công việc"
#define allctrinh		L"▶ Tất cả đối tượng"
#define tablenull		L"Bảng trắng"
#define khongsudung		L"Không sử dụng bảng mẫu (#)"

// Dlg_NhomBangmau dialog
IMPLEMENT_DYNAMIC(Dlg_NhomBangmau, CDialogEx)

Dlg_NhomBangmau::Dlg_NhomBangmau(CWnd* pParent /*=NULL*/)
	: CDialogEx(Dlg_NhomBangmau::IDD, pParent)
{
	iRowBMau = 0;
	iColBMau = 0;

	iRowActive = 0;
	iColumnActive = 0;
	szTenmau = L"";
}

Dlg_NhomBangmau::~Dlg_NhomBangmau()
{
	vecMSDSM.clear();
}

void Dlg_NhomBangmau::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, NKBMAU_L1, m_list);
	DDX_Control(pDX, NKBMAU_TXT1, m_txt_search);
	DDX_Control(pDX, NKBMAU_CBB1, m_cbb_cviec);
	DDX_Control(pDX, NKBMAU_CBB2, m_cbb_ctrinh);	
}

BEGIN_MESSAGE_MAP(Dlg_NhomBangmau, CDialogEx)
	ON_WM_SIZE()
	ON_WM_SIZING()
	ON_BN_CLICKED(NKBMAU_B1, &Dlg_NhomBangmau::OnBnClickedB1)
	ON_BN_CLICKED(NKBMAU_B2, &Dlg_NhomBangmau::OnBnClickedB2)
	ON_NOTIFY(NM_DBLCLK, NKBMAU_L1, &Dlg_NhomBangmau::OnNMDblclkL1)
	ON_NOTIFY(LVN_KEYDOWN, NKBMAU_L1, &Dlg_NhomBangmau::OnLvnKeydownL1)
	ON_EN_CHANGE(NKBMAU_TXT1, &Dlg_NhomBangmau::OnEnChangeTxt1)
	ON_CBN_SELCHANGE(NKBMAU_CBB1, &Dlg_NhomBangmau::OnCbnSelchangeCbb1)
	ON_CBN_SELCHANGE(NKBMAU_CBB2, &Dlg_NhomBangmau::OnCbnSelchangeCbb2)
	ON_BN_CLICKED(NKBMAU_B5, &Dlg_NhomBangmau::OnBnClickedB5)
	ON_BN_CLICKED(NKBMAU_B4, &Dlg_NhomBangmau::OnBnClickedB4)
	ON_NOTIFY(NM_CLICK, NKBMAU_L1, &Dlg_NhomBangmau::OnNMClickL1)
	ON_BN_CLICKED(IDC_CHECK1, &Dlg_NhomBangmau::OnBnClickedCheck1)
	ON_BN_CLICKED(NKBMAU_B6, &Dlg_NhomBangmau::OnBnClickedB6)
END_MESSAGE_MAP()

BEGIN_EASYSIZE_MAP(Dlg_NhomBangmau)
	EASYSIZE(NKBMAU_S1, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(NKBMAU_S2, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)

	EASYSIZE(NKBMAU_TXT1, ES_BORDER, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(NKBMAU_CBB1, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(NKBMAU_CBB2, ES_BORDER, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(NKBMAU_L1, ES_BORDER, ES_BORDER, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(NKBMAU_B1, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(NKBMAU_B2, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)

	EASYSIZE(IDC_CHECK1, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(NKBMAU_B4, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, 0)
	EASYSIZE(NKBMAU_B5, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, 0)
	EASYSIZE(NKBMAU_B6, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, 0)
END_EASYSIZE_MAP

// Dlg_NhomBangmau message handlers
BOOL Dlg_NhomBangmau::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	m_txt_search.SetCueBanner(L"Nhập từ khóa tìm kiếm...");
	m_cbb_cviec.SetCueBanner(L"-Loại công việc-");
	m_cbb_ctrinh.SetCueBanner(L"-Chọn đối tượng-");
	
	m_cbb_cviec.AddString(allcviec);
	m_cbb_ctrinh.AddString(allctrinh);

	m_cbb_cviec.AdjustDroppedWidth();
	m_cbb_cviec.SetMode(CComboBoxExt::MODE_AUTOCOMPLETE);
	m_cbb_cviec.SetCurSel(0);

	m_cbb_ctrinh.AdjustDroppedWidth();
	m_cbb_ctrinh.SetMode(CComboBoxExt::MODE_AUTOCOMPLETE);
	m_cbb_ctrinh.SetCurSel(0);

	m_list.InsertColumn(0, L"", LVCFMT_CENTER, 0);
	m_list.InsertColumn(1, L"Tên mẫu", LVCFMT_LEFT, _wLBMAU[0]);
	m_list.InsertColumn(2, L"Mô tả", LVCFMT_LEFT, _wLBMAU[1]);
	m_list.InsertColumn(3, L"Công việc", LVCFMT_LEFT, _wLBMAU[2]);
	m_list.InsertColumn(4, L"Đối tượng", LVCFMT_LEFT, _wLBMAU[3]);

	m_list.SetColumnColors(1, RGB(255, 255, 255), RGB(0, 0, 255));
	m_list.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);

	CButton* m_check = (CButton*)GetDlgItem(IDC_CHECK1);
	m_check->SetCheck(jCheckViewBM);

	_LoadDanhSach();

	INIT_EASYSIZE;

	// Set size dialog mfc
	if (irWBMAU > 0 && irHBMAU > 0)
	{
		this->SetWindowPos(NULL, 0, 0, irWBMAU, irHBMAU, SWP_NOMOVE | SWP_NOZORDER);

		CRect r;
		GetWindowRect(&r);
		ScreenToClient(&r);
		int isx = GetSystemMetrics(SM_CXSCREEN);
		int isy = GetSystemMetrics(SM_CYSCREEN);
		MoveWindow((isx - r.Width()) / 2, (isy - r.Height()) / 2, r.right - r.left, r.bottom - r.top, TRUE);
	}

	return TRUE;
}

// Bắt sự kiện click Enter
BOOL Dlg_NhomBangmau::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if (pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(NKBMAU_L1))
	{
		OnBnClickedB1();
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(NKBMAU_TXT1))
	{
		_Timkiemdulieu();
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		CDialog::OnCancel();
		return TRUE;
	}

	return FALSE;
}


void Dlg_NhomBangmau::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	UPDATE_EASYSIZE;
}

void Dlg_NhomBangmau::OnSizing(UINT fwSide, LPRECT pRect)
{
	CDialog::OnSizing(fwSide, pRect);
	EASYSIZE_MINSIZE(320, 240, fwSide, pRect);
}

void Dlg_NhomBangmau::_LoadDanhSach()
{
	slmau = 0;

	int colMota = 31;
	int colCviec = 32;
	int colCtrinh = 33;
	
	shBMau = name_sheet((_bstr_t)"shBangmauNT");
	psBMau = xl->Sheets->GetItem(shBMau);
	prBMau = psBMau->Cells;

	int slcv = 0, slctr = 0;
	CString szCviec[1000], szCtrinh[1000];

	CString szval = L"";
	int dem=0, icheck = 0;
	pRowEnd = prBMau->SpecialCells(xlCellTypeLastCell)->GetRow();

	vecMSDSM.clear();
	MyStrDSMau MSDSM;

	for (int i = 3; i <= pRowEnd; i++)
	{
		szval = GIOR(i, 1, prBMau, L"");
		szval.Trim();
		if (szval != L"" && szval != L"END")
		{
			MSDSM.irow = i;
			MSDSM.szten = szval;
			MSDSM.szmota = GIOR(i, colMota, prBMau, L"");
			MSDSM.szcviec = GIOR(i, colCviec, prBMau, L"");
			MSDSM.szctrinh = GIOR(i, colCtrinh, prBMau, L"");
			vecMSDSM.push_back(MSDSM);

			// Nhóm công việc
			szval = GIOR(i, colCviec, prBMau, L"");
			if (szval != L"")
			{
				icheck = 0;
				if (slcv > 0)
				{
					for (int j = 1; j <= slcv; j++)
					{
						if (szval == szCviec[j])
						{
							icheck = 1;
							break;
						}
					}
				}

				if (icheck == 0)
				{
					slcv++;
					szCviec[slcv] = szval;
				}
			}

			// Nhóm công trình
			szval = GIOR(i, colCtrinh, prBMau, L"");
			if (szval != L"")
			{
				icheck = 0;
				if (slctr > 0)
				{
					for (int j = 1; j <= slctr; j++)
					{
						if (szval == szCtrinh[j])
						{
							icheck = 1;
							break;
						}
					}
				}

				if (icheck == 0)
				{
					slctr++;
					szCtrinh[slctr] = szval;
				}
			}

			dem = 0;
		}
		else dem++;
	}

	if (slcv > 0)
	{
		for (int i = 1; i <= slcv; i++)
		{
			m_cbb_cviec.AddString(szCviec[i]);
		}
	}

	if (slctr > 0)
	{
		for (int i = 1; i <= slctr; i++)
		{
			m_cbb_ctrinh.AddString(szCtrinh[i]);
		}
	}

	slmau = (int)vecMSDSM.size();
	if (slmau > 0)
	{
		for (int i = 0; i < slmau; i++)
		{
			szval.Format(L"%d", vecMSDSM[i].irow);
			m_list.InsertItem(i, szval, 0);
			m_list.SetItemText(i, 1, vecMSDSM[i].szten);
			m_list.SetItemText(i, 2, vecMSDSM[i].szmota);
			m_list.SetItemText(i, 3, vecMSDSM[i].szcviec);
			m_list.SetItemText(i, 4, vecMSDSM[i].szctrinh);

			if (szTenmau != L"")
			{
				if (szTenmau == vecMSDSM[i].szten)
				{
					m_list.SetRowColors(i, RGB(255, 255, 153), RGB(0, 0, 0));
				}
			}
		}
	}

	m_list.InsertItem(0, L"", 0);	
	m_list.SetItemText(0, 1, tablenull);
	m_list.SetItemText(0, 2, khongsudung);

	if (szTenmau != L"")
	{
		int ivtrScl = -1;
		int ncount = m_list.GetItemCount();
		for (int i = 0; i < ncount; i++)
		{
			szval = m_list.GetItemText(i, 1);
			if (szTenmau == szval)
			{
				ivtrScl = i;
				break;
			}
		}

		if (ivtrScl >= 0)
		{
			RECT r;
			CRect rectCtrl;
			m_list.GetItemRect(0, &r, LVIR_BOUNDS);
			m_list.GetWindowRect(&rectCtrl);
			int a = r.bottom - r.top;
			int b = rectCtrl.Height();

			if (ivtrScl + 10 < ncount)
			{
				m_list.EnsureVisible(ivtrScl + (int)(b / a) - 3, FALSE);
			}
			else
			{
				m_list.EnsureVisible(ncount - 1, FALSE);
			}
		}
	}
}

void Dlg_NhomBangmau::_Timkiemdulieu()
{
	try
	{
		m_list.DeleteAllItems();
		ASSERT(m_list.GetItemCount() == 0);

		CString szval = L"";
		CString sztkhoa = L"";
		m_txt_search.GetWindowTextW(sztkhoa);
		sztkhoa.Trim();

		int nIndexCV = m_cbb_cviec.GetCurSel();
		int nIndexCT = m_cbb_ctrinh.GetCurSel();

		int dem = 0, icheck=1;
		if (slmau > 0)
		{
			for (int i = 0; i < slmau; i++)
			{
				icheck = 1;
				
				if (icheck == 1) icheck = _LocCongviec(nIndexCV, i);
				if (icheck == 1) icheck = _LocCongtrinh(nIndexCT, i);
				if (icheck == 1) icheck = _LocNoidung(sztkhoa, i);
				
				if (icheck == 1)
				{
					szval.Format(L"%d", vecMSDSM[i].irow);
					m_list.InsertItem(dem, szval, 0);
					m_list.SetItemText(dem, 1, vecMSDSM[i].szten);
					m_list.SetItemText(dem, 2, vecMSDSM[i].szmota);
					m_list.SetItemText(dem, 3, vecMSDSM[i].szcviec);
					m_list.SetItemText(dem, 4, vecMSDSM[i].szctrinh);
					dem++;
				}				
			}
		}

		m_list.InsertItem(0, L"", 0);
		m_list.SetItemText(0, 1, tablenull);
		m_list.SetItemText(0, 2, khongsudung);
	}
	catch(...){}
}


int Dlg_NhomBangmau::_LocCongviec(int nIndex, int vt)
{
	if (nIndex < 0) return 1;

	CString szval = L"";
	m_cbb_cviec.GetLBText(nIndex, szval);
	if (szval== allcviec || szval == vecMSDSM[vt].szcviec) return 1;

	return 0;
}


int Dlg_NhomBangmau::_LocCongtrinh(int nIndex, int vt)
{
	if (nIndex < 0) return 1;

	CString szval = L"";
	m_cbb_ctrinh.GetLBText(nIndex, szval);
	if (szval == allctrinh || szval == vecMSDSM[vt].szctrinh) return 1;

	return 0;
}

int Dlg_NhomBangmau::_LocNoidung(CString szNoidung, int vt)
{
	if (szNoidung == L"") return 1;

	szNoidung = _ConvertKhongDau(szNoidung);
	szNoidung.MakeLower();
	_TackTukhoaChuan(szNoidung, _T(" "), _T("+"), 1, 0);

	CString szval = vecMSDSM[vt].szten + L"+" + vecMSDSM[vt].szmota;
	szval = _ConvertKhongDau(szval);
	szval.MakeLower();

	for (int i = 1; i <= iKey; i++)
	{
		if (szval.Find(pKey[i]) == -1) return 0;
	}

	return 1;
}


void Dlg_NhomBangmau::OnBnClickedB1()
{
	POSITION pos = m_list.GetFirstSelectedItemPosition();
	if (pos != NULL)
	{
		for (int i = 0; i < m_list.GetItemCount(); i++)
		{ 
			if (m_list.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
			{
				if (i == 0) sTukhoa = L"#";
				else sTukhoa = m_list.GetItemText(i, 1);

				CDialog::OnOK();

				iRowBMau = 0;
				iColBMau = 0;
				return;
			}
		}
	}
}


void Dlg_NhomBangmau::OnBnClickedB2()
{
	iRowBMau = 0;
	iColBMau = 0;

	for (int i = 0; i < 4; i++)
	{
		_wL1[i] = m_list.GetColumnWidth(i+1);
	}

	f_Get_size();
	CDialog::OnCancel();
}


void Dlg_NhomBangmau::OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	OnBnClickedB1();
	*pResult = 0;
}


void Dlg_NhomBangmau::OnLvnKeydownL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);
	if (pLVKeyDow->wVKey == VK_SPACE) OnBnClickedB1();

	*pResult = 0;
}


void Dlg_NhomBangmau::OnEnChangeTxt1()
{
	_Timkiemdulieu();
}


void Dlg_NhomBangmau::OnCbnSelchangeCbb1()
{
	_Timkiemdulieu();
}


void Dlg_NhomBangmau::OnCbnSelchangeCbb2()
{
	_Timkiemdulieu();
}


void Dlg_NhomBangmau::OnBnClickedB5()
{
	POSITION pos = m_list.GetFirstSelectedItemPosition();
	if (pos != NULL)
	{
		for (int i = 0; i < m_list.GetItemCount(); i++)
		{
			if (m_list.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
			{
				if (i > 0)
				{
					CString szval = m_list.GetItemText(i, 0);
					int num = _wtoi(szval);
					if (num > 0)
					{
						int iVisible = psBMau->GetVisible();
						if (iVisible != -1) psBMau->PutVisible(XlSheetVisibility::xlSheetVisible);
						psBMau->Activate();

						RangePtr PRS = prBMau->GetItem(num, 1);
						PRS->Select();

						xl->ActiveWindow->PutScrollRow(num - 1);

						sTukhoa = L"";
						CDialog::OnOK();

						iRowBMau = iRowActive;
						iColBMau = iColumnActive;
						pShDMBMau = pShDMuc;
					}					
				}

				return;
			}
		}
	}
}


void Dlg_NhomBangmau::OnBnClickedB4()
{
	//int result = _yn(L"Bạn có chắc chắn tạo mới mẫu?", 0, 1);
	//if (result != 6) return;

	int num = 3;
	int ncount = m_list.GetItemCount();
	if (ncount > 0)
	{
		CString szval = m_list.GetItemText(ncount - 1, 0);
		num = _wtoi(szval);
		if (num > 0)
		{
			for (int i = num+1; i <= pRowEnd; i++)
			{
				szval = GIOR(i, 1, prBMau, L"");
				szval.Trim();
				if (szval == L"END")
				{
					num = i;
					break;
				}
			}
		}

		num += 2;
	}

	// Tạo mới
	int iVisible = psBMau->GetVisible();
	if (iVisible != -1) psBMau->PutVisible(XlSheetVisibility::xlSheetVisible);
	psBMau->Activate();

	RangePtr PRS = prBMau->GetItem(num, 1);
	PRS->Select();

	xl->ActiveWindow->PutScrollRow(num-1);

	int icong = 10;
	CString szTen = L"";
	szTen.Format(L"Mẫu số %d", ncount);
	prBMau->PutItem(num, 1, (_bstr_t)szTen);
	prBMau->PutItem(num, 31, (_bstr_t)L"<Mô tả bảng mẫu>");
	prBMau->PutItem(num, 32, (_bstr_t)L"<Loại công việc>");
	prBMau->PutItem(num, 33, (_bstr_t)L"<Đối tượng>");

	prBMau->PutItem(num+icong, 1, (_bstr_t)L"END");

	CString szcmt = L"Ứng với mỗi bảng bạn bắt buộc phải đặt tên cho bảng đó "
		L"để phần mềm có thể gọi được sang các sheet biên bản";

	PRS = prBMau->GetItem(num, 1);
	if (PRS->GetComment() != NULL) PRS->ClearComments();
	PRS->AddComment((_bstr_t)szcmt);
	PRS->Interior->PutColor(RGB(102, 255, 153));

	PRS = prBMau->GetItem(num + icong, 1);
	PRS->Interior->PutColor(RGB(102, 255, 153));

	PRS = psBMau->GetRange(prBMau->GetItem(num, 1), prBMau->GetItem(num + icong, 29));
	PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->Weight = xlThin;
	PRS->Borders->GetItem(XlBordersIndex::xlEdgeLeft)->Weight = xlThin;
	PRS->Borders->GetItem(XlBordersIndex::xlEdgeRight)->Weight = xlThin;
	PRS->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->Weight = xlThin;

	sTukhoa = L"";
	CDialog::OnOK();
}


void Dlg_NhomBangmau::OnNMClickL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	CButton* m_check = (CButton*)GetDlgItem(IDC_CHECK1);
	if ((int)m_check->GetCheck() == 1)
	{
		NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
		int nItem = pNMListView->iItem;
		int nSubItem = pNMListView->iSubItem;

		if (nItem > 0 && nSubItem == 1)
		{
			AFX_MANAGE_STATE(AfxGetStaticModuleState());
			Dlg_Viewbangmau p;
			p.szNumRow = m_list.GetItemText(nItem, 0);
			p.szTenbangmau = m_list.GetItemText(nItem, 1);
			p.DoModal();
		}
	}
}


void Dlg_NhomBangmau::OnBnClickedCheck1()
{
	CButton* m_check = (CButton*)GetDlgItem(IDC_CHECK1);
	jCheckViewBM = m_check->GetCheck();
}

void Dlg_NhomBangmau::f_Get_size()
{
	CRect rectCtrl;
	GetWindowRect(&rectCtrl);
	irHBMAU = rectCtrl.Height();
	irWBMAU = rectCtrl.Width();
}

void Dlg_NhomBangmau::OnBnClickedB6()
{
	try
	{
		int nvtri = 0;
		POSITION pos = m_list.GetFirstSelectedItemPosition();
		if (pos != NULL)
		{
			for (int i = 0; i < m_list.GetItemCount(); i++)
			{
				if (m_list.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					nvtri = i;
					break;
				}
			}
		}

		if (nvtri > 0)
		{
			int result = _yn(L"Bạn có chắc chắn nhân bản mẫu được chọn?", 0, 1);
			if (result == 6)
			{
				CString szval = m_list.GetItemText(nvtri, 0);
				int num = _wtoi(szval);
				if (num > 0)
				{
					int numEnd = num;
					// Xác định vị trí cuối
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

					if (numEnd > num)
					{
						int slchen = numEnd - num + 2;

						RangePtr PRS = psBMau->GetRange(
							prBMau->GetItem(numEnd + 1, 1), prBMau->GetItem(numEnd + slchen, 1));
						PRS->EntireRow->Insert(xlShiftDown);

						PRS = psBMau->GetRange(
							prBMau->GetItem(num, 1), prBMau->GetItem(numEnd, 1));
						PRS->EntireRow->Copy(prBMau->GetItem(numEnd + 2, 1));

						xl->PutCutCopyMode(XlCutCopyMode(false));

						szval.Format(L"%d", num);
						m_list.InsertItem(nvtri + 1, szval, 0);
						m_list.SetItemText(nvtri + 1, 1, m_list.GetItemText(nvtri, 1));
						m_list.SetItemText(nvtri + 1, 2, m_list.GetItemText(nvtri, 2));
						m_list.SetItemText(nvtri + 1, 3, m_list.GetItemText(nvtri, 3));
						m_list.SetItemText(nvtri + 1, 4, m_list.GetItemText(nvtri, 4));

						int ncount = m_list.GetItemCount();

						for (int i = nvtri + 1; i < ncount; i++)
						{
							szval.Format(L"%d", _wtoi(m_list.GetItemText(i, 0)) + slchen);
							m_list.SetItemText(i, 0, szval);
						}

						nvtri = -1;
						ncount = (int)vecMSDSM.size();

						for (int i = 0; i < ncount; i++)
						{
							if (vecMSDSM[i].irow == num)
							{
								nvtri = i;
								break;
							}
						}

						if (nvtri >= 0)
						{
							for (int i = nvtri; i < ncount; i++)
							{
								vecMSDSM[i].irow += slchen;
							}

							MyStrDSMau MSDSM;

							MSDSM.irow = num;
							MSDSM.szten = vecMSDSM[nvtri].szten;
							MSDSM.szmota = vecMSDSM[nvtri].szmota;
							MSDSM.szcviec = vecMSDSM[nvtri].szcviec;
							MSDSM.szctrinh = vecMSDSM[nvtri].szctrinh;

							vecMSDSM.insert(vecMSDSM.begin() + nvtri, MSDSM);
						}
					}
					else
					{
						_s(L"Sao chép không thành công.", 2);
					}
				}
			}
		}
	}
	catch (...) {}
}
