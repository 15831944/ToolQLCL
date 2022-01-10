#include "stdafx.h"
#include "Dlg_TimeDM.h"

// Dlg_TimeDM dialog
IMPLEMENT_DYNAMIC(Dlg_TimeDM, CDialogEx)

Dlg_TimeDM::Dlg_TimeDM(CWnd* pParent /*=NULL*/)
	: CDialogEx(Dlg_TimeDM::IDD, pParent)
{

}

Dlg_TimeDM::~Dlg_TimeDM()
{
}

void Dlg_TimeDM::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, TDM_T1_V1, m_txt[0]);
	DDX_Control(pDX, TDM_T1_V2, m_txt[1]);
	DDX_Control(pDX, TDM_T1_V3, m_txt[2]);
	DDX_Control(pDX, TDM_T1_V4, m_txt[3]);

	DDX_Control(pDX, TDM_SP1_V1, m_spin[0]);
	DDX_Control(pDX, TDM_SP1_V2, m_spin[1]);
	DDX_Control(pDX, TDM_SP1_V3, m_spin[2]);
	DDX_Control(pDX, TDM_SP1_V4, m_spin[3]);

	DDX_Control(pDX, TDM_L1, m_list[0]);
	DDX_Control(pDX, TDM_L2, m_list[1]);

	DDX_Control(pDX, TDM_CK1, m_chk1);
	DDX_Control(pDX, TDM_S2, m_chk2);	
}


BEGIN_MESSAGE_MAP(Dlg_TimeDM, CDialogEx)
	ON_BN_CLICKED(TDM_B2, &Dlg_TimeDM::OnBnClickedB2)
	ON_NOTIFY(UDN_DELTAPOS, TDM_SP1_V1, &Dlg_TimeDM::OnDeltaposSp1V1)
	ON_NOTIFY(UDN_DELTAPOS, TDM_SP1_V2, &Dlg_TimeDM::OnDeltaposSp1V2)
	ON_NOTIFY(UDN_DELTAPOS, TDM_SP1_V3, &Dlg_TimeDM::OnDeltaposSp1V3)
	ON_NOTIFY(UDN_DELTAPOS, TDM_SP1_V4, &Dlg_TimeDM::OnDeltaposSp1V4)
	ON_BN_CLICKED(TDM_B1, &Dlg_TimeDM::OnBnClickedB1)
	ON_EN_KILLFOCUS(TDM_T1_V1, &Dlg_TimeDM::OnEnKillfocusT1V1)
	ON_EN_KILLFOCUS(TDM_T1_V3, &Dlg_TimeDM::OnEnKillfocusT1V3)
	ON_EN_KILLFOCUS(TDM_T1_V2, &Dlg_TimeDM::OnEnKillfocusT1V2)
	ON_EN_KILLFOCUS(TDM_T1_V4, &Dlg_TimeDM::OnEnKillfocusT1V4)
	ON_LBN_DBLCLK(TDM_L1, &Dlg_TimeDM::OnLbnDblclkL1)
	ON_LBN_DBLCLK(TDM_L2, &Dlg_TimeDM::OnLbnDblclkL2)
	ON_LBN_SELCHANGE(TDM_L1, &Dlg_TimeDM::OnLbnSelchangeL1)
	ON_LBN_SELCHANGE(TDM_L2, &Dlg_TimeDM::OnLbnSelchangeL2)
	ON_BN_CLICKED(TDM_S2, &Dlg_TimeDM::OnBnClickedS2)
END_MESSAGE_MAP()


// Dlg_TimeDM message handlers
BOOL Dlg_TimeDM::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);
	
	// Biến kiểm tra giới hạn các cận
	iChk=0, iNext=0;

	m_chk2.SetCheck(iCheckTimerDM);

	// Xác định vị trí hiển thị hộp thoại
	f_Set_location();

	// Load listbox
	f_Load_listbox();

	// Xét các nút spin
	f_Set_spin();

	// Lấy thời gian hiện hành
	f_Set_time();	

	return TRUE; 
}


// Bắt sự kiện click Enter
BOOL Dlg_TimeDM::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if (pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN && pWndWithFocus == GetDlgItem(TDM_L1))
	{
		f_Get_time(0);
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN && pWndWithFocus == GetDlgItem(TDM_L2))
	{
		f_Get_time(1);
		return TRUE;
	}	
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(TDM_T1_V1))
	{
		iNext=0;
		GotoDlgCtrl(GetDlgItem(TDM_T1_V3));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(TDM_T1_V3))
	{
		iNext=1;
		GotoDlgCtrl(GetDlgItem(TDM_T1_V2));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(TDM_T1_V2))
	{
		iNext=0;
		GotoDlgCtrl(GetDlgItem(TDM_T1_V4));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(TDM_T1_V4))
	{
		iNext=0;
		GotoDlgCtrl(GetDlgItem(TDM_T1_V1));
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		CDialog::OnCancel();
		return TRUE; 
	}

	return FALSE;
}


void Dlg_TimeDM::OnBnClickedB2()
{
	CDialog::OnCancel();
}


void Dlg_TimeDM::f_Set_location()
{
	try
	{
		// Xác định vị trí chỉ chuột
		POINT p;
		if (GetCursorPos(&p))
		{
			CRect r;
			GetWindowRect(&r);
			ScreenToClient(&r);

			// Xác định kích thước hộp thoại
			int h = r.Height();
			int w = r.Width();

			// Xác định  kích thước màn hình
			int isx = GetSystemMetrics(SM_CXSCREEN);
			int isy = GetSystemMetrics(SM_CYSCREEN);

			// Xác định vị trí hiển thị hộp thoại
			if(p.x+w+10 >= isx) p.x-=(w+10);
			if(p.y+h-100 >= isy) p.y-=50;
			MoveWindow(p.x + 10, p.y - 200 , r.right - r.left, r.bottom - r.top, TRUE);
		}
	}
	catch(...){}
}


void Dlg_TimeDM::f_Load_listbox()
{
	try
	{
		// Xác định tất cả ngày có trên sheet danh mục
		pSh1 = pWb->ActiveSheet;
		pRg1 = pSh1->Cells;

		_irow = pSh1->Application->ActiveCell->Row;
		_icol = pSh1->Application->ActiveCell->Column;

		int pos=-1;
		int iktr=0,slg=0;
		CString sLuu[100];
		CString szval=L"";
		int iEnd = FindComment(pSh1,1,(_bstr_t)"END");
		for (int i = 8; i < iEnd; i++)
		{
			szval = GIOR(i,_icol,pRg1,L"");
			szval.Trim();
			if(szval!=L"")
			{
				pos = FindMulti(szval,L"h|-");
				if(pos>0)
				{
					iktr=0;
					if(slg>1)
					{
						for (int j = 1; j <= slg; j++)
						{
							if(szval==sLuu[j])
							{
								iktr=1;
								break;
							}
						}
					}

					if(iktr==0 && slg<99)
					{
						if(szval.Mid(1,1)==L"h") szval = L"0" + szval;
						if(_wtoi(szval.Left(2))<=12) m_list[0].AddString(szval);
						else m_list[1].AddString(szval);

						slg++;
						sLuu[slg] = szval;
					}
				}
			}
		}

		irts = 3+getRow(psTS,"CF_EVDATE");
		icts = 1+getColumn(psTS,"CF_EVDATE");
		int dem=0;
		for (int i = 0; i < 100; i++)
		{
			szval = GIOR(irts+i,icts,prTS,L"");
			szval.Trim();
			if(szval!=L"")
			{
				dem=0;
				pos = FindMulti(szval,L"h|-");
				if(pos>0)
				{
					iktr=0;
					if(slg>1)
					{
						for (int j = 1; j <= slg; j++)
						{
							if(szval==sLuu[j])
							{
								iktr=1;
								break;
							}
						}
					}

					if(iktr==0 && slg<99)
					{
						if(szval.Mid(1,1)==L"h") szval = L"0" + szval;
						if(_wtoi(szval.Left(2))<=12) m_list[0].AddString(szval);
						else m_list[1].AddString(szval);

						slg++;
						sLuu[slg] = szval;
					}
				}
			}
			else dem++;

			if(dem==10) break;
		}
	}
	catch(...){}
}


void Dlg_TimeDM::f_Set_spin()
{
	try
	{
		// Xét các nút spin số 1 - Giờ
		m_spin[0].SetDecimalPlaces (0);		// Làm tròn sau i số thập phân
		m_spin[0].SetTrimTrailingZeros (FALSE);		// Xóa số 0 đằng sau số thập phân (nếu có)
		m_spin[0].SetRangeAndDelta (0, 23, 1);	// Giá trị min-max & bước nhảy
		m_spin[0].SetPos (8);	// Giá trị hiển thị

		// Xét các nút spin số 2 - Giờ
		m_spin[1].SetDecimalPlaces (0);
		m_spin[1].SetTrimTrailingZeros (FALSE);
		m_spin[1].SetRangeAndDelta (0, 23, 1);
		m_spin[1].SetPos (8);

		// Xét các nút spin số 3 - Phút
		m_spin[2].SetDecimalPlaces (0);
		m_spin[2].SetTrimTrailingZeros (FALSE);
		m_spin[2].SetRangeAndDelta (0, 45, 15);
		m_spin[2].SetPos (0);

		// Xét các nút spin số 4 - Phút
		m_spin[3].SetDecimalPlaces (0);
		m_spin[3].SetTrimTrailingZeros (FALSE);
		m_spin[3].SetRangeAndDelta (0, 45, 15);
		m_spin[3].SetPos (30);
	}
	catch(...){}
}


void Dlg_TimeDM::f_Set_time()
{
	try
	{
		CString szval=L"";
		szval = GIOR(_irow,_icol,pRg1,L"");
		szval.Trim();

		// gTime
		if(szval!=L"") f_Tack_thoigian(szval);

		// Load txt
		if(gTime[0]>=0 && gTime[0]<24) m_spin[0].SetPos (gTime[0]);	// giờ bd
		if(gTime[1]>=0 && gTime[1]<24) m_spin[1].SetPos (gTime[1]);	// giờ kt
		if(gTime[2]>=0 && gTime[2]<60) m_spin[2].SetPos (gTime[2]);	// phút bd
		if(gTime[3]>=0 && gTime[3]<60) m_spin[3].SetPos (gTime[3]);	// phút kt
	}
	catch(...){}
}


void Dlg_TimeDM::f_Tack_thoigian(CString stime)
{
	try
	{
		int pos = stime.Find(L"-");
		if(pos>0)
		{
			// Tách thời gian bắt đầu & kết thúc
			CString st[2]={L"",L""};
			st[0] = stime.Left(pos);
			st[1] = stime.Right(stime.GetLength()-pos-1);

			// Tách bắt đầu
			pos = st[0].Find(L"h");
			if(pos>0)
			{
				gTime[0] = _wtoi(st[0].Left(pos));
				gTime[2] = _wtoi(st[0].Right(st[0].GetLength()-pos-1));
			}

			// Tách kết thúc
			pos = st[1].Find(L"h");
			if(pos>0)
			{
				gTime[1] = _wtoi(st[1].Left(pos));
				gTime[3] = _wtoi(st[1].Right(st[1].GetLength()-pos-1));
			}				
		}
	}
	catch(...){}
}


void Dlg_TimeDM::f_Tack_time2(CString stime)
{
	try
	{
		int pos = stime.Find(L"-");
		if(pos>0)
		{
			// Tách thời gian bắt đầu & kết thúc
			int igt[4]={-1,-1,-1,-1};
			CString st[2]={L"",L""};
			st[0] = stime.Left(pos);
			st[1] = stime.Right(stime.GetLength()-pos-1);

			// Tách bắt đầu
			pos = st[0].Find(L"h");
			if(pos>0)
			{
				igt[0] = _wtoi(st[0].Left(pos));
				igt[2] = _wtoi(st[0].Right(st[0].GetLength()-pos-1));
			}

			// Tách kết thúc
			pos = st[1].Find(L"h");
			if(pos>0)
			{
				igt[1] = _wtoi(st[1].Left(pos));
				igt[3] = _wtoi(st[1].Right(st[1].GetLength()-pos-1));
			}

			// Load txt
			if(igt[0]>=0 && igt[0]<24) m_spin[0].SetPos (igt[0]);	// giờ bd
			if(igt[1]>=0 && igt[1]<24) m_spin[1].SetPos (igt[1]);	// giờ kt
			if(igt[2]>=0 && igt[2]<60) m_spin[2].SetPos (igt[2]);	// phút bd
			if(igt[3]>=0 && igt[3]<60) m_spin[3].SetPos (igt[3]);	// phút kt
		}
	}
	catch(...){}
}

void Dlg_TimeDM::f_Check_spin(int type, int num)
{
	try
	{
		CString szval=L"";
		m_txt[type].GetWindowTextW(szval);
		szval.Trim();
		int N = _wtoi(szval);
		if(N==0 || N==num) iChk++;
		else iChk=0;

		if(iChk==2)
		{
			iChk=0;
			szval.Format(L"%d",num);
			if(N==0) m_txt[type].SetWindowText(szval);
			else m_txt[type].SetWindowText(L"0");
		}
	}
	catch(...){}
}


void Dlg_TimeDM::f_Check_int(int type, int num)
{
	try
	{
		CString szval=L"";
		m_txt[type].GetWindowTextW(szval);
		szval.Trim();
		int N = _wtoi(szval);
		if(N<0) N=0;
		else if(N>num) N=num;
		szval.Format(L"%d",N);
		m_txt[type].SetWindowText(szval);
	}
	catch(...){}
}


void Dlg_TimeDM::f_Auto_time()
{
	try
	{
		CString sz1=L"";
		m_txt[0].GetWindowTextW(sz1);
		sz1.Trim();
		int N1 = _wtoi(sz1);

		CString sz2=L"";
		m_txt[2].GetWindowTextW(sz2);
		sz2.Trim();
		int N2 = _wtoi(sz2);
		N2+=30;

		if(N2>=60)
		{
			N2-=60;
			N1++;
			if(N1==24) N1=0;			
		}

		sz1.Format(L"%d",N1);
		sz2.Format(L"%d",N2);
		m_txt[1].SetWindowText(sz1);
		m_txt[3].SetWindowText(sz2);
	}
	catch(...){}
}


void Dlg_TimeDM::OnDeltaposSp1V1(NMHDR *pNMHDR, LRESULT *pResult)
{
	f_Check_spin(0,23);
	f_Auto_time();
}


void Dlg_TimeDM::OnDeltaposSp1V2(NMHDR *pNMHDR, LRESULT *pResult)
{
	f_Check_spin(1,23);	
}


void Dlg_TimeDM::OnDeltaposSp1V3(NMHDR *pNMHDR, LRESULT *pResult)
{
	f_Check_spin(2,45);
	f_Auto_time();
}


void Dlg_TimeDM::OnDeltaposSp1V4(NMHDR *pNMHDR, LRESULT *pResult)
{
	f_Check_spin(3,45);
}


void Dlg_TimeDM::f_Get_time(int num)
{
	try
	{
		CString szval=L"";
		m_list[num].GetText(m_list[num].GetCurSel(), szval);
		szval.Trim();
		pRg1->PutItem(_irow,_icol,(_bstr_t)szval);
		f_Tack_thoigian(szval);
		f_Luu_time_TS(szval);
		CDialog::OnOK();
	}
	catch(...){}
}


void Dlg_TimeDM::OnBnClickedB1()
{
	_irow = pSh1->Application->ActiveCell->Row;
	_icol = pSh1->Application->ActiveCell->Column;

	CString szval[4];
	for (int i = 0; i < 4; i++)
	{
		m_txt[i].GetWindowTextW(szval[i]);
		szval[i].Trim();
		gTime[i] = _wtoi(szval[i]);
		if(gTime[i]<=9) szval[i] = L"0" + szval[i];
	}
	
	CString sTime=L"";
	iCheckTimerDM = m_chk2.GetCheck();
	if(iCheckTimerDM==1) sTime.Format(L"%sh%s-%sh%s",szval[0],szval[2],szval[1],szval[3]);
	else sTime.Format(L"%sh%s",szval[0],szval[2]);

	pRg1->PutItem(_irow,_icol,(_bstr_t)sTime);	
	f_Luu_time_TS(sTime);
	
	int nck = m_chk1.GetCheck();
	if(nck==0) CDialog::OnOK();
}


void Dlg_TimeDM::f_Luu_time_TS(CString sLuu)
{
	try
	{
		CString szval=L"";
		int count=0;
		for (int i = 0; i <= 1; i++)
		{
			count = m_list[i].GetCount();
			if(count>0)
			for (int j = 0; j < count; j++)
			{
				m_list[i].GetText(j,szval);
				if(szval==sLuu) return;
			}
		}

		int vtri=4,dem=0;
		for (int i = irts; i < irts+100; i++)
		{
			szval = GIOR(i,icts,prTS,L"");
			szval.Trim();
			if(szval==L"") dem++;
			else dem=0;

			if(dem==10)
			{
				vtri = i-9;
				break;
			}
		}

		prTS->PutItem(vtri,icts,(_bstr_t)sLuu);		
	}
	catch(...){}
}

void Dlg_TimeDM::f_Delete_list(int num)
{
	m_list[num].ResetContent();
	ASSERT(m_list[num].GetCount() == 0);
}


void Dlg_TimeDM::OnEnKillfocusT1V1()
{
	iNext=0;
	f_Check_int(0,23);
	f_Auto_time();
}


void Dlg_TimeDM::OnEnKillfocusT1V3()
{
	f_Check_int(2,59);
	f_Auto_time();
	if(iNext==1) GotoDlgCtrl(GetDlgItem(TDM_T1_V2));
	iNext=0;
}


void Dlg_TimeDM::OnEnKillfocusT1V2()
{
	iNext=0;
	f_Check_int(1,23);
}


void Dlg_TimeDM::OnEnKillfocusT1V4()
{
	iNext=0;
	f_Check_int(3,59);
}


void Dlg_TimeDM::OnLbnDblclkL1()
{
	f_Get_time(0);
}


void Dlg_TimeDM::OnLbnDblclkL2()
{
	f_Get_time(1);
}


void Dlg_TimeDM::OnLbnSelchangeL1()
{
	CString szval=L"";
	m_list[0].GetText(m_list[0].GetCurSel(), szval);
	szval.Trim();
	f_Tack_time2(szval);
}


void Dlg_TimeDM::OnLbnSelchangeL2()
{
	CString szval=L"";
	m_list[1].GetText(m_list[0].GetCurSel(), szval);
	szval.Trim();
	f_Tack_time2(szval);
}

void Dlg_TimeDM::OnBnClickedS2()
{
	iCheckTimerDM = m_chk2.GetCheck();
}
