#include "Dlg_Setup_Print.h"

Dlg_Setup_Print::Dlg_Setup_Print() : CDialog(Dlg_Setup_Print::IDD)
{
}

void Dlg_Setup_Print::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, STP_C1, fcbb[1]);
	DDX_Control(pDX, STP_C2, fcbb[2]);
	DDX_Control(pDX, STP_C3, fcbb[3]);
	DDX_Control(pDX, STP_C4, fcbb[4]);
	DDX_Control(pDX, STP_C5, fcbb[5]);
	DDX_Control(pDX, STP_C6, fcbb[6]);
	DDX_Control(pDX, STP_C7, fcbb[7]);
	DDX_Control(pDX, STP_C8, fcbb[8]);
	DDX_Control(pDX, STP_C9, fcbb[9]);
	DDX_Control(pDX, STP_C10, fcbb[10]);
	DDX_Control(pDX, STP_C11, fcbb[11]);
	DDX_Control(pDX, STP_C12, fcbb[12]);
	DDX_Control(pDX, STP_T1, txt[1]);
	DDX_Control(pDX, STP_T2, txt[2]);
	DDX_Control(pDX, STP_T3, txt[3]);
	DDX_Control(pDX, STP_T4, txt[4]);
	DDX_Control(pDX, STP_CK1, fchk[1]);
	DDX_Control(pDX, STP_CK2, fchk[2]);
	DDX_Control(pDX, STP_CK3, fchk[3]);
	DDX_Control(pDX, STP_CK4, fchk[4]);
	DDX_Control(pDX, STP_CK5, fchk[5]);
	DDX_Control(pDX, STP_CK6, fchk[6]);
	DDX_Control(pDX, STP_CK7, fchk[7]);
	DDX_Control(pDX, STP_CK8, fchk[8]);
	DDX_Control(pDX, STP_CK9, fchk[9]);
	DDX_Control(pDX, STP_CK10, fchk[10]);
	DDX_Control(pDX, STP_CK11, fchk[11]);
	DDX_Control(pDX, STP_CK12, fchk[12]);
	DDX_Control(pDX, STP_CK13, fchk[13]);
}

BEGIN_MESSAGE_MAP(Dlg_Setup_Print, CDialog)
	ON_BN_CLICKED(STP_B1, &Dlg_Setup_Print::OnBnClickedB1)
	ON_BN_CLICKED(STP_B2, &Dlg_Setup_Print::OnBnClickedB2)
	ON_BN_CLICKED(STP_CK6, &Dlg_Setup_Print::OnBnClickedCk6)
	ON_BN_CLICKED(STP_CK7, &Dlg_Setup_Print::OnBnClickedCk7)
	ON_BN_CLICKED(STP_CK8, &Dlg_Setup_Print::OnBnClickedCk8)
	ON_BN_CLICKED(STP_CK9, &Dlg_Setup_Print::OnBnClickedCk9)
	ON_BN_CLICKED(STP_CK10, &Dlg_Setup_Print::OnBnClickedCk10)
	ON_EN_KILLFOCUS(STP_T3, &Dlg_Setup_Print::OnEnKillfocusT3)
	ON_BN_CLICKED(STP_CK11, &Dlg_Setup_Print::OnBnClickedCk11)
	ON_EN_KILLFOCUS(STP_T4, &Dlg_Setup_Print::OnEnKillfocusT4)
	ON_BN_CLICKED(STP_CK12, &Dlg_Setup_Print::OnBnClickedCk12)
	ON_BN_CLICKED(STP_CK1, &Dlg_Setup_Print::OnBnClickedCk1)
	ON_BN_CLICKED(STP_B3, &Dlg_Setup_Print::OnBnClickedB3)
	ON_BN_CLICKED(STP_CK13, &Dlg_Setup_Print::OnBnClickedCk13)
	ON_CBN_SELCHANGE(STP_C12, &Dlg_Setup_Print::OnCbnSelchangeC12)
END_MESSAGE_MAP()

BOOL Dlg_Setup_Print::OnInitDialog()
{
CDialog::OnInitDialog();
	// Load Icon dự toán
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	f_Load_ToolTip();
	f_thietlap_macdinh();

	return TRUE; 
}

// Bắt sự kiện click Enter
BOOL Dlg_Setup_Print::PreTranslateMessage(MSG* pMsg)
{
	 CWnd* pWndWithFocus = GetFocus();
	if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(STP_B1))
	{
		OnBnClickedB1();
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(STP_T1))
	{
		GotoDlgCtrl(GetDlgItem(STP_T2));
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(STP_T1))
	{
		GotoDlgCtrl(GetDlgItem(STP_B2));
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		CDialog::OnCancel();
		return TRUE; 
	}
	else if( pMsg->message == WM_LBUTTONDOWN ||
		pMsg->message == WM_LBUTTONUP ||
		pMsg->message == WM_MOUSEMOVE )
	{
		m_ToolTips.RelayEvent(pMsg);
		return CDialog::PreTranslateMessage(pMsg);
	}

	return FALSE;
}


void Dlg_Setup_Print::f_Load_ToolTip()
	{
		
		CButton * chk0 = (CButton *)GetDlgItem(STP_CK1);
		CButton * chk1 = (CButton *)GetDlgItem(STP_CK5);
		CButton * chk12 = (CButton *)GetDlgItem(STP_CK12);
		
		m_ToolTips.Create(this);
		m_ToolTips.SetMaxTipWidth(600);
		m_ToolTips.SetDelayTime(900);

		m_ToolTips.AddTool(chk0, L"Tự động đánh số thứ tự trang in. "
			L"\nCó thể đánh số trang bắt đầu bằng 1 số bất kỳ.");

		if(iDlog==0)
			m_ToolTips.AddTool(chk12, L"Khi tiến hành in hồ sơ, tích để đánh số trang lại từ đầu "
			L"\nkhi có nhiều bộ hồ sơ, biên bản hoặc số lượng công việc nhiều.");
		else
			m_ToolTips.AddTool(chk12, L"Khi tiến hành in nhật ký, "
				L"\ntích để đánh số trang lại từ đầu bắt đầu ngày mới.");

		m_ToolTips.AddTool(chk1, L"Trong quá trình tính toán, có thể xảy ra các lỗi không hiện giá trị."
			L"\nBạn có thể tích chọn để bỏ trống các giá trị đó khi in."); 
	}


void Dlg_Setup_Print::OnBnClickedB1()
{
	try
	{
		int n0 = _myPrint.inumpage;
		int n1 = _myPrint.isecpage;

		// num page
		CButton *m_ctlCheck1 = (CButton*) GetDlgItem(STP_CK1);
		if(m_ctlCheck1->GetCheck() == BST_UNCHECKED) n0=0;
		else n0=1;

		n1 = fcbb[11].GetCurSel();
		if(n0==1 && _ick[2]==n1) 
			MessageBox(L"Vị trí tiêu đề chân trang & đánh số trang bị trùng nhau.",
			L"Thông báo",MB_OK | MB_ICONWARNING);
		else
		{
			_myPrint.itop = f_setCombobox(1);
			_myPrint.ileft = f_setCombobox(2);
			_myPrint.iright = f_setCombobox(3);
			_myPrint.ibottom = f_setCombobox(4);
			_myPrint.ihtop = f_setCombobox(9);
			_myPrint.ifbot = f_setCombobox(10);
			
			_myPrint.isizehd = f_setCombobox(5);
			_myPrint.isizeft = f_setCombobox(7);

			_myPrint.isecpage = n1;
			if(iDlog==0) _myPrint.inextpage = fcbb[12].GetCurSel();

			_myPrint.istylehd = fcbb[6].GetCurSel();
			if(_myPrint.istylehd==1) _myPrint.cstylehd = L"Bold";
			else if(_myPrint.istylehd==2) _myPrint.cstylehd = L"Italic";
			else if(_myPrint.istylehd==3) _myPrint.cstylehd = L"Bold Italic";
			else _myPrint.cstylehd = L"";

			_myPrint.istyleft = fcbb[8].GetCurSel();
			if(_myPrint.istyleft==1) _myPrint.cstyleft = L"Bold";
			else if(_myPrint.istyleft==2) _myPrint.cstyleft = L"Italic";
			else if(_myPrint.istyleft==3) _myPrint.cstyleft = L"Bold Italic";
			else _myPrint.cstyleft = L"";

			// num page
			_myPrint.inumpage=n0;

			// hoz
			CButton *m_ctlCheck2 = (CButton*) GetDlgItem(STP_CK2);
			if(m_ctlCheck2->GetCheck() == BST_UNCHECKED) _myPrint.ihor=0;
			else  _myPrint.ihor=1;

			// ver
			CButton *m_ctlCheck3 = (CButton*) GetDlgItem(STP_CK3);
			if(m_ctlCheck3->GetCheck() == BST_UNCHECKED) _myPrint.iver=0;
			else  _myPrint.iver=1;

			// BW
			CButton *m_ctlCheck4 = (CButton*) GetDlgItem(STP_CK4);
			if(m_ctlCheck4->GetCheck() == BST_UNCHECKED) _myPrint.ibwhite=0;
			else  _myPrint.ibwhite=1;

			// Error
			CButton *m_ctlCheck5 = (CButton*) GetDlgItem(STP_CK5);
			if(m_ctlCheck5->GetCheck() == BST_UNCHECKED) _myPrint.ierror=0;
			else  _myPrint.ierror=1;

			// Breaks
			CButton *m_ctlCheck12 = (CButton*) GetDlgItem(STP_CK12);
			if(m_ctlCheck12->GetCheck() == BST_UNCHECKED) _myPrint.ibreaks=0;
			else  _myPrint.ibreaks=1;

			_myPrint.isechd = _ick[1];
			_myPrint.isecft = _ick[2];

			_myPrint.ctxthd = f_Gettext(1);
			_myPrint.ctxtft = f_Gettext(2);
	
			_myPrint.izoom = _wtoi(f_Gettext(3));
			_myPrint.ifirst = _wtoi(f_Gettext(4));
		}

		CDialog::OnOK();
	}
	catch(...){_s(L"#QL6.46: Lỗi lưu định dạng bản in.");}
}


void Dlg_Setup_Print::OnBnClickedB2()
{
	f_thietlap_macdinh();
}


void Dlg_Setup_Print::f_thietlap_macdinh()
{
	// Trạng thái mặc định từng trang in
	if(_myPrint.iprivate==1)
	{
		fchk[13].SetCheck(TRUE);
		f_Enable_private(0);
	}

	// Load dữ liệu combobox
	f_load_combobox();

	// Thông tin tiêu đề
	txt[1].SetWindowText(_myPrint.ctxthd);
	txt[1].SetCueBanner(_T("www.giaxaydung.vn"), TRUE);
	txt[2].SetWindowText(_myPrint.ctxtft);
	txt[2].SetCueBanner(_T("Phần mềm Quản lý chất lượng công trình"), TRUE);
	
	txt[3].SetCueBanner(L"Auto");
	if(_myPrint.iprivate==0)
	{
		if(_myPrint.izoom>0)
		{
			msg.Format(L"%d",_myPrint.izoom);
			txt[3].SetWindowText(msg);
		}
	}

	fchk[1].SetCheck(TRUE);
	fchk[1].EnableWindow(FALSE);
	txt[4].EnableWindow(FALSE);
	fcbb[11].EnableWindow(TRUE);

	msg.Format(L"%d",_myPrint.ifirst);
	txt[4].SetWindowText(msg);

	// Đánh số trang
	fchk[1].SetCheck(_myPrint.inumpage);
	if(_myPrint.inumpage==0)
	{
		txt[4].EnableWindow(FALSE);
		fcbb[11].EnableWindow(FALSE);
	}
	else
	{
		txt[4].EnableWindow(TRUE);
		fcbb[11].EnableWindow(TRUE);
	}

	// Căn chỉnh lề
	fchk[2].SetCheck(_myPrint.ihor);
	fchk[3].SetCheck(_myPrint.iver);
	fchk[4].SetCheck(_myPrint.ibwhite);
	fchk[5].SetCheck(_myPrint.ierror);

	// Ngắt trang
	fchk[12].SetCheck(_myPrint.ibreaks);
	if(_myPrint.ibreaks==0) fcbb[12].EnableWindow(FALSE);
	else fcbb[12].EnableWindow(TRUE);
	
	// Đầu trang
	if(_myPrint.isechd==1)
	{
		fchk[6].SetCheck(0);
		fchk[7].SetCheck(1);
		fchk[8].SetCheck(0);
	}
	else if(_myPrint.isechd==2)
	{
		fchk[6].SetCheck(0);
		fchk[7].SetCheck(0);
		fchk[8].SetCheck(1);
	}
	else
	{
		fchk[6].SetCheck(1);
		fchk[7].SetCheck(0);
		fchk[8].SetCheck(0);
	}

	// Chân trang
	if(_myPrint.isecft==1)
	{
		fchk[9].SetCheck(0);
		fchk[10].SetCheck(1);
		fchk[11].SetCheck(0);
	}
	else if(_myPrint.isecft==2)
	{
		fchk[9].SetCheck(0);
		fchk[10].SetCheck(0);
		fchk[11].SetCheck(1);
	}
	else
	{
		fchk[9].SetCheck(1);
		fchk[10].SetCheck(0);
		fchk[11].SetCheck(0);
	}

	_ick[1]=_myPrint.isechd;
	_ick[2]=_myPrint.isecft;

}

void Dlg_Setup_Print::f_load_combobox()
{
	for (int i = 1; i <= 12; i++)
	{
		fcbb[i].ResetContent();
		ASSERT(fcbb[i].GetCount() == 0);
	}

	// Căn trên dưới
	fcbb[1].AddString(L"");
	fcbb[4].AddString(L"");
	for (int i = 1; i <= 50; i++)
	{
		msg.Format(L"%d",i);
		fcbb[1].AddString(msg);
		fcbb[4].AddString(msg);
	}

	if(_myPrint.iprivate==0) fcbb[1].SetCurSel(_myPrint.itop);
	if(_myPrint.iprivate==0) fcbb[4].SetCurSel(_myPrint.ibottom);

	// Căn trái phải
	fcbb[2].AddString(L"");
	fcbb[3].AddString(L"");
	for (int i = 1; i <= 50; i++)
	{
		msg.Format(L"%d",i);
		fcbb[2].AddString(msg);
		fcbb[3].AddString(msg);
	}

	if(_myPrint.iprivate==0) fcbb[2].SetCurSel(_myPrint.ileft);
	if(_myPrint.iprivate==0) fcbb[3].SetCurSel(_myPrint.iright);

	// Căn lề trên dưới
	fcbb[9].AddString(L"");
	fcbb[10].AddString(L"");
	for (int i = 1; i <= 50; i++)
	{
		msg.Format(L"%d",i);
		fcbb[9].AddString(msg);
		fcbb[10].AddString(msg);
	}

	if(_myPrint.iprivate==0) fcbb[9].SetCurSel(_myPrint.ihtop);
	if(_myPrint.iprivate==0) fcbb[10].SetCurSel(_myPrint.ifbot);

	// Size chữ
	for (int i = 6; i <= 20; i++)
	{
		msg.Format(L"%d",i);
		fcbb[5].AddString(msg);
		fcbb[7].AddString(msg);
	}

	fcbb[5].SetCurSel(_myPrint.isizehd-6);
	fcbb[7].SetCurSel(_myPrint.isizeft-6);

	// Định dạng chữ
	fcbb[6].AddString(L"Mặc định");
	fcbb[6].AddString(L"Chữ đậm");
	fcbb[6].AddString(L"Chữ nghiêng");
	fcbb[6].AddString(L"Chữ đậm & nghiêng");
	fcbb[6].AddString(L"Gạch chân");

	fcbb[8].AddString(L"Mặc định");
	fcbb[8].AddString(L"Chữ đậm");
	fcbb[8].AddString(L"Chữ nghiêng");
	fcbb[8].AddString(L"Chữ đậm & nghiêng");
	fcbb[8].AddString(L"Gạch chân");

	fcbb[6].SetCurSel(_myPrint.istylehd);
	fcbb[8].SetCurSel(_myPrint.istyleft);

	fcbb[11].AddString(L"Trái");
	fcbb[11].AddString(L"Giữa");
	fcbb[11].AddString(L"Phải");
	fcbb[11].SetCurSel(_myPrint.isecpage);

	if(iDlog==0)
	{
		fcbb[12].AddString(L"Hồ sơ mới");
		fcbb[12].AddString(L"Công việc mới");
		fcbb[12].AddString(L"Biên bản mới");
		fcbb[12].SetCurSel(_myPrint.inextpage);
	}
	else
	{
		fcbb[12].AddString(L"Ngày mới");
		fcbb[12].SetCurSel(0);
	}
}


int Dlg_Setup_Print::f_setCombobox(int n)
{
	fchon[n] = L"";
	int a = fcbb[n].GetCurSel();
	if(a>=0) fcbb[n].GetLBText(a,fchon[n]);
	return _wtoi(fchon[n]);
}


CString Dlg_Setup_Print::f_Gettext(int n)
{
	CString gt = L"";
	iLen = txt[n].LineLength(txt[n].LineIndex(0));
	lp = gt.GetBuffer(iLen);
	txt[n].GetLine(0, lp,iLen);
	gt.ReleaseBuffer();

	return gt;
}

void Dlg_Setup_Print::OnBnClickedCk6()
{
	_ick[1] = 0;
	fchk[6].SetCheck(TRUE);
	fchk[7].SetCheck(FALSE);
	fchk[8].SetCheck(FALSE);
}


void Dlg_Setup_Print::OnBnClickedCk7()
{
	_ick[1] = 1;
	fchk[6].SetCheck(FALSE);
	fchk[7].SetCheck(TRUE);
	fchk[8].SetCheck(FALSE);
}


void Dlg_Setup_Print::OnBnClickedCk8()
{
	_ick[1] = 2;
	fchk[6].SetCheck(FALSE);
	fchk[7].SetCheck(FALSE);
	fchk[8].SetCheck(TRUE);
}


void Dlg_Setup_Print::OnBnClickedCk9()
{
	_ick[2] = 0;
	fchk[9].SetCheck(TRUE);
	fchk[10].SetCheck(FALSE);
	fchk[11].SetCheck(FALSE);
}


void Dlg_Setup_Print::OnBnClickedCk10()
{
	_ick[2] = 1;
	fchk[9].SetCheck(FALSE);
	fchk[10].SetCheck(TRUE);
	fchk[11].SetCheck(FALSE);
}



void Dlg_Setup_Print::OnBnClickedCk11()
{
	_ick[2] = 2;
	fchk[9].SetCheck(FALSE);
	fchk[10].SetCheck(FALSE);
	fchk[11].SetCheck(TRUE);
}


void Dlg_Setup_Print::OnEnKillfocusT3()
{
	CString ktra = f_Gettext(3);
	if(_wtoi(ktra)<=0) txt[3].SetWindowText(L"");
	else if(_wtoi(ktra)>400) txt[3].SetWindowText(L"400");
	else
	{
		msg.Format(L"%d",_wtoi(ktra));
		txt[3].SetWindowText(msg);
	}
}


void Dlg_Setup_Print::OnEnKillfocusT4()
{
	CString ktra = f_Gettext(4);
	if(_wtoi(ktra)<=0) txt[4].SetWindowText(L"1");
	else if(_wtoi(ktra)>10000) txt[4].SetWindowText(L"10000");
	else
	{
		msg.Format(L"%d",_wtoi(ktra));
		txt[4].SetWindowText(msg);
	}
}


void Dlg_Setup_Print::OnBnClickedCk12()
{
	CButton *m_ctlCheck1 = (CButton*) GetDlgItem(STP_CK12);
	if(m_ctlCheck1->GetCheck() == BST_UNCHECKED) fcbb[12].EnableWindow(FALSE);
	else fcbb[12].EnableWindow(TRUE);
}


void Dlg_Setup_Print::OnBnClickedCk1()
{
	CButton *m_ctlCheck1 = (CButton*) GetDlgItem(STP_CK1);
	if(m_ctlCheck1->GetCheck() == BST_UNCHECKED)
	{
		txt[4].EnableWindow(FALSE);
		fcbb[11].EnableWindow(FALSE);
	}
	else
	{
		txt[4].EnableWindow(TRUE);
		fcbb[11].EnableWindow(TRUE);
	}
}


void Dlg_Setup_Print::OnBnClickedB3()
{
	OnCancel();
}


void Dlg_Setup_Print::OnBnClickedCk13()
{
	CButton *m_ctlCheck1 = (CButton*) GetDlgItem(STP_CK13);
	if(m_ctlCheck1->GetCheck() == BST_CHECKED)
	{
		f_Enable_private(0);
		_myPrint.iprivate=1;
		txt[3].SetWindowText(L"");
		for (int i = 1; i <= 4; i++) fcbb[i].SetCurSel(0);
		fcbb[9].SetCurSel(0);
		fcbb[10].SetCurSel(0);
	}
	else
	{
		f_Enable_private(1);
		_myPrint.iprivate=0;
		if(_myPrint.izoom>0)
		{
			msg.Format(L"%d",_myPrint.izoom);
			txt[3].SetWindowText(msg);
		}

		fcbb[1].SetCurSel(_myPrint.itop);
		fcbb[2].SetCurSel(_myPrint.ileft);
		fcbb[3].SetCurSel(_myPrint.iright);
		fcbb[4].SetCurSel(_myPrint.ibottom);
		fcbb[9].SetCurSel(_myPrint.ihtop);
		fcbb[10].SetCurSel(_myPrint.ifbot);
	}
}


void Dlg_Setup_Print::f_Enable_private(int nType)
{
	fcbb[1].EnableWindow(nType);
	fcbb[2].EnableWindow(nType);
	fcbb[3].EnableWindow(nType);
	fcbb[4].EnableWindow(nType);
	fcbb[9].EnableWindow(nType);
	fcbb[10].EnableWindow(nType);
	txt[3].EnableWindow(nType);
}

void Dlg_Setup_Print::OnCbnSelchangeC12()
{
	// TODO: Add your control notification handler code here
}
