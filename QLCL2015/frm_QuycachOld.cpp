#include "frm_QuycachOld.h"
#include "LaymauNgang.h"

// frm_QuycachOld dialog
IMPLEMENT_DYNAMIC(frm_QuycachOld, CDialogEx)

frm_QuycachOld::frm_QuycachOld(CWnd* pParent /*=NULL*/)
	: CDialogEx(frm_QuycachOld::IDD, pParent)
{
	sslvl=0,sslpp=0;
	szKichthuoc=L"", szTansuat=L"";
}

frm_QuycachOld::~frm_QuycachOld()
{
}

void frm_QuycachOld::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, LMO_CBB_1, m_cbb_ten);
	DDX_Control(pDX, LMO_TXT_1, m_txt[1]);
	DDX_Control(pDX, LMO_TXT_2, m_txt[2]);
	DDX_Control(pDX, LMO_TXT_3, m_txt[3]);
	DDX_Control(pDX, LMO_TXT_4, m_txt[4]);
	DDX_Control(pDX, LMO_TXT_5, m_txt[5]);
	DDX_Control(pDX, LMO_TXT_6, m_txt[6]);
	DDX_Control(pDX, LMO_TXT_GC, m_txt_gc);
	DDX_Control(pDX, LMO_TXT_HD, m_txt_hd);
	DDX_Control(pDX, LMO_CBB_2, m_cbb_qcach);	
	DDX_Control(pDX, LMO_CHKNT, m_chk_nthu);
}


BEGIN_MESSAGE_MAP(frm_QuycachOld, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_CBN_SELCHANGE(LMO_CBB_1, &frm_QuycachOld::OnCbnSelchangeCombo1)
	ON_BN_CLICKED(LMO_BTN_OK, &frm_QuycachOld::OnBnClickedBtnOk)
	ON_BN_CLICKED(LMO_BTN_CANCEL, &frm_QuycachOld::OnBnClickedBtnCancel)
	ON_BN_CLICKED(LMO_BTN_MAUKHAC, &frm_QuycachOld::OnBnClickedBtnMaukhac)
	ON_CBN_SELCHANGE(LMO_CBB_2, &frm_QuycachOld::OnCbnSelchangeCbb2)
	ON_EN_SETFOCUS(LMO_TXT_6, &frm_QuycachOld::OnEnSetfocusTxt6)
	ON_EN_CHANGE(LMO_TXT_1, &frm_QuycachOld::OnEnChangeTxt1)
END_MESSAGE_MAP()


// frm_QuycachOld message handlers
BOOL frm_QuycachOld::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	if(getPathQCLM()==L"") return FALSE;

	XacdinhQuycach();

	return TRUE; 
}

// Bắt sự kiện click Enter
BOOL frm_QuycachOld::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(LMO_TXT_1))
	{
		GotoDlgCtrl(GetDlgItem(LMO_TXT_2));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(LMO_TXT_2))
	{
		GotoDlgCtrl(GetDlgItem(LMO_TXT_3));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(LMO_TXT_3))
	{
		GotoDlgCtrl(GetDlgItem(LMO_TXT_4));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(LMO_TXT_4))
	{
		GotoDlgCtrl(GetDlgItem(LMO_CBB_2));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(LMO_CBB_2))
	{
		GotoDlgCtrl(GetDlgItem(LMO_TXT_5));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(LMO_TXT_5))
	{
		GotoDlgCtrl(GetDlgItem(LMO_TXT_6));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(LMO_TXT_6))
	{
		GotoDlgCtrl(GetDlgItem(LMO_TXT_GC));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(LMO_TXT_GC))
	{
		GotoDlgCtrl(GetDlgItem(LMO_CBB_1));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(LMO_CBB_1))
	{
		GotoDlgCtrl(GetDlgItem(LMO_TXT_1));
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		OnBnClickedBtnCancel();
		return TRUE; 
	}

	return FALSE;
}

void frm_QuycachOld::OnSysCommand(UINT nID, LPARAM lParam)
{
	if(nID == SC_CLOSE) OnBnClickedBtnCancel();
	else CDialog::OnSysCommand(nID, lParam);
}


void frm_QuycachOld::OnCbnSelchangeCombo1()
{
	int nIndex = m_cbb_ten.GetCurSel();
	if(nIndex>=0) Loaddulieu(MSVlieu[nIndex+1].szMa);
}


void frm_QuycachOld::OnBnClickedBtnOk()
{
	CString szNoidung=L"",szval=L"";
	m_cbb_ten.GetWindowTextW(szval);
	szNoidung+=szval;
	szNoidung+=L";";

	m_txt[1].GetWindowTextW(szval);
	szNoidung+=szval;
	szNoidung+=L" ";

	for (int i = 2; i <= 6; i++)
	{
		m_txt[i].GetWindowTextW(szval);
		szNoidung+=szval;
		szNoidung+=L";";
	}

	m_txt_gc.GetWindowTextW(szval);
	szNoidung+=szval;

	pRange->PutItem(iRow,iCol,(_bstr_t)szNoidung);
	RangePtr PRS = pRange->GetItem(iRow,iCol);	
	if(PRS->GetComment()!=NULL) PRS->ClearComments();
	PRS->Interior->PutColor(xlNone);
	PRS->Font->PutColor(RGB(0,0,0));
	PRS->Font->PutItalic(1);
	PRS->Font->PutBold(0);
	PRS->EntireRow->Rows->AutoFit();
	if((int)PRS->GetRowHeight()>80) PRS->PutWrapText(FALSE);

	int icheckLM=0;
	for (int i = iRow-1; i >=1; i--)
	{
		szval = GIOR(i,icolLM,pRange,L"");
		if(szval.Left(2)==L"LM" || szval.Left(2)==L"NT")
		{
			PRS = pRange->GetItem(i,iCol);
			if(PRS->GetComment()==NULL)
			{
				icheckLM = 1;
				break;
			}
		}

		szval = GIOR(i,icotMAVL,pRange,L"");
		if(szval!=L"") break;
	}

	if (icheckLM == 0)
	{
		if ((int)m_chk_nthu.GetCheck() == 1)
		{
			pRange->PutItem(iRow, icolLM, (_bstr_t)L"NT");
		}
		else
		{
			pRange->PutItem(iRow, icolLM, (_bstr_t)L"LM");
		}		
	}

	CDialog::OnOK();
}


void frm_QuycachOld::OnBnClickedBtnCancel()
{
	CDialog::OnCancel();
}


void frm_QuycachOld::OnBnClickedBtnMaukhac()
{
	CDialog::OnOK();

	AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
	LaymauNgang _dlg;
	_dlg.DoModal();
}

void frm_QuycachOld::XacdinhQuycach()
{
	try
	{
		pSheet= pWb->ActiveSheet;
		pRange= pSheet->Cells;

		iRow = pSheet->Application->ActiveCell->Row;
		iCol = pSheet->Application->ActiveCell->Column;
		CString szNoidung = GIOR(iRow,iCol,pRange,L"");
		
		CString szval=L"";
		icolSTT = getColumn(pSheet,L"DMVL_STT");
		icotMAVL = getColumn(pSheet,L"DMVL_MAVL");
		icolLM = getColumn(pSheet,L"DMVL_MHDG");
		icolNdung = getColumn(pSheet,L"DMVL_ND");

		if(szNoidung==L"")
		{
			RangePtr PRS;
			for (int i = iRow-1; i >=1; i--)
			{
				szval = GIOR(i,icolLM,pRange,L"");
				if(szval.Left(2)==L"LM")
				{
					PRS = pRange->GetItem(i,iCol);
					if(PRS->GetComment()==NULL)
					{
						szNoidung = GIOR(i,iCol,pRange,L"");
						break;
					}
				}

				szval = GIOR(i,icotMAVL,pRange,L"");
				if(szval!=L"") break;
			}
		}

		for (int i = 1; i <= 7; i++) pKey[i]=L"";
		_TackTukhoaChuan(szNoidung,L";",L";",0,0);
		if(iKey>0)
		{
			m_cbb_ten.SetWindowText(pKey[1]);		
			m_txt_gc.SetWindowText(pKey[7]);

			int pos = pKey[2].Find(L" ");
			if(pos>0)
			{
				m_txt[1].SetWindowText(pKey[2].Left(pos));
				m_txt[2].SetWindowText(pKey[2].Right(pKey[2].GetLength()-pos-1));
			}

			m_txt[3].SetWindowText(pKey[3]);
			m_txt[4].SetWindowText(pKey[4]);
			m_txt[6].SetWindowText(pKey[6]);

			CString szKLNhap = L"";
			m_txt[1].GetWindowTextW(szKLNhap);
			
			szTansuat = pKey[5];
			m_txt[5].SetWindowText(AutoTinhSLMau(szKLNhap, szTansuat));
		}		

		int ivtri=0;		
		for (int i = iRow-1; i >= 1; i--)
		{
			szval = GIOR(i,icolSTT,pRange,L"");
			if(_wtoi(szval)>0)
			{
				ivtri = i;
				break;
			}
		}

		if(ivtri==0) return;

		szval = GIOR(ivtri,icotMAVL,pRange,L"");
		_TackTukhoaChuan(szval,L";",L"|",0,0);
		
		sslvl = iKey;
		if(sslvl==0) return;

		for (int i = 1; i <= iKey; i++)
		{
			MSVlieu[i].szMa = pKey[i];
			MSVlieu[i].szTen = L"";
		}

		szval = GIOR(ivtri,icolNdung,pRange,L"");
		_TackTukhoaChuan(szval,L";",L"|",0,0);		
		for (int i = 1; i <= iKey; i++) MSVlieu[i].szTen = pKey[i];

		int dem=0;
		for (int i = 1; i <= sslvl; i++)
		{
			if(MSVlieu[i].szTen!=L"")
			{
				dem++;
				m_cbb_ten.AddString(MSVlieu[i].szTen);
			}
		}

		if(dem>0 && szNoidung==L"")
		{
			m_cbb_ten.SetCurSel(0);
			OnCbnSelchangeCombo1();
		}
	}
	catch(...){}
}

void frm_QuycachOld::Loaddulieu(CString szMahieu)
{
	try
	{
		sslpp=0;
		szKichthuoc=L"";
		m_txt[6].SetCueBanner(L"");
		m_txt[2].SetWindowText(L"");
		m_txt[5].SetWindowText(L"");
		m_txt_hd.SetWindowText(L"");
		m_cbb_qcach.ResetContent();
		ASSERT(m_cbb_qcach.GetCount()==0);

		if(getPathQLCL()==L"") return;
		if(connectDsn(pathMDB)==-1) return;
		ObjConn.OpenAccessDB(pathMDB, L"", L"");

		CString szMavl=L"";
		SqlString.Format(L"SELECT * FROM tbMaVL_tenVL WHERE MaVL LIKE '%s';",szMahieu);
		CRecordset* Recset = ObjConn.Execute(SqlString);
		while( !Recset->IsEOF())
		{
			Recset->GetFieldValue(L"MaVL",szMavl);
			Recset->GetFieldValue(L"Kichthuocmau",szKichthuoc);
			m_txt[6].SetCueBanner(szKichthuoc);
			break;
		}		
		Recset->Close();

		if(szMavl!=L"")
		{
			int dem=0;
			CString szPPLM[100];
			SqlString.Format(L"SELECT * FROM tbMaVL_Quydinhlaymau WHERE MaVL LIKE '%s';",szMavl);
			Recset = ObjConn.Execute(SqlString);
			while( !Recset->IsEOF() )
			{
				dem++;
				Recset->GetFieldValue(L"MaLayMau",szPPLM[dem]);				
				Recset->MoveNext();
			}
			Recset->Close();

			if(dem>0)
			{
				for (int i = 1; i <= dem; i++)
				{
					SqlString.Format(L"SELECT * FROM tbMau WHERE MaLayMau LIKE '%s';",szPPLM[i]);
					Recset = ObjConn.Execute(SqlString);
					while( !Recset->IsEOF() )
					{
						sslpp++;						
						Recset->GetFieldValue(L"MaLayMau",MSPP[sslpp].szMa);
						Recset->GetFieldValue(L"Phuongphaplaymau",MSPP[sslpp].szTenPP);
						Recset->GetFieldValue(L"TansuatLaymau",MSPP[sslpp].szTSuat);
						Recset->GetFieldValue(L"Donvi",MSPP[sslpp].szDVT);
						m_cbb_qcach.AddString(MSPP[sslpp].szMa);
						break;
					}
					Recset->Close();
				}
			}
		}

		ObjConn.CloseDatabase();

		if(sslpp>0)
		{
			m_cbb_qcach.EnableWindow(1);
			m_cbb_qcach.SetCurSel(0);
			OnCbnSelchangeCbb2();
		}
		else m_cbb_qcach.EnableWindow(0);
	}
	catch(...){}
}

void frm_QuycachOld::OnCbnSelchangeCbb2()
{
	int nIndex = m_cbb_qcach.GetCurSel();
	if(nIndex>=0)
	{
		m_txt_hd.SetWindowText(MSPP[nIndex+1].szTenPP);
		m_txt[2].SetWindowText(MSPP[nIndex+1].szDVT);
		
		CString szKLNhap = L"";
		m_txt[1].GetWindowTextW(szKLNhap);

		szTansuat = MSPP[nIndex + 1].szTSuat;
		m_txt[5].SetWindowText(AutoTinhSLMau(szKLNhap, szTansuat));
	}
}

void frm_QuycachOld::OnEnSetfocusTxt6()
{
	if(szKichthuoc!=L"") m_txt[6].ShowBalloonTip(L"Gợi ý",szKichthuoc,TTI_INFO);
}

void frm_QuycachOld::OnEnChangeTxt1()
{
	CString szKLNhap = L"";
	m_txt[1].GetWindowTextW(szKLNhap);
	m_txt[5].SetWindowText(AutoTinhSLMau(szKLNhap, szTansuat));
}

CString frm_QuycachOld::AutoTinhSLMau(CString szKLNhap, CString szTansuat)
{
	try
	{
		szKLNhap.Trim(), szTansuat.Trim();
		szKLNhap.Replace(L",", L".");
		szTansuat.Replace(L",", L".");		

		int iSLM = 0;
		int iKL = _wtoi(szKLNhap);
		int iTS = _wtoi(szTansuat);
		if (iKL <= 0 || iTS <= 0) return L"";

		iSLM = (int)(iKL/iTS);
		if (szKLNhap.Find(L".") > 0 || iKL % iTS > 0) iSLM++;

		CString szSL = L"";
		szSL.Format(L"%d", iSLM);
		return szSL;
	}
	catch (...) {}
	return L"";
}
