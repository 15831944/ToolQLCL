#include "stdafx.h"
#include "Dlg_AutoSTT_HSNT.h"

// Dlg_AutoSTT_HSNT dialog
IMPLEMENT_DYNAMIC(Dlg_AutoSTT_HSNT, CDialogEx)

Dlg_AutoSTT_HSNT::Dlg_AutoSTT_HSNT(CWnd* pParent /*=NULL*/)
	: CDialogEx(Dlg_AutoSTT_HSNT::IDD, pParent)
{

}

Dlg_AutoSTT_HSNT::~Dlg_AutoSTT_HSNT()
{
}

void Dlg_AutoSTT_HSNT::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, HSNT_RD1, m_rdSel);
	DDX_Control(pDX, HSNT_RD2, m_rdHM);
	DDX_Control(pDX, HSNT_RD3, m_rdGD);
	DDX_Control(pDX, HSNT_CB1, m_cbbHM);
	DDX_Control(pDX, HSNT_T1, m_cbbTen);
	DDX_Control(pDX, HSNT_T2, m_txtSTT);
	DDX_Control(pDX, HSNT_SP1, m_spinSTT);
	DDX_Control(pDX, HSNT_CK1, m_ckTT);
	DDX_Control(pDX, HSNT_CK2, m_ckHT);

	DDX_Control(pDX, HSNT_T3, m_txtBD);
	DDX_Control(pDX, HSNT_T4, m_txtKT);
	DDX_Control(pDX, HSNT_SP2, m_spinBD);
	DDX_Control(pDX, HSNT_SP3, m_spinKT);
}


BEGIN_MESSAGE_MAP(Dlg_AutoSTT_HSNT, CDialogEx)
	ON_BN_CLICKED(HSNT_B1, &Dlg_AutoSTT_HSNT::OnBnClickedB1)
	ON_BN_CLICKED(HSNT_B2, &Dlg_AutoSTT_HSNT::OnBnClickedB2)
	ON_BN_CLICKED(HSNT_RD2, &Dlg_AutoSTT_HSNT::OnBnClickedRd2)
	ON_BN_CLICKED(HSNT_RD3, &Dlg_AutoSTT_HSNT::OnBnClickedRd3)
	ON_BN_CLICKED(HSNT_RD1, &Dlg_AutoSTT_HSNT::OnBnClickedRd1)
	ON_BN_CLICKED(HSNT_CK1, &Dlg_AutoSTT_HSNT::OnBnClickedCk1)
	ON_BN_CLICKED(HSNT_CK2, &Dlg_AutoSTT_HSNT::OnBnClickedCk2)
END_MESSAGE_MAP()


// Dlg_AutoSTT_HSNT message handlers
BOOL Dlg_AutoSTT_HSNT::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	iUpdate=0;
	m_rdSel.SetCheck(1);
	m_ckTT.SetCheck(1);
	m_cbbHM.EnableWindow(0);
	
	_Xacdinhvitri();

	m_cbbHM.SetCueBanner(L"Nhập tên...");
	m_cbbTen.AdjustDroppedWidth();
	m_cbbTen.SetMode(CComboBoxExt::MODE_AUTOCOMPLETE);

	_LoadMaHSNT(0);

	m_spinSTT.SetDecimalPlaces (0);
	m_spinSTT.SetTrimTrailingZeros (FALSE);
	m_spinSTT.SetRangeAndDelta (1, 9999, 1);
	m_spinSTT.SetPos(1);

	m_spinBD.SetDecimalPlaces (0);
	m_spinBD.SetTrimTrailingZeros (FALSE);
	m_spinBD.SetRangeAndDelta (8, ixdEnd-1, 10);
	m_spinBD.SetPos(8);

	m_spinKT.SetDecimalPlaces (0);
	m_spinKT.SetTrimTrailingZeros (FALSE);
	m_spinKT.SetRangeAndDelta (8, ixdEnd-1, 10);
	m_spinKT.SetPos(ixdEnd-1);

	return TRUE; 
}


void Dlg_AutoSTT_HSNT::_Xacdinhvitri()
{
	try
	{
		pSheet = pWb->ActiveSheet;
		pRange = pSheet->Cells;

		_bstr_t sh= pSheet->CodeName;
		if (sh==(_bstr_t)L"shHSNTCV")
		{
			iCotID = getColumn(pSheet,L"DMBB_STT");
			iCotMaCV = getColumn(pSheet,L"DMBB_MACV");
			iCotMaHS = getColumn(pSheet,L"DMBB_MAHS");
			iCotND = getColumn(pSheet,L"DMBB_ND");
		}
		else
		{
			iCotID = getColumn(pSheet,L"DMVL_STT");
			iCotMaCV = getColumn(pSheet,L"DMVL_MAVL");
			iCotMaHS = getColumn(pSheet,L"DMVL_MAHS");
			iCotND = getColumn(pSheet,L"DMVL_ND");
		}

		
		ixdEnd = FindCEnd(pSheet,iCotID,L"END",8);	
		if(ixdEnd<8) ixdEnd=8;

		CString szval=L"";

		iSLTT=0,iSLHT=0;
		_CtrTiento[0] = L"LEOGXD";
		_CtrHauto[0] = L"LEOGXD";

		slHM=0,slGD=0;
		for (int i = 8; i < ixdEnd; i++)
		{
			szval = GIOR(i,iCotMaCV,pRange,L"");
			szval.MakeLower();
			if(szval==L"hm")
			{
				slHM++;
				sTenHM[slHM] = GIOR(i,iCotND,pRange,L"");
				iVtrHM[slHM] = i;
			}
			else if(szval==L"gd" || szval==L"gđ" || szval==L"gĐ")
			{
				slGD++;
				sTenGD[slGD] = GIOR(i,iCotND,pRange,L"");
				iVtrGD[slGD] = i;
			}

			szval = GIOR(i,iCotID,pRange,L"");
			if(_wtoi(szval)>0)
			{
				szval = GIOR(i,iCotMaHS,pRange,L"");
				szval.Trim();
				if(szval!=L"") _CheckMaHSNT(szval);
			}			
		}		

		iVtrHM[slHM+1] = ixdEnd;
		iVtrGD[slGD+1] = ixdEnd;
	}
	catch(...){}
}


void Dlg_AutoSTT_HSNT::_CheckMaHSNT(CString sTen)
{
	try
	{
		int ikt = -1;
		CString szval=L"";		
		for (int i = sTen.GetLength()-1; i >=0 ; i--)
		{
			szval = sTen.Mid(i,1);
			if(_wtoi(szval)==0 && szval!=L"0")
			{
				ikt=i;
				break;
			}
		}

		szval=sTen;
		if(ikt>=0) szval = sTen.Left(ikt+1);

		ikt=0;
		for (int i = 0; i <= iSLTT; i++)
		{
			if(szval==_CtrTiento[i])
			{
				ikt=1;
				break;
			}
		}

		if(ikt==0)
		{
			iSLTT++;
			_CtrTiento[iSLTT] = szval;
		}

///////////////////////////////////////////////////////////////////////////////

		ikt = -1;
		szval=L"";		
		for (int i = 0; i < sTen.GetLength() ; i++)
		{
			szval = sTen.Mid(i,1);
			if(_wtoi(szval)==0 && szval!=L"0")
			{
				ikt=i;
				break;
			}
		}

		szval=sTen;
		if(ikt>=0) szval = sTen.Right(sTen.GetLength()-ikt);

		ikt=0;
		for (int i = 0; i <= iSLHT; i++)
		{
			if(szval==_CtrHauto[i])
			{
				ikt=1;
				break;
			}
		}

		if(ikt==0)
		{
			iSLHT++;
			_CtrHauto[iSLHT] = szval;
		}
	}
	catch(...){}
}

void Dlg_AutoSTT_HSNT::OnBnClickedB1()
{
	try
	{
		CString sTen=L"";
		m_cbbTen.GetWindowTextW(sTen);
		sTen.Trim();

		if(sTen==L"")
		{
			_s(L"Bạn chưa nhập tên mã HSNT.");
			return;	
		}

		iVBatdau=8;
		iVKetthuc=ixdEnd;
		if(iVKetthuc<=0) iVKetthuc=8;

		int numSt = 0;
		CString szval=L"";
		m_txtSTT.GetWindowTextW(szval);
		numSt = _wtoi(szval);
		if(numSt<1) numSt=1;
		else if(numSt>9999) numSt=9999;

		int nck = m_rdSel.GetCheck();
		if(nck==1)
		{
			CString sbd=L"",skt=L"";
			m_txtBD.GetWindowTextW(sbd);
			m_txtKT.GetWindowTextW(skt);
			sbd.Trim();
			skt.Trim();

			int ibd = _wtoi(sbd);
			int ikt = _wtoi(skt);

			if(ibd<=0) ibd=8;
			if(ikt<=0 || ikt>=ixdEnd) ikt=ixdEnd-1;

			iVBatdau = ibd;
			iVKetthuc = ikt;
		}
		else
		{
			nck = m_rdHM.GetCheck();
			if(nck==1)
			{
				nck = m_cbbHM.GetCurSel();
				if(nck>0)
				{
					iVBatdau = iVtrHM[nck];
					iVKetthuc = iVtrHM[nck+1];
				}
			}
			else
			{
				nck = m_cbbHM.GetCurSel();
				if(nck>0)
				{
					iVBatdau = iVtrGD[nck];
					iVKetthuc = iVtrGD[nck+1];
				}
			}
		}

		int cong=0;
		xl->PutStatusBar((_bstr_t)L"Đang tiến hành đánh lại mã HSNT. Vui lòng đợi trong giây lát...");

		for (int i = iVBatdau; i < iVKetthuc; i++)
		{
			szval = GIOR(i,iCotID,pRange,L"");
			if(_wtoi(szval)>0)
			{
				if(m_ckTT.GetCheck()==1) szval.Format(L"%s%d",sTen,numSt+cong);
				else szval.Format(L"%d%s",numSt+cong,sTen);

				pRange->PutItem(i,iCotMaHS,(_bstr_t)szval);
				cong++;
			}
		}

		xl->PutStatusBar((_bstr_t)L"Ready");
		CDialog::OnOK();
	}
	catch(...){}	
}


void Dlg_AutoSTT_HSNT::OnBnClickedB2()
{
	CDialog::OnCancel();
}


void Dlg_AutoSTT_HSNT::OnSysCommand(UINT nID, LPARAM lParam)
{
	if(nID == SC_CLOSE) OnBnClickedB2();
	else CDialog::OnSysCommand(nID, lParam);
}


BOOL Dlg_AutoSTT_HSNT::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		OnBnClickedB2();		
		return TRUE;
	}

	return FALSE;
}


void Dlg_AutoSTT_HSNT::_XoaCBB()
{
	m_cbbHM.ResetContent();
	ASSERT(m_cbbHM.GetCount() == 0);
}


void Dlg_AutoSTT_HSNT::OnBnClickedRd2()
{
	_XoaCBB();
	m_cbbHM.EnableWindow(1);
	_SetEnableBDKT(0);
	m_cbbHM.AddString(L"Tất cả hạng mục");
	if(slHM>0)
		for (int i = 1; i <= slHM; i++) m_cbbHM.AddString(sTenHM[i]);

	m_cbbHM.SetCurSel(0);
}


void Dlg_AutoSTT_HSNT::OnBnClickedRd3()
{
	_XoaCBB();
	m_cbbHM.EnableWindow(1);
	_SetEnableBDKT(0);
	m_cbbHM.AddString(L"Tất cả giai đoạn");
	if(slGD>0)
		for (int i = 1; i <= slGD; i++) m_cbbHM.AddString(sTenGD[i]);

	m_cbbHM.SetCurSel(0);
}

void Dlg_AutoSTT_HSNT::_SetEnableBDKT(int itype)
{
	m_txtBD.EnableWindow(itype);
	m_txtKT.EnableWindow(itype);
	m_spinBD.EnableWindow(itype);
	m_spinKT.EnableWindow(itype);
}

void Dlg_AutoSTT_HSNT::OnBnClickedRd1()
{
	m_cbbHM.EnableWindow(0);
	_SetEnableBDKT(1);
}


void Dlg_AutoSTT_HSNT::OnBnClickedCk1()
{
	int ick = m_ckTT.GetCheck();
	if(ick==1)
	{
		m_ckHT.SetCheck(0);
		_LoadMaHSNT(0);
	}
	else m_ckHT.SetCheck(1);
}


void Dlg_AutoSTT_HSNT::OnBnClickedCk2()
{
	iUpdate=1;
	int ick = m_ckHT.GetCheck();
	if(ick==1)
	{
		m_ckTT.SetCheck(0);
		_LoadMaHSNT(1);
	}
	else m_ckTT.SetCheck(1);
}


void Dlg_AutoSTT_HSNT::_LoadMaHSNT(int itype)
{
	try
	{
		m_cbbTen.ResetContent();
		ASSERT(m_cbbTen.GetCount() == 0);

		if(itype==0) for (int i = 1; i <= iSLTT; i++) m_cbbTen.AddString(_CtrTiento[i]);
		else for (int i = 1; i <= iSLHT; i++) m_cbbTen.AddString(_CtrHauto[i]);
	}
	catch(...){}
}