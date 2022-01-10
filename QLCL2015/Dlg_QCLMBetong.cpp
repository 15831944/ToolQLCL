#include "stdafx.h"
#include "Dlg_QCLMBetong.h"

// Dlg_QCLMBetong dialog
IMPLEMENT_DYNAMIC(Dlg_QCLMBetong, CDialog)

Dlg_QCLMBetong::Dlg_QCLMBetong(CWnd* pParent /*=NULL*/)
	: CDialog(Dlg_QCLMBetong::IDD, pParent)
{

}

Dlg_QCLMBetong::~Dlg_QCLMBetong()
{
}

void Dlg_QCLMBetong::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, QCLM_S1, m_stt[1]);
	DDX_Control(pDX, QCLM_S2, m_stt[2]);
	DDX_Control(pDX, QCLM_S3, m_stt[3]);
	DDX_Control(pDX, QCLM_S4, m_stt[4]);
	DDX_Control(pDX, QCLM_S5, m_stt[5]);

	DDX_Control(pDX, IDC_ETN_NSX, m_txt[1]);
	DDX_Control(pDX, IDC_ETN_MAC, m_txt[2]);
	DDX_Control(pDX, IDC_ETN_SUT, m_txt[3]);
	DDX_Control(pDX, IDC_ETN_SL, m_txt[4]);
	DDX_Control(pDX, IDC_ETN_KT, m_txt[5]);
	
	DDX_Control(pDX, IDC_EDIT_XM, m_txt[6]);
	DDX_Control(pDX, IDC_EDIT_XM_DV, m_txt[7]);
	DDX_Control(pDX, IDC_EDIT_XM_KL, m_txt[8]);

	DDX_Control(pDX, IDC_EDIT_CAT, m_txt[9]);
	DDX_Control(pDX, IDC_EDIT_CAT_DV, m_txt[10]);
	DDX_Control(pDX, IDC_EDIT_CAT_KL, m_txt[11]);

	DDX_Control(pDX, IDC_EDIT_DA, m_txt[12]);
	DDX_Control(pDX, IDC_EDIT_DA_DV, m_txt[13]);
	DDX_Control(pDX, IDC_EDIT_DA_KL, m_txt[14]);

	DDX_Control(pDX, IDC_EDIT_NC, m_txt[15]);
	DDX_Control(pDX, IDC_EDIT_NC_DV, m_txt[16]);
	DDX_Control(pDX, IDC_EDIT_NC_KL, m_txt[17]);

	DDX_Control(pDX, IDC_EDIT_PG1, m_txt[18]);
	DDX_Control(pDX, IDC_EDIT_PG1_DV, m_txt[19]);
	DDX_Control(pDX, IDC_EDIT_PG1_KL, m_txt[20]);

	DDX_Control(pDX, IDC_EDIT_PG2, m_txt[21]);
	DDX_Control(pDX, IDC_EDIT_PG2_DV, m_txt[22]);
	DDX_Control(pDX, IDC_EDIT_PG2_KL, m_txt[23]);
}


BEGIN_MESSAGE_MAP(Dlg_QCLMBetong, CDialog)
	ON_BN_CLICKED(ID_MTN_OK, &Dlg_QCLMBetong::OnBnClickedMtnOk)
	ON_BN_CLICKED(ID_MTN_CANCEL, &Dlg_QCLMBetong::OnBnClickedMtnCancel)
	ON_BN_CLICKED(ID_MTN_OK2, &Dlg_QCLMBetong::OnBnClickedMtnOk2)
	ON_BN_CLICKED(ID_MTN_OK3, &Dlg_QCLMBetong::OnBnClickedMtnOk3)
	ON_BN_CLICKED(ID_MTN_OK4, &Dlg_QCLMBetong::OnBnClickedMtnOk4)
	ON_BN_CLICKED(ID_MTN_OK5, &Dlg_QCLMBetong::OnBnClickedMtnOk5)
	ON_BN_CLICKED(IDC_AUTO_CHECK, &Dlg_QCLMBetong::OnBnClickedAutoCheck)
END_MESSAGE_MAP()


// Dlg_QCLMBetong message handlers
BOOL Dlg_QCLMBetong::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	for (int i = 1; i <= 5; i++) m_stt[i].SetWindowText(pKLMBT[i]);

	// Tách chuỗi	
	if(schuoiMTN!=L"")
	{
		f_Tack_chuoi(schuoiMTN);
		if(iKey>0) for (int i = 1; i <= iKey; i++) m_txt[i].SetWindowText(pKey[i]);
	}

	return TRUE; 
}


// Bắt sự kiện click Enter
BOOL Dlg_QCLMBetong::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(IDC_ETN_NSX))
	{
		GotoDlgCtrl(GetDlgItem(IDC_ETN_MAC));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(IDC_ETN_MAC))
	{
		GotoDlgCtrl(GetDlgItem(IDC_ETN_SUT));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(IDC_ETN_SUT))
	{
		GotoDlgCtrl(GetDlgItem(IDC_ETN_SL));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(IDC_ETN_SL))
	{
		GotoDlgCtrl(GetDlgItem(IDC_ETN_KT));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(IDC_ETN_KT))
	{
		GotoDlgCtrl(GetDlgItem(IDC_EDIT_XM));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(IDC_EDIT_XM))
	{
		GotoDlgCtrl(GetDlgItem(IDC_EDIT_XM_DV));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(IDC_EDIT_XM_DV))
	{
		GotoDlgCtrl(GetDlgItem(IDC_EDIT_XM_KL));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(IDC_EDIT_XM_KL))
	{
		GotoDlgCtrl(GetDlgItem(IDC_EDIT_CAT));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(IDC_EDIT_CAT))
	{
		GotoDlgCtrl(GetDlgItem(IDC_EDIT_CAT_DV));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(IDC_EDIT_CAT_DV))
	{
		GotoDlgCtrl(GetDlgItem(IDC_EDIT_CAT_KL));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(IDC_EDIT_CAT_KL))
	{
		GotoDlgCtrl(GetDlgItem(IDC_EDIT_DA));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(IDC_EDIT_DA))
	{
		GotoDlgCtrl(GetDlgItem(IDC_EDIT_DA_DV));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(IDC_EDIT_DA_DV))
	{
		GotoDlgCtrl(GetDlgItem(IDC_EDIT_DA_KL));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(IDC_EDIT_DA_KL))
	{
		GotoDlgCtrl(GetDlgItem(IDC_EDIT_NC));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(IDC_EDIT_NC))
	{
		GotoDlgCtrl(GetDlgItem(IDC_EDIT_NC_DV));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(IDC_EDIT_NC_DV))
	{
		GotoDlgCtrl(GetDlgItem(IDC_EDIT_NC_KL));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(IDC_EDIT_NC_KL))
	{
		GotoDlgCtrl(GetDlgItem(IDC_EDIT_PG1));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(IDC_EDIT_PG1))
	{
		GotoDlgCtrl(GetDlgItem(IDC_EDIT_PG1_DV));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(IDC_EDIT_PG1_DV))
	{
		GotoDlgCtrl(GetDlgItem(IDC_EDIT_PG1_KL));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(IDC_EDIT_PG1_KL))
	{
		GotoDlgCtrl(GetDlgItem(IDC_EDIT_PG2));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(IDC_EDIT_PG2))
	{
		GotoDlgCtrl(GetDlgItem(IDC_EDIT_PG2_DV));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(IDC_EDIT_PG2_DV))
	{
		GotoDlgCtrl(GetDlgItem(IDC_EDIT_PG2_KL));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(IDC_EDIT_PG2_KL))
	{
		GotoDlgCtrl(GetDlgItem(IDC_ETN_NSX));
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		CDialog::OnCancel();
		return TRUE; 
	}

	return FALSE;
}


void Dlg_QCLMBetong::f_Tack_chuoi(CString sKey)
{
	try
	{
		if(sKey.Right(1)!=L";") sKey+=L";";
		iKey=0;
		int vtri=0;
		CString szval=L"";
		for (int i = 0; i < sKey.GetLength(); i++)
		{
			szval = sKey.Mid(i,1);
			if((szval==L";") && i>vtri)
			{
				iKey++;
				pKey[iKey] = sKey.Mid(vtri,i-vtri);
				pKey[iKey].Trim();
				vtri=i+1;

				if(iKey==23) break;
			}
		}
	}
	catch(...){}
}


void Dlg_QCLMBetong::OnBnClickedMtnOk()
{
	try
	{
		iKqua=1;
		schuoiMTN=L"";

		CString szval=L"";
		for (int i = 1; i <= 23; i++)
		{
			m_txt[i].GetWindowTextW(szval);
			szval.Trim();
			if(szval==L"") szval=L" ";
		
			schuoiMTN+=szval;
			if(i<23) schuoiMTN+=L";";
		}

		CDialog::OnOK();
	}
	catch(...){}
}


void Dlg_QCLMBetong::OnBnClickedMtnCancel()
{
	iKqua=0;
	CDialog::OnCancel();
}


void Dlg_QCLMBetong::OnBnClickedMtnOk2()
{
	CString szval=L"";
	m_txt[3].GetWindowTextW(szval);
	szval+=L"±";
	m_txt[3].SetWindowText(szval);
}


void Dlg_QCLMBetong::OnBnClickedMtnOk3()
{
	CString szval=L"";
	m_txt[3].GetWindowTextW(szval);
	szval+=L"÷";
	m_txt[3].SetWindowText(szval);
}


void Dlg_QCLMBetong::OnBnClickedMtnOk4()
{
	CDialog::OnOK();
	OQCLMAU();
}

void Dlg_QCLMBetong::OnBnClickedMtnOk5()
{ 
	for (int i = 1; i <= 23; i++) m_txt[i].SetWindowText(L"");
}

void Dlg_QCLMBetong::OnBnClickedAutoCheck()
{
	try
	{
		CButton *btn_check = (CButton*)GetDlgItem(IDC_AUTO_CHECK);
		if((int)btn_check->GetCheck()==1)
		{
			CString szAuto[24];
			for (int i = 6; i <= 23; i++) m_txt[6].GetWindowTextW(szAuto[i]);
	
			int icheck=0;
			for (int i = 2; i <= 7; i++)
			{
				if(szAuto[3*i]!=L"")
				{
					icheck=1;
					break;
				}
			}

			if(icheck==0)
			{
				m_txt[6].SetWindowText(L"Xi măng");
				m_txt[7].SetWindowText(L"kg");
			
				m_txt[9].SetWindowText(L"Cát");
				m_txt[10].SetWindowText(L"m3");
			
				m_txt[12].SetWindowText(L"Đá");
				m_txt[13].SetWindowText(L"m3");
			
				m_txt[15].SetWindowText(L"Nước");
				m_txt[16].SetWindowText(L"lít");

				m_txt[18].SetWindowText(L"Phụ gia.. (nếu có)");
				m_txt[19].SetWindowText(L"kg");
			}
		}
		else
		{
			for (int i = 6; i <= 23; i++) m_txt[i].SetWindowText(L"");
		}		
	}
	catch(...){}
}
