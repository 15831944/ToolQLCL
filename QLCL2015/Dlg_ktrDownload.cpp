#include "stdafx.h"
#include "Dlg_ktrDownload.h"


// Dlg_ktrDownload dialog
IMPLEMENT_DYNAMIC(Dlg_ktrDownload, CDialogEx)

Dlg_ktrDownload::Dlg_ktrDownload(CWnd* pParent /*=NULL*/)
	: CDialogEx(Dlg_ktrDownload::IDD, pParent)
{

}

Dlg_ktrDownload::~Dlg_ktrDownload()
{
}

void Dlg_ktrDownload::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, CKDW_T1, m_stt1);
	DDX_Control(pDX, CKDW_T2, m_txt2);
}


BEGIN_MESSAGE_MAP(Dlg_ktrDownload, CDialogEx)
	ON_BN_CLICKED(IDC_BUTTON1, &Dlg_ktrDownload::OnBnClickedButton1)
END_MESSAGE_MAP()


// Dlg_ktrDownload message handlers
BOOL Dlg_ktrDownload::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	CFont m_Font;
	m_Font.CreateFont( 18, 0, 0, 0, FW_BOLD, false, false,
		0, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY,
		FIXED_PITCH|FF_MODERN, _T("MS Shell Dlg") );
	m_stt1.SetFont(&m_Font,TRUE);

	iDem=0,iSai=0;
	_SetRandomQuit();

	return TRUE; 
}

// Bắt sự kiện click Enter
BOOL Dlg_ktrDownload::PreTranslateMessage(MSG* pMsg)
{
	 CWnd* pWndWithFocus = GetFocus();
	if (pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_TAB && pWndWithFocus == GetDlgItem(CKDW_T1))
	{
		GotoDlgCtrl(GetDlgItem(CKDW_T2));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_TAB && pWndWithFocus == GetDlgItem(CKDW_T2))
	{
		GotoDlgCtrl(GetDlgItem(CKDW_T1));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN && pWndWithFocus == GetDlgItem(CKDW_T2))
	{
		_CheckKetqua();
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		CDialog::OnCancel();
		return TRUE; 
	}

	return FALSE;
}


void Dlg_ktrDownload::_SetRandomQuit()
{
	try
	{
		iDem++;
		if(iDem>=6) return;
		iSai=0;

		CString szval=L"1 + 1 = ?";	

		time_t now = time(0);
		tm *ltm = localtime(&now);
		int iMin = ltm->tm_min;
		int iSec = ltm->tm_sec;
		int rd = (iMin+iSec+rand())%2;
		if(rd==1)
		{
			szval.Format(L"%s%s%s",MySymABC[rand()%36],MySymABC[rand()%36],MySymABC[rand()%36]);
			szval.MakeUpper();
		}
		else
		{
			rd = (iMin+iSec+rand())%4;
			int so1 = rand()%10;
			int so2 = rand()%10;

			if(rd==0) szval.Format(L"%d + %d = ?",so1,so2);
			else if(rd==1)
			{
				so1+=so2;
				szval.Format(L"%d - %d = ?",so1,so2);
			}
			else if(rd==2) szval.Format(L"%d x %d = ?",so1,so2);
			else if(rd==3)
			{
				so1 = so2*(rand()%5);
				szval.Format(L"%d / %d = ?",so1,so2);
			}
		}

		m_stt1.SetWindowText(szval);
	}
	catch(...){}
}


void Dlg_ktrDownload::_CheckKetqua()
{
	try
	{
		if(iSai>=6)
		{
			CDialog::OnOK();
			return;
		}

		sDapan=L"LEOGXD";
		CString szval=L"";
		m_stt1.GetWindowTextW(szval);		
		if(szval.Right(1)!=L"?") sDapan= szval;
		else
		{
			szval.Replace(L"?",L"");
			szval.Replace(L"=",L"");
			szval.Replace(L" ",L"");

			int itype=0;
			int pos = szval.Find(L"+");
			if(pos==-1)
			{
				itype=1;
				pos = szval.Find(L"-");
			}

			if(pos==-1)
			{
				itype=2;
				pos = szval.Find(L"x");
			}

			if(pos==-1)
			{
				itype=3;
				pos = szval.Find(L"/");
			}

			if(pos>0)
			{
				int so1 = _wtoi(szval.Left(pos));
				int so2 = _wtoi(szval.Right(szval.GetLength()-pos-1));

				if(itype==0) sDapan.Format(L"%d",so1+so2);
				else if(itype==1) sDapan.Format(L"%d",so1-so2);
				else if(itype==2) sDapan.Format(L"%d",so1*so2);
				else if(itype==3) sDapan.Format(L"%d",(int)(so1/so2));
			}				
		}

		m_txt2.GetWindowTextW(szval);
		if(szval==sDapan)
		{
			iCheckDownTC=0;
			CDialog::OnOK();
		}
		else iSai++;
	}
	catch(...){}
}

void Dlg_ktrDownload::OnBnClickedButton1()
{
	_SetRandomQuit();
	GotoDlgCtrl(GetDlgItem(CKDW_T2));
}
