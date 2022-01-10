#include "stdafx.h"
#include "Dlg_ChangeLM3.h"


// Dlg_ChangeLM3 dialog

IMPLEMENT_DYNAMIC(Dlg_ChangeLM3, CDialog)

Dlg_ChangeLM3::Dlg_ChangeLM3(CWnd* pParent /*=NULL*/)
	: CDialog(Dlg_ChangeLM3::IDD, pParent)
{

}

Dlg_ChangeLM3::~Dlg_ChangeLM3()
{
}

void Dlg_ChangeLM3::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, CGLM3_T1, m_t1);
	DDX_Control(pDX, CGLM3_T2, m_t2);
}


BEGIN_MESSAGE_MAP(Dlg_ChangeLM3, CDialog)
	ON_BN_CLICKED(IDC_BUTTON7, &Dlg_ChangeLM3::OnBnClickedButton7)
	ON_BN_CLICKED(IDC_BUTTON8, &Dlg_ChangeLM3::OnBnClickedButton8)
	ON_BN_CLICKED(IDC_NEX1, &Dlg_ChangeLM3::OnBnClickedNex1)
	ON_BN_CLICKED(IDC_PRV1, &Dlg_ChangeLM3::OnBnClickedPrv1)
END_MESSAGE_MAP()


// Dlg_ChangeLM3 message handlers
BOOL Dlg_ChangeLM3::OnInitDialog()
{
	CDialog::OnInitDialog();

	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	if(iKqua==-1)
	{
		this->SetWindowTextW(L"Đổi tên mẫu");

		CButton* btn1 = (CButton*)GetDlgItem(IDC_PRV1);
		CButton* btn2 = (CButton*)GetDlgItem(IDC_NEX1);
		CStatic* stt3 = (CStatic*)GetDlgItem(IDC_HDAN);

		btn1->EnableWindow(0);
		btn2->EnableWindow(0);
		stt3->EnableWindow(0);
	}
	else xl_tukhoa = _NameColumnHeading(0,iKqua,L"");
	
	m_t1.SetWindowText(xl_tukhoa);
	m_t2.SetCueBanner(L"Nhập tên mới");

	return TRUE;
}


BOOL Dlg_ChangeLM3::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	
	if(pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(CGLM3_T2))
	{
		if(iKqua>=0) OnBnClickedButton7();
		else GotoDlgCtrl(GetDlgItem(CGLM3_T1));
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(CGLM3_T1))
	{
		GotoDlgCtrl(GetDlgItem(CGLM3_T2));
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_TAB &&
		pWndWithFocus == GetDlgItem(CGLM3_T2))
	{
		GotoDlgCtrl(GetDlgItem(CGLM3_T1));
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_TAB &&
		pWndWithFocus == GetDlgItem(CGLM3_T1))
	{
		GotoDlgCtrl(GetDlgItem(CGLM3_T2));
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(IDC_NEX1))
	{
		OnBnClickedNex1();
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		CDialog::OnCancel();
		return TRUE; 
	}

	return FALSE;
}


void Dlg_ChangeLM3::OnBnClickedButton7()
{
	if(iKqua>=0)
	{
		CString s=L"";
		m_t2.GetWindowTextW(s);
		s.Trim();
		if(s==L"") s=L"(null)";
		xl_tukhoa=s;

		_NameColumnHeading(1,iKqua,xl_tukhoa);
		GotoDlgCtrl(GetDlgItem(IDC_NEX1));
	}
	else
	{
		xl_tukhoa=L"";
		m_t2.GetWindowTextW(xl_tukhoa);
		xl_tukhoa.Trim();
		if(xl_tukhoa==L"")
		{
			_s(L"Vui lòng nhập tên mới của mẫu");
			GotoDlgCtrl(GetDlgItem(CGLM3_T2));
			return;
		}

		if (connectDsn(pathQCLM) == -1) return;
		ObjConn.OpenAccessDB(pathQCLM, L"", L"");
		SqlString.Format(L"UPDATE tbTenMau SET TenMau='%s' WHERE ID='%s';", xl_tukhoa, sIDmau);

		ObjConn.ExecuteRB(SqlString);
		ObjConn.CloseDatabase();
		CDialog::OnOK();
		iKqua=-2;
	}	
}


void Dlg_ChangeLM3::OnBnClickedButton8()
{
	CDialog::OnCancel();
}


void Dlg_ChangeLM3::OnBnClickedNex1()
{
	iKqua++;

	int sl = litLMN3.GetColumnCount();
	if(iKqua==sl) iKqua = 1;
	CString szval=L"";
	szval.Format(L"Thay đổi tên cột (%d)",iKqua);
	this->SetWindowTextW(szval);

	xl_tukhoa = _NameColumnHeading(0,iKqua,L"");
	m_t1.SetWindowText(xl_tukhoa);
	m_t2.SetWindowText(L"");
	GotoDlgCtrl(GetDlgItem(CGLM3_T2));
}


void Dlg_ChangeLM3::OnBnClickedPrv1()
{
	iKqua--;
	if(iKqua==0) iKqua = litLMN3.GetColumnCount()-1;
	CString szval=L"";
	szval.Format(L"Thay đổi tên cột (%d)",iKqua);
	this->SetWindowTextW(szval);

	xl_tukhoa = _NameColumnHeading(0,iKqua,L"");
	m_t1.SetWindowText(xl_tukhoa);
	m_t2.SetWindowText(L"");
	GotoDlgCtrl(GetDlgItem(CGLM3_T2));
}
