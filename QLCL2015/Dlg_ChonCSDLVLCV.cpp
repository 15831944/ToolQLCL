#include "stdafx.h"
#include "Dlg_ChonCSDLVLCV.h"

IMPLEMENT_DYNAMIC(Dlg_ChonCSDLVLCV, CDialogEx)

Dlg_ChonCSDLVLCV::Dlg_ChonCSDLVLCV(CWnd* pParent /*=NULL*/)
	: CDialogEx(Dlg_ChonCSDLVLCV::IDD, pParent)
{

}

Dlg_ChonCSDLVLCV::~Dlg_ChonCSDLVLCV()
{
}

void Dlg_ChonCSDLVLCV::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, DLCVVL_L1, m_list);
}


BEGIN_MESSAGE_MAP(Dlg_ChonCSDLVLCV, CDialogEx)
	ON_BN_CLICKED(DLCVVL_B2, &Dlg_ChonCSDLVLCV::OnBnClickedB2)
	ON_BN_CLICKED(DLCVVL_B1, &Dlg_ChonCSDLVLCV::OnBnClickedB1)
	ON_NOTIFY(NM_DBLCLK, DLCVVL_L1, &Dlg_ChonCSDLVLCV::OnNMDblclkL1)
	ON_NOTIFY(LVN_KEYDOWN, DLCVVL_L1, &Dlg_ChonCSDLVLCV::OnLvnKeydownL1)
END_MESSAGE_MAP()


// Dlg_ChonCSDLVLCV message handlers
// Dlg_NhomkyBB message handlers
BOOL Dlg_ChonCSDLVLCV::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	_LoadDS();
	return TRUE; 
}

// Bắt sự kiện click Enter
BOOL Dlg_ChonCSDLVLCV::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if (pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(DLCVVL_L1))
	{
		OnBnClickedB1();
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		CDialog::OnCancel();
		return TRUE; 
	}

	return FALSE;
}


void Dlg_ChonCSDLVLCV::_LoadDS()
{
	try
	{
		m_list.InsertColumn(0,L"",LVCFMT_CENTER,0);
		m_list.InsertColumn(1,L"Tệp tin",LVCFMT_LEFT,400);
		m_list.InsertColumn(2,L"Chọn",LVCFMT_CENTER,80);
		m_list.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);

		CString szval = L"";
		CString sfile = GIOR(1,7,prTS,L"");

		CString sFd = L"";
		sFd.Format(L"%s%s",_fGPathFolder(L"ToolQLCL.dll"),_pathFILE);

		f_Add_Csdl(sFd,L".mdb");
		int ncount = _finder.GetFileCount();		
		if(ncount>0)
		{
			for (int i = 0; i < ncount; i++)
			{
				szval = _finder.GetFilePath(i).GetLongPath();
				m_list.InsertItem(i,L"",0);
				m_list.SetItemText(i,1,szval);

				if(szval==sfile) m_list.SetItemText(i,2,L"X");
			}
		}
	}
	catch(...){}
}


void Dlg_ChonCSDLVLCV::f_Add_Csdl(CString m_sBaseFolder, CString m_sFileMask)
{
	try
	{
		CFileFinder::CFindOpts	opts;

		opts.sBaseFolder = m_sBaseFolder;
		opts.sFileMask.Format(L"*%s*", m_sFileMask);
		opts.FindNormalFiles();

		_finder.RemoveAll();
		_finder.Find(opts);
	}
	catch(...){}
}


void Dlg_ChonCSDLVLCV::OnBnClickedB2()
{
	CDialog::OnCancel();
}


void Dlg_ChonCSDLVLCV::OnBnClickedB1()
{
	POSITION pos=m_list.GetFirstSelectedItemPosition();
	if(pos!=NULL)
	{
		for (int i = 0; i < m_list.GetItemCount(); i++)
		if (m_list.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
		{
			CString szval = m_list.GetItemText(i,1);
			prTS->PutItem(1,7,(_bstr_t)szval);
			CDialog:OnOK();
			return;
		}
	}
}


void Dlg_ChonCSDLVLCV::OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	OnBnClickedB1();
	*pResult = 0;
}


void Dlg_ChonCSDLVLCV::OnLvnKeydownL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);
	if(pLVKeyDow->wVKey == VK_SPACE) OnBnClickedB1();

	*pResult = 0;
}
