#include "stdafx.h"
#include "Dlg_listNKyMDB.h"

// Dlg_listNKyMDB dialog
IMPLEMENT_DYNAMIC(Dlg_listNKyMDB, CDialogEx)

Dlg_listNKyMDB::Dlg_listNKyMDB(CWnd* pParent /*=NULL*/)
	: CDialogEx(Dlg_listNKyMDB::IDD, pParent)
{

}

Dlg_listNKyMDB::~Dlg_listNKyMDB()
{
}

void Dlg_listNKyMDB::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX,NKYMDB_L1,m_list);
}


BEGIN_MESSAGE_MAP(Dlg_listNKyMDB, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_BN_CLICKED(NKYMDB_B1, &Dlg_listNKyMDB::OnBnClickedB1)
	ON_BN_CLICKED(NKYMDB_B2, &Dlg_listNKyMDB::OnBnClickedB2)
	ON_NOTIFY(NM_DBLCLK, NKYMDB_L1, &Dlg_listNKyMDB::OnNMDblclkL1)
END_MESSAGE_MAP()


// Dlg_listNKyMDB message handlers
BOOL Dlg_listNKyMDB::OnInitDialog()
{
	CDialog::OnInitDialog();

	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	m_list.InsertColumn(0,L"",LVCFMT_CENTER,0);
	m_list.InsertColumn(1,L"Tên tệp tin",LVCFMT_LEFT,400);
	m_list.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);

	CString szval=L"";
	szval.Format(L"%s%s",_fGPathFolder(L"ToolQLCL.dll"),_pathDlieu);
	szval.Replace(L"////",L"");
	f_Add_Csdl(szval,L".mdb");
	if(_finder.GetFileCount()>0)
	{
		int dem=0;
		for(int i = 0; i< _finder.GetFileCount(); i++)
		{
			szval = _finder.GetFilePath(i).GetFileName();
			if(szval.Left(6)==L"Nhatky")
			{
				m_list.InsertItem(dem,_finder.GetFilePath(i).GetLongPath(),0);
				m_list.SetItemText(dem,1,szval);
				dem++;
			}
		}
	}

	return TRUE;
}


BOOL Dlg_listNKyMDB::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if (pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(NKYMDB_L1))
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


void Dlg_listNKyMDB::OnSysCommand(UINT nID, LPARAM lParam)
{
	if(nID == SC_CLOSE) CDialog::OnCancel();
	else CDialog::OnSysCommand(nID, lParam);
}


void Dlg_listNKyMDB::OnBnClickedB1()
{
	CDialog::OnOK();
	POSITION pos = m_list.GetFirstSelectedItemPosition();
	if(pos!=NULL)
	{
		for (int i = 0; i < m_list.GetItemCount();i++ )
			if (m_list.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
			{
				sTukhoa = m_list.GetItemText(i,0);
				return;
			}
	}
}


void Dlg_listNKyMDB::OnBnClickedB2()
{
	CDialog::OnCancel();
}


void Dlg_listNKyMDB::f_Add_Csdl(CString m_sBaseFolder, CString m_sFileMask)
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


void Dlg_listNKyMDB::OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	OnBnClickedB1();
}
