#include "Dlg_Cpys.h"

Dlg_Cpys::Dlg_Cpys() : CDialog(Dlg_Cpys::IDD)
{
}

void Dlg_Cpys::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, CPYS_L1, m_list55);
	DDX_Control(pDX, CPYS_CK1, m_ck1);
}

BEGIN_MESSAGE_MAP(Dlg_Cpys, CDialog)

	ON_BN_CLICKED(CPYS_B1, &Dlg_Cpys::OnBnClickedB1)
	ON_BN_CLICKED(CPYS_B2, &Dlg_Cpys::OnBnClickedB2)
	ON_BN_CLICKED(CPYS_CK1, &Dlg_Cpys::OnBnClickedCk1)
END_MESSAGE_MAP()

BOOL Dlg_Cpys::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);
	
	f_Style_list();
	f_Load_dsanh();

	return TRUE; 
}

// Bắt sự kiện click Enter
BOOL Dlg_Cpys::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		CDialog::OnCancel();
		return TRUE; 
	}

	return FALSE;
}


void Dlg_Cpys::f_Style_list()
{
	m_list55.InsertColumn(0,L"",LVCFMT_CENTER,22);
	m_list55.InsertColumn(1,L"",LVCFMT_LEFT,0);		// Code name
	m_list55.InsertColumn(2,L"Tên bảng tính",LVCFMT_LEFT,400);
	m_list55.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_CHECKBOXES);
}


void Dlg_Cpys::f_Load_dsanh()
{
	try
	{
		int dem=0;
		int ick=1;
		CString szval=L"";
		for (int i = 0; i < isxbs; i++)
		{
			m_list55.InsertItem(i,L"",0);
			m_list55.SetItemText(i,1,SXI_BS_cn[i+1]);
			m_list55.SetItemText(i,2,SXI_BS_name[i+1]);

			ick=1;
			if(ixkoin>0)
			{
				for (int j = 1; j <= ixkoin; j++)
				{
					if(SXI_BS_cn[i+1]==SXI_KO_cn[j])
					{
						ick=0;
						break;
					}
				}
			}
			
			if(ick==1)
			{
				dem++;
				m_list55.SetCheck(i,1);
			}			
		}

		if(dem>0 && dem==isxbs) m_ck1.SetCheck(1);
	}
	catch(...){}
}


void Dlg_Cpys::OnBnClickedB1()
{
	// Lọc ra các biên bản không in
	ixkoin=0;
	for (int i = 0; i < m_list55.GetItemCount(); i++)
	{
		if((int)m_list55.GetCheck(i)==0)
		{
			ixkoin++;
			SXI_KO_cn[ixkoin] = m_list55.GetItemText(i,1);
		}
	}

	CDialog::OnOK();
}


void Dlg_Cpys::OnBnClickedB2()
{
	CDialog::OnCancel();
}


void Dlg_Cpys::OnBnClickedCk1()
{
	CButton *m_ctlCheck = (CButton*) GetDlgItem(CPYS_CK1);
	if(m_ctlCheck->GetCheck() == BST_UNCHECKED) f_SetCheck(0);
	else f_SetCheck(1);
}

void Dlg_Cpys::f_SetCheck(int n)
{
	int ncount = m_list55.GetItemCount();
	if(ncount>0)
		for (int i = 0; i < ncount; i++) m_list55.SetCheck(i,n);
}
