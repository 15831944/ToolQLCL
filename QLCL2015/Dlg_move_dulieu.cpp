// Dlg_move_dulieu.cpp : implementation file
#include "stdafx.h"
#include "Dlg_move_dulieu.h"

// Dlg_move_dulieu dialog
IMPLEMENT_DYNAMIC(Dlg_move_dulieu, CDialog)

Dlg_move_dulieu::Dlg_move_dulieu(CWnd* pParent /*=NULL*/)
	: CDialog(Dlg_move_dulieu::IDD, pParent)
{

}

Dlg_move_dulieu::~Dlg_move_dulieu()
{
}

void Dlg_move_dulieu::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, BGHD_CB1, m_cbb1);
	DDX_Control(pDX, BGHD_CB2, m_cbb2);
	DDX_Control(pDX, BGHD_CK1, m_chk1);
	DDX_Control(pDX, BGHD_CK2, m_chk2);
	DDX_Control(pDX, BGHD_CK3, m_chk3);
	DDX_Control(pDX, BGHD_CK4, m_chk4);
}


BEGIN_MESSAGE_MAP(Dlg_move_dulieu, CDialog)
	ON_BN_CLICKED(BGHD_B1, &Dlg_move_dulieu::OnBnClickedB1)
	ON_BN_CLICKED(BGHD_B2, &Dlg_move_dulieu::OnBnClickedB2)
	ON_CBN_SELCHANGE(BGHD_CB1, &Dlg_move_dulieu::OnCbnSelchangeCb1)
	ON_CBN_SELCHANGE(BGHD_CB2, &Dlg_move_dulieu::OnCbnSelchangeCb2)
END_MESSAGE_MAP()


// Dlg_move_dulieu message handlers
BOOL Dlg_move_dulieu::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	m_cbb1.SetCueBanner(L"Kích chọn hạng mục");
	m_cbb2.SetCueBanner(L"Kích chọn giai đoạn");

	m_chk1.SetCheck(1);
	m_chk2.SetCheck(1);
	m_chk3.SetCheck(1);
	m_chk4.SetCheck(1);

	f_Xac_dinh_hangmuc();

	return TRUE; 
}


// Bắt sự kiện click Enter
BOOL Dlg_move_dulieu::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		CDialog::OnCancel();
		return TRUE; 
	}

	return FALSE;
}


void Dlg_move_dulieu::f_Xac_dinh_hangmuc()
{
	try
	{
		_WorksheetPtr pSh1 = pWb->ActiveSheet;
		_bstr_t _bsh1 = pSh1->CodeName;
		if(_bsh1==(_bstr_t)L"shDTXD")
		{
			// Công việc
			shDMCV = name_sheet((_bstr_t)"shHSNTCV");
			psDMCV = xl->Sheets->GetItem(shDMCV);
			prDMCV = psDMCV->Cells;

			iCdm1 = getColumn(psDMCV,L"DMBB_STT");
			iCdm4 = getColumn(psDMCV,L"DMBB_MACV");
			iCdm7 = getColumn(psDMCV,L"DMBB_ND");
			iCdm12 = getColumn(psDMCV,L"DMBB_TC");
		}
		else
		{
			// Vật liệu
			shDMCV = name_sheet((_bstr_t)"shHSNTVL");
			psDMCV = xl->Sheets->GetItem(shDMCV);
			prDMCV = psDMCV->Cells;

			iCdm1 = getColumn(psDMCV,L"DMVL_STT");
			iCdm4 = getColumn(psDMCV,L"DMVL_MAVL");
			iCdm7 = getColumn(psDMCV,L"DMVL_ND");
			iCdm12 = getColumn(psDMCV,L"DMVL_TC");
		}
		
		iEnd = FindComment(psDMCV,iCdm1,"END");
		vtketiep = iEnd;
		vthm=6,vtgd=6;

		int vt1=-1;
		int vt2=-1;
		int dem=0;
		CString szval=L"";
		
		// Xác định vị trí chứa hạng mục active
		for (int i = 8; i < iEnd; i++)
		{
			szval = GIOR(i,iCdm4,prDMCV,L"");
			if(szval==L"HM")
			{
				szval = GIOR(i,iCdm7,prDMCV,L"NULL");
				m_cbb1.AddString(szval);
				if(szval==sbgHM)
				{
					vt1=dem;	// vị trí combobox
					vthm=i;		// vị trí hạng mục
				}

				dem++;
			}			
		}

		// vt1>=0 tương ứng có hạng mục active
		if(vt1>=0)
		{
			m_cbb1.SetCurSel(vt1);

			// Xác định GD tương ứng
			dem=0;
			for (int i = vthm+1; i < iEnd; i++)
			{
				szval = GIOR(i,iCdm4,prDMCV,L"");
				if(szval==L"GD")
				{
					// Kiểm tra giai đoạn active
					szval = GIOR(i,iCdm7,prDMCV,L"NULL");
					m_cbb2.AddString(szval);
					if(szval==sbgGD)
					{
						vt2=dem;	// vị trí combobox
						vtgd=i;		// vị trí giai đoạn
					}

					dem++;
				}
				else if(szval==L"HM")
				{
					vtketiep=i;		// nếu không tồn tại giai đoạn
					break;
				}
			}

			// tồn tại giai đoạn active
			if(vt2>=0)
			{
				m_cbb2.SetCurSel(vt2);

				// Xác định vị trí cuối
				for (int i = vtgd+1; i < iEnd; i++)
				{
					szval = GIOR(i,iCdm4,prDMCV,L"");
					if(szval==L"HM" || szval==L"GD")
					{
						vtketiep=i;		// tìm thấy vị trí kế tiếp
						break;
					}
				}
			}
		}
		else
		{
			// không tìm thấy vị trí active hạng mục --> có thể tồn tại hạng mục hoặc không
			dem=0;
			for (int i = 8; i < iEnd; i++)
			{
				szval = GIOR(i,iCdm4,prDMCV,L"");
				if(szval==L"GD")
				{
					szval = GIOR(i,iCdm7,prDMCV,L"NULL");
					m_cbb2.AddString(szval);
					if(szval==sbgGD)
					{
						vt2=dem;
						vtgd=i;
					}

					dem++;
				}
			}

			if(vt2>=0)
			{
				m_cbb2.SetCurSel(vt2);

				// Xác định vị trí cuối
				for (int i = vtgd+1; i < iEnd; i++)
				{
					szval = GIOR(i,iCdm4,prDMCV,L"");
					if(szval==L"HM" || szval==L"GD")
					{
						vtketiep=i;		// tìm thấy vị trí kế tiếp
						break;
					}
				}
			}
		}		
	}
	catch(...){_s(L"Lỗi xác định hạng mục!");}
}


void Dlg_move_dulieu::OnBnClickedB2()
{
	CDialog::OnCancel();
}


void Dlg_move_dulieu::OnCbnSelchangeCb1()
{
	try
	{
		vtketiep=iEnd;

		m_cbb2.ResetContent();
		ASSERT(m_cbb2.GetCount() == 0);

		// Xác định vị trí chứa hạng mục tương ứng
		int nvtri = m_cbb1.GetCurSel();
		m_cbb1.GetLBText(nvtri,sbgHM);

		vthm=6;
		int dem=0;
		CString szval=L"";
		for (int i = 8; i < iEnd; i++)
		{
			szval = GIOR(i,iCdm4,prDMCV,L"");
			if(szval==L"HM")
			{
				dem++;
				if(nvtri+1==dem)
				{
					vthm=i;
					break;
				}
			}
		}

		if(vthm==0) return;

		for (int i = vthm+1; i < iEnd; i++)
		{
			szval = GIOR(i,iCdm4,prDMCV,L"");
			if(szval==L"GD")
			{
				szval = GIOR(i,iCdm7,prDMCV,L"NULL");
				m_cbb2.AddString(szval);
			}
			else if(szval==L"HM")
			{
				vtketiep=i;
				break;
			}
		}
	}
	catch(...){}	
}


void Dlg_move_dulieu::OnCbnSelchangeCb2()
{
	try
	{
		vtketiep=iEnd;
		int dem=0;
		int nvtri = m_cbb2.GetCurSel();
		m_cbb2.GetLBText(nvtri,sbgGD);
		if(vthm==0) vtgd=7;

		CString szval=L"";
		for (int i = vthm+1; i < iEnd; i++)
		{
			szval = GIOR(i,iCdm4,prDMCV,L"");
			if(szval==L"GD")
			{
				dem++;
				if(nvtri+1==dem)
				{
					vtgd=i;
					break;
				}
			}
		}

		for (int i = vtgd+1; i < iEnd; i++)
		{
			szval = GIOR(i,iCdm4,prDMCV,L"");
			if(szval==L"HM" || szval==L"GD")
			{
				vtketiep=i;
				break;
			}
		}
	}
	catch(...){}	
}


void Dlg_move_dulieu::OnBnClickedB1()
{
	// Kiểm tra check
	ibgTG = m_chk1.GetCheck();
	ibgTC = m_chk2.GetCheck();
	ibgTM = m_chk3.GetCheck();
	ibgDMNC = m_chk4.GetCheck();

	ivtdchuyen=9;
	CString s1=L"",s2=L"";
	for (int i = vtketiep-1; i > 7; i--)
	{
		s1 = GIOR(i,iCdm7,prDMCV,L"");
		s2 = GIOR(i,iCdm12,prDMCV,L"");
		if(s1!=L"" || s2!=L"")
		{
			ivtdchuyen=i+1;
			break;
		}
	}

	CDialog::OnOK();
}
