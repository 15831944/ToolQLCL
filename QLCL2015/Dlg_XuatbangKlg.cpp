#include "stdafx.h"
#include "Dlg_XuatbangKlg.h"

// Dlg_XuatbangKlg dialog
IMPLEMENT_DYNAMIC(Dlg_XuatbangKlg, CDialog)

Dlg_XuatbangKlg::Dlg_XuatbangKlg(CWnd* pParent /*=NULL*/)
	: CDialog(Dlg_XuatbangKlg::IDD, pParent)
{

}

Dlg_XuatbangKlg::~Dlg_XuatbangKlg()
{
}

void Dlg_XuatbangKlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, XKL_CB1, m_cbb1);
	DDX_Control(pDX, XKL_CB2, m_cbb2);
	DDX_Control(pDX, XKL_CK1, m_chk);
}

BEGIN_MESSAGE_MAP(Dlg_XuatbangKlg, CDialog)
	ON_BN_CLICKED(XKL_B1, &Dlg_XuatbangKlg::OnBnClickedB1)
	ON_BN_CLICKED(XKL_B2, &Dlg_XuatbangKlg::OnBnClickedB2)
END_MESSAGE_MAP()

// Dlg_XuatbangKlg message handlers
BOOL Dlg_XuatbangKlg::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	// Thiết lập mặc định
	m_chk.SetCheck(1);

	// Xác định sheet hiện hành
	pSheet = pWb->ActiveSheet;
	_bsCNM = pSheet->CodeName;	// shHSNTVL hoặc shHSNTCV
	_bsName = pSheet->Name;
	sNameSh = (LPCTSTR)_bsName;
	pRange = pSheet->Cells;

	// Xác định các giai đoạn tồn tại
	f_Load_hangmuc();

	return TRUE; 
}


// Bắt sự kiện click Enter
BOOL Dlg_XuatbangKlg::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		CDialog::OnCancel();
		return TRUE; 
	}

	return FALSE;
}

void Dlg_XuatbangKlg::xl_xdcot_DMCV(_WorksheetPtr pSh)
{
	_iCot[1] = getColumn(pSh,"DMBB_STT");
	_iCot[2] = getColumn(pSh,"DMBB_MACV");
	_iCot[4] = getColumn(pSh,"DMBB_MAHS");
	_iCot[5] = getColumn(pSh,"DMBB_ND");
	_iCot[8] = getColumn(pSh,"DMBB_TC");
	_iCot[21] = getColumn(pSh,"DMBB_PS");
	_iCot[26] = getColumn(pSh,"DMBB_TDDBT_GIO");

	iNgay[0] = getColumn(pSh,"DMBB_TN_NGAY");
	iNgay[1] = getColumn(pSh,"DMBB_NB_NGAY");
	iNgay[2] = getColumn(pSh,"DMBB_PHIEUYC");
	iNgay[3] = getColumn(pSh,"DMBB_AB_NGAY");
	iNgay[4] = getColumn(pSh,"DMBB_HK_NGAYBD");
	iNgay[5] = getColumn(pSh,"DMBB_DKDBT_NGAY");
	iNgay[6] = getColumn(pSh,"DMBB_TDDBT_NGAY");
}


void Dlg_XuatbangKlg::f_Load_hangmuc()
{
	try
	{
		// Xác định các cột sheet danh mục công việc
		xl_xdcot_DMCV(pSheet);

		// Vùng sắp xếp
		m_cbb1.AddString(L"Tất cả các hạng mục");

		CString szval=L"";
		iEnd = FindComment(pSheet,_iCot[1],"END");
		ibd=8,ikt=iEnd;

		int dem=0;
		for (int i = 8; i < iEnd; i++)
		{
			// Xác định hạng mục
			szval = GIOR(i,_iCot[2],pRange,L"");
			if(szval==L"HM")
			{
				dem++;
				Msrt[dem] = i;
				szval = GIOR(i,_iCot[5],pRange,L"");
				m_cbb1.AddString(szval);
			}
		}

		Msrt[dem+1] = iEnd;
		m_cbb1.SetCurSel(0);

		// Không tồn tại hạng mục
		if(dem==0) m_cbb1.EnableWindow(0);

		// Xác định các giai đoạn
		f_xacdinh_gdoan(ibd,ikt);
	}
	catch(...){_s(L"Lỗi xác định hạng mục");}
}


void Dlg_XuatbangKlg::f_xacdinh_gdoan(int num1,int num2)
{
	try
	{
		m_cbb2.ResetContent();
		ASSERT(m_cbb2.GetCount() == 0);
		m_cbb2.AddString(L"Tất cả các giai đoạn");

		int dem=0;
		CString szval=L"";
		for (int i = num1+1; i < num2; i++)
		{
			szval = GIOR(i,_iCot[2],pRange,L"");
			if(szval==L"GD")
			{
				dem++;
				Msgd[dem] = i;
				szval = GIOR(i,_iCot[5],pRange,L"");
				m_cbb2.AddString(szval);
			}
		}

		Msgd[dem+1] = ikt;
		m_cbb2.SetCurSel(0);
	}
	catch(...){_s(L"Lỗi xác định giai đoạn");}
}


void Dlg_XuatbangKlg::OnBnClickedB1()
{
	CDialog::OnOK();
}


void Dlg_XuatbangKlg::OnBnClickedB2()
{
	CDialog::OnCancel();
}
