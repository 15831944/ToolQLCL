#include "stdafx.h"
#include "Dlg_noidungppktra.h"


// Dlg_noidungppktra dialog
IMPLEMENT_DYNAMIC(Dlg_noidungppktra, CDialogEx)

Dlg_noidungppktra::Dlg_noidungppktra(CWnd* pParent /*=NULL*/)
	: CDialogEx(Dlg_noidungppktra::IDD, pParent)
{

}

Dlg_noidungppktra::~Dlg_noidungppktra()
{
}

void Dlg_noidungppktra::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, NDPPKT_T1, m_Txt);
	DDX_Control(pDX, NDPPKT_T2, m_Them);
	DDX_Control(pDX, NDPPKT_L1, m_List);
}


BEGIN_MESSAGE_MAP(Dlg_noidungppktra, CDialogEx)
	ON_BN_CLICKED(NDPPKT_B2, &Dlg_noidungppktra::OnBnClickedB2)
	ON_BN_CLICKED(NDPPKT_B3, &Dlg_noidungppktra::OnBnClickedB3)
	ON_BN_CLICKED(NDPPKT_B1, &Dlg_noidungppktra::OnBnClickedB1)
	ON_BN_CLICKED(NDPPKT_CHK1, &Dlg_noidungppktra::OnBnClickedChk1)
	ON_BN_CLICKED(NDPPKT_B4, &Dlg_noidungppktra::OnBnClickedB4)
	ON_NOTIFY(NM_DBLCLK, NDPPKT_L1, &Dlg_noidungppktra::OnNMDblclkL1)
END_MESSAGE_MAP()


// Dlg_noidungppktra message handlers
BOOL Dlg_noidungppktra::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	iClose = 0,iTang = 0;

	CString szval=L"Nội dung kiểm tra";
	if(jTMCV==2) szval=L"Phương pháp kiểm tra";
	this->SetWindowTextW(szval);

	m_List.InsertColumn(0,L"",LVCFMT_LEFT,0);
	m_List.InsertColumn(1,szval,LVCFMT_LEFT,500);
	m_List.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);
	
	m_Txt.SetWindowText(xl_tukhoa);
	m_Txt.SetCueBanner(L"Nhập từ khóa tìm kiếm (Ví dụ: Đóng và ép cọc,...)");
	m_Them.SetCueBanner(L"Nhập nội dung/ phương pháp cần thêm mới (Ví dụ: Biện pháp gia công,...)");	
	f_Load_DL();

	return TRUE; 
}


BOOL Dlg_noidungppktra::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if (pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN && pWndWithFocus == GetDlgItem(NDPPKT_L1)) OnBnClickedB1();
	else if (pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN && pWndWithFocus == GetDlgItem(NDPPKT_T1)) OnBnClickedB3();
	else if (pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN && pWndWithFocus == GetDlgItem(NDPPKT_T2)) OnBnClickedB4();
	else if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		CDialog::OnCancel();
		return TRUE; 
	}

	return FALSE;
}


void Dlg_noidungppktra::f_Load_DL()
{
	TRY
	{
		CString szval=L"NoidungCV";
		if(jTMCV==2) szval=L"PPkiemtra";

		// Kết nối SQL
		ObjConn.OpenAccessDB(pathMDB,L"",L"");
		CRecordset* Recset = ObjConn.Execute(SqlString);
		int dem=0;
		CString sGan=L"";
		while( !Recset->IsEOF() )
		{
			Recset->GetFieldValue(szval,sGan);
			m_List.InsertItem(dem,L"",0);
			m_List.SetItemText(dem,1,sGan);
			dem++;

			Recset->MoveNext();
		}
		Recset->Close();
		ObjConn.CloseDatabase();
		if(dem>0) GotoDlgCtrl(GetDlgItem(NDPPKT_L1));
		else m_Txt.SetSel(0,-1);
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("...")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_noidungppktra::f_Tack_tukhoa(CString sKey)
{
	try
	{
		sKey+=L"+";
		iKey=0;
		int vtri=0;
		CString val=L"";
		for (int i = 0; i < sKey.GetLength(); i++)
		{
			val = sKey.Mid(i,1);
			if((val==L" " || val==L"+") && i>vtri)
			{
				iKey++;
				pKey[iKey] = sKey.Mid(vtri,i-vtri);
				vtri=i+1;

				if(iKey==98) break;
			}
		}
	}
	catch(...){}
}


void Dlg_noidungppktra::f_Tao_truyvan()
{
	try
	{
		CString sztbl=L"tbNoidungCV",sft=L"NoidungCV";
		if(jTMCV==2)
		{
			sztbl=L"tbPhuongphapKT";
			sft=L"PPkiemtra";
		}

		CString szval=L"";
		if(iKey==1) SqlString.Format(L"SELECT * FROM %s WHERE (%s LIKE '%s%s%s');",sztbl,sft,L"%",pKey[1],L"%");
		else
		{
			SqlString.Format(L"SELECT * FROM %s WHERE ((%s LIKE '%s%s%s')",sztbl,sft,L"%",pKey[1],L"%");
			for (int i = 2; i <= iKey; i++)
			{
				szval.Format(L" AND (%s LIKE '%s%s%s')",sft,L"%",pKey[i],L"%");
				SqlString+=szval;
			}
			SqlString+= L");";
		}
	}
	catch(...){}
}


void Dlg_noidungppktra::f_Truy_van_DL()
{
	TRY
	{
		CString sft=L"NoidungCV";
		if(jTMCV==2) sft=L"PPkiemtra";

		ObjConn.OpenAccessDB(pathMDB,L"",L"");
		CRecordset* Recset = ObjConn.Execute(SqlString);
		iKqua=0;
		while( !Recset->IsEOF() )
		{
			iKqua++;
			if(iKqua>1) break;
			Recset->GetFieldValue(sft,XLXD[1].CDG[0]);
			Recset->MoveNext();
		}

		Recset->Close();
		ObjConn.CloseDatabase();
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("...")+e->m_strError);
	}
	END_CATCH;
}


int Dlg_noidungppktra::f_Tim_kiem_NDPP()
{
	try
	{
		CString sztbl=L"tbNoidungCV",sft=L"NoidungCV";
		if(jTMCV==2)
		{
			sztbl=L"tbPhuongphapKT";
			sft=L"PPkiemtra";
		}

		// Tạo truy vấn tra cứu công việc		
		if(xl_tukhoa==L"") SqlString.Format(L"SELECT * FROM %s ORDER BY %s ASC;",sztbl,sft);
		else
		{
			f_Tack_tukhoa(xl_tukhoa);
			f_Tao_truyvan();
		}		

		// Truy vấn dữ liệu
		f_Truy_van_DL();
		if(iKqua==1)
		{
			lit_ndkt.SetItemText(CLRow,jTMCV,XLXD[1].CDG[0]);
			lit_ndkt.SetRowColors(CLRow,RGB(255,255,255),RGB(0,0,0));
		}
		else if(iKqua>1)
		{
			// Có nhiều kết quả -> hiển thị hộp thoại
			AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
			DoModal();

			return 1;
		}
		else lit_ndkt.SetRowColors(CLRow,RGB(255,255,255),RGB(255,0,0));

		return 0;
	}
	catch(...){}

	return 0;
}


void Dlg_noidungppktra::OnBnClickedB2()
{
	CDialog::OnCancel();
}


void Dlg_noidungppktra::OnBnClickedB3()
{
	try
	{
		CString sztbl=L"tbNoidungCV",sft=L"NoidungCV";
		if(jTMCV==2)
		{
			sztbl=L"tbPhuongphapKT";
			sft=L"PPkiemtra";
		}

		// Xóa toàn bộ dữ liệu trên List Control
		m_List.DeleteAllItems();
		ASSERT(m_List.GetItemCount() == 0);

		// Xác định từ khóa
		CString szval=L"";
		m_Txt.GetWindowTextW(szval);
		szval.Replace(L"'",L"");
		szval.Trim();

		// Truy vấn & load dữ liệu
		if(szval==L"") SqlString.Format(L"SELECT * FROM %s ORDER BY %s ASC;",sztbl,sft);
		else
		{
			f_Tack_tukhoa(szval);
			f_Tao_truyvan();
		}

		f_Load_DL();
	}
	catch(...){}
}


void Dlg_noidungppktra::OnBnClickedB1()
{
	try
	{
		POSITION pos=m_List.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			int nCount = m_List.GetItemCount();
			for(int i = 0; i < nCount; i++)		
			{
				if (m_List.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					// Đổ dữ liệu lên bảng công việc
					if(iTang>0) lit_ndkt.InsertItem(CLRow+iTang,L"",0);
					lit_ndkt.SetItemText(CLRow+iTang,0,L"1");
					lit_ndkt.SetItemText(CLRow+iTang,jTMCV,m_List.GetItemText(i,1));
					lit_ndkt.SetRowColors(CLRow+iTang,RGB(255,255,255),RGB(0,0,255));

					break;
				}
			}

			if(iClose==0) CDialog::OnOK();
			else iTang++;
		}
	}
	catch(...){}
}


void Dlg_noidungppktra::OnBnClickedChk1()
{
	CButton *m_ctlCheck = (CButton*) GetDlgItem(NDPPKT_CHK1);
	if(m_ctlCheck->GetCheck() == BST_UNCHECKED) iClose=0;
	else iClose=1;
}


void Dlg_noidungppktra::OnBnClickedB4()
{
	m_Them.SetSel(0,-1);
}


void Dlg_noidungppktra::OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	OnBnClickedB1();
}
