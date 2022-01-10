#include "stdafx.h"
#include "Dlg_dsach_ctieu.h"


// Dlg_dsach_ctieu dialog
IMPLEMENT_DYNAMIC(Dlg_dsach_ctieu, CDialogEx)

Dlg_dsach_ctieu::Dlg_dsach_ctieu(CWnd* pParent /*=NULL*/)
	: CDialogEx(Dlg_dsach_ctieu::IDD, pParent)
{

}

Dlg_dsach_ctieu::~Dlg_dsach_ctieu()
{
}

void Dlg_dsach_ctieu::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, DSCT_L1, m_list);
	DDX_Control(pDX, DSCT_CK1, m_check);
	DDX_Control(pDX, DSCT_T1, m_txt);
	DDX_Control(pDX, DSCT_T2, m_t1);
	DDX_Control(pDX, DSCT_T3, m_t2);
}


BEGIN_MESSAGE_MAP(Dlg_dsach_ctieu, CDialogEx)
	ON_BN_CLICKED(IDC_BUTTON3, &Dlg_dsach_ctieu::OnBnClickedButton3)
	ON_BN_CLICKED(IDC_BUTTON1, &Dlg_dsach_ctieu::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON4, &Dlg_dsach_ctieu::OnBnClickedButton4)
	ON_NOTIFY(LVN_KEYDOWN, DSCT_L1, &Dlg_dsach_ctieu::OnLvnKeydownL1)
	ON_BN_CLICKED(IDC_BUTTON5, &Dlg_dsach_ctieu::OnBnClickedButton5)
	ON_NOTIFY(NM_DBLCLK, DSCT_L1, &Dlg_dsach_ctieu::OnNMDblclkL1)
END_MESSAGE_MAP()


// Dlg_dsach_ctieu message handlers
BOOL Dlg_dsach_ctieu::OnInitDialog()
{
	CDialog::OnInitDialog();

	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	m_t1.SetCueBanner(L"Thêm tên chỉ tiêu");
	m_t2.SetCueBanner(L"Đơn vị");
	m_txt.SetWindowText(xl_tukhoa);
	m_txt.SetCueBanner(L"Nhập từ khóa tìm kiếm");
	f_Style_dsach();
	f_Load_dulieu();

	return TRUE;
}


BOOL Dlg_dsach_ctieu::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if (pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_RETURN &&	pWndWithFocus == GetDlgItem(DSCT_T1))
	{
		OnBnClickedButton4();
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_RETURN &&	pWndWithFocus == GetDlgItem(DSCT_L1))
	{
		OnBnClickedButton1();
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		CDialog::OnCancel();
		return TRUE; 
	}

	return FALSE;
}


void Dlg_dsach_ctieu::f_Style_dsach()
{
	m_list.InsertColumn(1,L"",LVCFMT_CENTER,0);
	m_list.InsertColumn(2,L"Tên chỉ tiêu",LVCFMT_LEFT,300);
	m_list.InsertColumn(3,L"Đơn vị",LVCFMT_CENTER,80);

	m_list.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);
}


void Dlg_dsach_ctieu::f_Delete_list()
{
	m_list.DeleteAllItems();
	ASSERT(m_list.GetItemCount() == 0);
}


void Dlg_dsach_ctieu::OnBnClickedButton3()
{
	OnCancel();
}


void Dlg_dsach_ctieu::OnBnClickedButton1()
{
	try
	{
		// Xác định vị trí sẽ thêm mới vào list quy cách
		int dem = listQclm.GetItemCount();
		int vtr=-1;
		POSITION pos=listQclm.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			for (int i = 0; i < dem; i++)
				if (listQclm.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					vtr=i;
					break;
				}
		}

		if(vtr==-1) vtr = dem;

		// Xác định các giá trị sẽ sử dụng tại list chỉ tiêu mẫu
		int count = m_list.GetItemCount();
		if(count==0) return;

		pos=m_list.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			dem = 0;
			CString szval=L"",vl1=L"",vl2=L"";
			for (int i = 0; i < count; i++)
				if (m_list.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					if(vtr+dem<9) szval.Format(L"0%d",vtr+dem+1);
					else szval.Format(L"%d",vtr+dem+1);

					vl1 = m_list.GetItemText(i,1);
					vl2 = m_list.GetItemText(i,2);

					listQclm.InsertItem(vtr+dem,szval,0);
					listQclm.SetItemText(vtr+dem,1,vl1);
					listQclm.SetItemText(vtr+dem,2,vl2);

					// Tô màu
					if(vl1.Left(1)==L"*") listQclm.SetRowColors(vtr+dem,YELLOW,BLUE);
					else listQclm.SetRowColors(vtr+dem,WHITE,BLUE);

					dem++;
				}

			iChekUp=1;
			f_Sort_stt_qc();
		}

		if(m_check.GetCheck()==0) OnOK();
	}
	catch(...){}
}


void Dlg_dsach_ctieu::OnBnClickedButton4()
{
	// Tìm kiếm dữ liệu
	f_Load_dulieu();
}


void Dlg_dsach_ctieu::f_Tack_tukhoa(CString sKey)
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
			}
		}
	}
	catch(...){}
}


void Dlg_dsach_ctieu::f_Tao_truyvan(CString sBosung)
{
	try
	{
		CString val=L"";
		if(iKey==1)
		{
			SqlString.Format(L"SELECT * FROM tbDanhsach "
			L"WHERE ((DSTen LIKE '%s%s%s' OR Donvi LIKE '%s%s%s')%s) ORDER BY DSTen ASC;",L"%",pKey[1],L"%",L"%",pKey[1],L"%",sBosung);
		}
		else
		{
			SqlString.Format(L"SELECT * FROM tbDanhsach " 
			L"WHERE ((DSTen LIKE '%s%s%s' OR Donvi LIKE '%s%s%s')",L"%",pKey[1],L"%",L"%",pKey[1],L"%");

			for (int i = 2; i <= iKey; i++)
			{
				val.Format(
					L" AND (DSTen LIKE '%s%s%s' OR Donvi LIKE '%s%s%s')",L"%",pKey[i],L"%",L"%",pKey[i],L"%");
				SqlString+=val;
			}

			val.Format(L"%s%s) ORDER BY DSTen ASC;",SqlString,sBosung);
			SqlString = val;
		}
	}
	catch(...){}
}


void Dlg_dsach_ctieu::f_Load_dulieu()
{
	TRY
	{
		if(connectDsn(pathQCLM)==-1) return;
		ObjConn.OpenAccessDB(pathQCLM, L"", L"");

		CString szval=L"";
		m_txt.GetWindowTextW(szval);
		szval.Trim();
		szval.Replace(L"'",L"");

		int dem = m_list.GetItemCount();
		if(dem>0) f_Delete_list();

		iKqua=0;
		CString vl1=L"",vl2=L"",vl3=L"";
		
		if(szval==L"") SqlString = L"SELECT * FROM tbDanhsach WHERE (ID <> '' AND ID IS NOT NULL) ORDER BY DSTen ASC;";
		else
		{
			f_Tack_tukhoa(szval);
			f_Tao_truyvan(L"");
		}

		CRecordset* Recset = ObjConn.Execute(SqlString);
		while( !Recset->IsEOF() )
		{
			Recset->GetFieldValue(L"ID",vl1);
			Recset->GetFieldValue(L"DSTen",vl2);
			Recset->GetFieldValue(L"Donvi",vl3);

			m_list.InsertItem(iKqua,vl1,0);
			m_list.SetItemText(iKqua,1,vl2);
			m_list.SetItemText(iKqua,2,vl3);

			iKqua++;
			Recset->MoveNext();
		}
		Recset->Close();

		ObjConn.CloseDatabase();
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("Lỗi xác định cơ sở dữ liệu.")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_dsach_ctieu::f_Delete_item()
{
	TRY
	{
		int count = m_list.GetItemCount();
		if(count==0) return;

		if(_yn(L"Chỉ tiêu được lựa chọn sẽ bị xóa khỏi CSDL! Bạn có chắc chắn xóa?",0)!=6) return;

		if(connectDsn(pathQCLM)==-1) return;
		ObjConn.OpenAccessDB(pathQCLM, L"", L"");

		POSITION pos=m_list.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			CString vl1=L"",vl2=L"";
			for (int i = count; i >= 0; i--)
				if (m_list.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					vl1 = m_list.GetItemText(i,1);
					vl2 = m_list.GetItemText(i,2);

					m_list.DeleteItem(i);
					SqlString.Format(L"DELETE FROM tbDanhsach WHERE (DSTen='%s' AND Donvi='%s');",vl1,vl2);
					ObjConn.ExecuteRB(SqlString);
				}
		}

		ObjConn.CloseDatabase();
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("Lỗi xác định cơ sở dữ liệu.")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_dsach_ctieu::OnLvnKeydownL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);
	if(pLVKeyDow->wVKey == VK_DELETE) f_Delete_item();


	*pResult = 0;
}


void Dlg_dsach_ctieu::OnBnClickedButton5()
{
	TRY
	{
		CString svl1=L"",svl2=L"";
		m_t1.GetWindowTextW(svl1);
		m_t2.GetWindowTextW(svl2);
		svl1.Trim();
		svl1.Replace(L"'",L"");
		svl2.Trim();
		svl2.Replace(L"'",L"");

		if(svl1!=L"")
		{
			if(connectDsn(pathQCLM)==-1) return;
			ObjConn.OpenAccessDB(pathQCLM, L"", L"");
			CRecordset* Recset;
			CString svl0=L"",sdem=L"";

			// Xác định ID mới (Dxxx)
			for (int j = 1; j <= 999; j++)
			{
				if(j<=9) svl0.Format(L"D00%d",j);
				else if(j>9 && j<=99) svl0.Format(L"D0%d",j);
				else svl0.Format(L"D%d",j);

				SqlString.Format(L"SELECT COUNT(*) AS iDem FROM tbDanhsach WHERE (ID = '%s');",svl0);
				Recset = ObjConn.Execute(SqlString);
				Recset->GetFieldValue(L"iDem",sdem);
				Recset->Close();

				if(_wtoi(sdem)==0) break;
			}

			// Thêm mới vào CSDL
			SqlString.Format(
				L"INSERT INTO tbDanhsach (ID,DSTen,Donvi) VALUES ('%s','%s','%s');",svl0,svl1,svl2);
			ObjConn.ExecuteRB(SqlString);

			// Thêm mới vào list
			int count = m_list.GetItemCount();
			m_list.InsertItem(count,svl0,0);
			m_list.SetItemText(count,1,svl1);
			m_list.SetItemText(count,2,svl2);

			ObjConn.CloseDatabase();
		}
		else GotoDlgCtrl(GetDlgItem(DSCT_T2));
		
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("Lỗi xác định cơ sở dữ liệu.")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_dsach_ctieu::OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	OnBnClickedButton1();
}