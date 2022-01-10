#include "Dlg_all_tieuchuan.h"

Dlg_all_tieuchuan::Dlg_all_tieuchuan(CWnd* pParent /*=NULL*/)
	: CDialog(Dlg_all_tieuchuan::IDD, pParent)
{
	
}

void Dlg_all_tieuchuan::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, ALTC_T1, m_Txt);
	DDX_Control(pDX, ALTC_T2, m_Matc);
	DDX_Control(pDX, ALTC_T3, m_Tentc);
	DDX_Control(pDX, ALTC_T4, m_Loaitc);
	DDX_Control(pDX, ALTC_L1, m_List);
	DDX_Control(pDX, ALTC_G1, m_Group);
	DDX_Control(pDX, ALTC_B3, m_btn[3]);	
}

BEGIN_MESSAGE_MAP(Dlg_all_tieuchuan, CDialog)
	ON_WM_SIZE()
	ON_WM_SIZING()
	ON_WM_CONTEXTMENU()
	ON_BN_CLICKED(ALTC_B1, &Dlg_all_tieuchuan::OnBnClickedB1)
	ON_BN_CLICKED(ALTC_B2, &Dlg_all_tieuchuan::OnBnClickedB2)
	ON_BN_CLICKED(ALTC_B4, &Dlg_all_tieuchuan::OnBnClickedB4)
	ON_BN_CLICKED(ALTC_B3, &Dlg_all_tieuchuan::OnBnClickedB3)
	ON_BN_CLICKED(ALTC_B5, &Dlg_all_tieuchuan::OnBnClickedB5)
	ON_NOTIFY(LVN_KEYDOWN, ALTC_L1, &Dlg_all_tieuchuan::OnLvnKeydownL1)
	ON_BN_CLICKED(ALTC_CK1, &Dlg_all_tieuchuan::OnBnClickedCk1)
	ON_NOTIFY(NM_CLICK, ALTC_L1, &Dlg_all_tieuchuan::OnNMClickL1)	
	ON_NOTIFY(NM_DBLCLK, ALTC_L1, &Dlg_all_tieuchuan::OnNMDblclkL1)	
	ON_NOTIFY(NM_RCLICK, ALTC_L1, &Dlg_all_tieuchuan::OnNMRClickL1)
	ON_COMMAND(ID_TI40039_8, &Dlg_all_tieuchuan::OnTi40039)
	ON_COMMAND(ID_TI40040_8, &Dlg_all_tieuchuan::OnTi40040)
END_MESSAGE_MAP()

BEGIN_EASYSIZE_MAP(Dlg_all_tieuchuan)
	EASYSIZE(ALTC_S1,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(ALTC_T1,ES_BORDER,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(ALTC_B1,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)

	EASYSIZE(ALTC_G1,ES_BORDER,ES_BORDER,ES_BORDER,ES_BORDER,0)
	EASYSIZE(ALTC_L1,ES_BORDER,ES_BORDER,ES_BORDER,ES_BORDER,0)

	EASYSIZE(ALTC_B2,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)
	EASYSIZE(ALTC_T2,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)
	EASYSIZE(ALTC_T3,ES_BORDER,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(ALTC_T4,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(ALTC_B4,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)

	EASYSIZE(ALTC_CK1,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)
	EASYSIZE(ALTC_B3,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(ALTC_B5,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
END_EASYSIZE_MAP

BOOL Dlg_all_tieuchuan::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	iTCclick=-1;
	sDownTC=L"",sFullTC=L"";

	iVtTC = CLRow;
	iClose = 0,iTang = 0;	
	if(jTTTC==0) m_btn[3].EnableWindow(1);

	f_Load_ToolTip();
	m_Txt.SetWindowText(xl_tukhoa);
	m_Txt.SetCueBanner(L"Nhập từ khóa tìm kiếm (Ví dụ: TCVN6355, Thép cốt bê tông 2009, ...)");
	m_Matc.SetCueBanner(L"Mã TC");
	m_Tentc.SetCueBanner(L"Nhập tên tiêu chuẩn mới vào đây...");
	m_Loaitc.SetCueBanner(L"Loại TC");

	f_StyleList();
	f_Load_DL();

	if(qThemTC==1)
	{
		CButton *m1 = (CButton*) GetDlgItem(ALTC_CK1);
		m1->EnableWindow(0);

		CButton *m2 = (CButton*) GetDlgItem(ALTC_B3);
		m2->EnableWindow(0);
	}

	INIT_EASYSIZE;

	return TRUE; 
}

// Bắt sự kiện click Enter
BOOL Dlg_all_tieuchuan::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if (pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(ALTC_T1))
	{
		OnBnClickedB1();
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(ALTC_T2))
	{
		GotoDlgCtrl(GetDlgItem(ALTC_T3));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(ALTC_T3))
	{
		GotoDlgCtrl(GetDlgItem(ALTC_T4));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(ALTC_T4))
	{
		GotoDlgCtrl(GetDlgItem(ALTC_T2));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(ALTC_L1))
	{
		OnBnClickedB3();
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		f_SaveWidth();
		CDialog::OnCancel();
		return TRUE; 
	}
	else if( pMsg->message == WM_LBUTTONDOWN ||
		pMsg->message == WM_LBUTTONUP ||
		pMsg->message == WM_MOUSEMOVE )
	{
		m_ToolTips.RelayEvent(pMsg);
		return CDialog::PreTranslateMessage(pMsg);
	}
	else if(pMsg->message == WM_CHAR)
	{
		TCHAR chr = static_cast<TCHAR>(pMsg->wParam);
		if(pWndWithFocus == GetDlgItem(ALTC_L1))
		{
			switch (chr)
			{
				case 0x04:
				{
					OnTi40039();
					break;
				}
				case 0x07:
				{
					OnTi40040();
					break;
				}

				default:
					break;
			}
		}
	}

	return FALSE;
}

void Dlg_all_tieuchuan::OnSize(UINT nType, int cx, int cy) 
{
	CDialog::OnSize(nType, cx, cy);
	UPDATE_EASYSIZE;
}

void Dlg_all_tieuchuan::OnSizing(UINT fwSide, LPRECT pRect) 
{
	CDialog::OnSizing(fwSide, pRect);
	EASYSIZE_MINSIZE(320,240,fwSide,pRect);
}


void Dlg_all_tieuchuan::f_Load_ToolTip()
{
	CListCtrl * pls1 = (CListCtrl *)GetDlgItem(ALTC_L1);
	CButton * btn1 = (CButton *)GetDlgItem(ALTC_B2);
	CButton * btn2 = (CButton *)GetDlgItem(ALTC_B4);
	
	m_ToolTips.Create(this);
	m_ToolTips.SetMaxTipWidth(600);
	m_ToolTips.SetDelayTime(900);

	m_ToolTips.AddTool(pls1, L"Kích đúp để áp dụng tiêu chuẩn. Giữ Shift/Ctrl để chọn nhiều");
	m_ToolTips.AddTool(btn1, L"Tạo mới tiêu chuẩn.");
	m_ToolTips.AddTool(btn2, L"Cập nhật tiêu chuẩn chỉnh sửa.");
}


void Dlg_all_tieuchuan::f_StyleList()
{
	m_List.InsertColumn(0,L"",LVCFMT_LEFT,0);
	m_List.InsertColumn(1,L"Mã TC",LVCFMT_LEFT,_wL4[0]);
	m_List.InsertColumn(2,L"Tên tiêu chuẩn",LVCFMT_LEFT,_wL4[1]);
	m_List.InsertColumn(3,L"Loại tiêu chuẩn",LVCFMT_LEFT,_wL4[2]);
	m_List.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);
}


void Dlg_all_tieuchuan::f_SaveWidth()
{
	for (int i = 0; i < 3; i++) _wL4[i] = m_List.GetColumnWidth(i+1);
}


void Dlg_all_tieuchuan::f_Tack_tukhoa(CString sKey)
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


void Dlg_all_tieuchuan::f_Tao_truyvan()
{
	try
	{
		CString val=L"";
		if(iKey==1)
		{
			SqlString.Format(L"SELECT * FROM tbTentieuchuan "
			L"WHERE (MaTC LIKE '%s%s%s' OR TenTC LIKE '%s%s%s' OR Ghichu LIKE '%s%s%s');"
			,L"%",pKey[1],L"%",L"%",pKey[1],L"%",L"%",pKey[1],L"%");
		}
		else
		{
			SqlString.Format(L"SELECT * FROM tbTentieuchuan " 
			L"WHERE ((MaTC LIKE '%s%s%s' OR TenTC LIKE '%s%s%s' OR Ghichu LIKE '%s%s%s')"
			,L"%",pKey[1],L"%",L"%",pKey[1],L"%",L"%",pKey[1],L"%");

			for (int i = 2; i <= iKey; i++)
			{
				val.Format(
					L" AND (MaTC LIKE '%s%s%s' OR TenTC LIKE '%s%s%s' OR Ghichu LIKE '%s%s%s')"
					,L"%",pKey[i],L"%",L"%",pKey[i],L"%",L"%",pKey[i],L"%");
				SqlString+=val;
			}
			SqlString+= L");";
		}
	}
	catch(...){}
}


void Dlg_all_tieuchuan::OnLvnKeydownL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);
	if(pLVKeyDow->wVKey == VK_SPACE) OnBnClickedB3();
	else if(pLVKeyDow->wVKey == VK_DELETE) f_Xoa_tieuchuan();

	*pResult = 0;
}


void Dlg_all_tieuchuan::f_Xoa_tieuchuan()
{
	TRY
	{
		int nCount = m_List.GetItemCount();
		if(nCount<=0) return;

		POSITION pos=m_List.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			int dem=0;
			for (int i = 0; i < nCount; i++)
				if (m_List.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED) dem++;

			CString sTvan,val = L"";
			if(dem==1) val = L"Bạn có chắc chắn sẽ xóa tiêu chuẩn này trong dữ liệu?";
			else val = L"Bạn có chắc chắn sẽ xóa những tiêu chuẩn này trong dữ liệu?";
			
			int n = _yn(val,0);
			if (n == 6)
			{
				// Kết nối SQL
				ObjConn.OpenAccessDB(pathMDB,L"",L"");
				
				for (int i = nCount-1; i >=0; i--)
				{
					if (m_List.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
					{
						// Xóa trong CSDL
						val = m_List.GetItemText(i,1);
						sTvan.Format(L"DELETE FROM tbTentieuchuan WHERE MaTC='%s';",val);
						ObjConn.ExecuteRB(sTvan);

						// Xóa trên list control
						m_List.DeleteItem(i);
					}
				}

				ObjConn.CloseDatabase();
			}
		}
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[224] Lỗi xóa dữ liệu tiêu chuẩn")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_all_tieuchuan::f_Truy_van_DL()
{
	TRY
	{
		ObjConn.OpenAccessDB(pathMDB,L"",L"");
		CRecordset* Recset = ObjConn.Execute(SqlString);
		iKqua=0;
		while( !Recset->IsEOF() )
		{
			iKqua++;
			if(iKqua>1) break;
			Recset->GetFieldValue(L"MaTC",XLXD[1].CDG[0]);
			Recset->GetFieldValue(L"TenTC",XLXD[1].CDG[1]);
			Recset->GetFieldValue(L"LoaiTC",XLXD[1].CDG[2]);
			Recset->MoveNext();
		}
		Recset->Close();
		ObjConn.CloseDatabase();
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[1] Lỗi tìm kiếm dữ liệu công việc")+e->m_strError);
	}
	END_CATCH;
}


int Dlg_all_tieuchuan::f_Tim_kiem_TC()
{
	try
	{
		// Tạo truy vấn tra cứu công việc
		if(xl_tukhoa==L"") SqlString = L"SELECT * FROM tbTentieuchuan ORDER BY MaTC ASC;";
		else
		{
			f_Tack_tukhoa(xl_tukhoa);
			f_Tao_truyvan();
		}		

		// Truy vấn dữ liệu
		f_Truy_van_DL();

		if(iKqua==1)
		{
			if(_pcaltc==0)
			{
				// Đổ dữ liệu lên bảng công việc
				txtE_TCCV.SetWindowText(XLXD[1].CDG[1]);
				lit_tccv.SetItemText(CLRow,1,XLXD[1].CDG[0]);
				lit_tccv.SetItemText(CLRow,2,XLXD[1].CDG[1]);
				lit_tccv.SetRowColors(CLRow,RGB(255,255,255),RGB(0,0,0));
			}
			else
			{
				// Đổ dữ liệu lên bảng vật liệu
				txtE_TCVL .SetWindowText(XLXD[1].CDG[1]);
				lit_tcvl.SetItemText(CLRow,1,XLXD[1].CDG[0]);
				lit_tcvl.SetItemText(CLRow,2,XLXD[1].CDG[1]);
				lit_tcvl.SetRowColors(CLRow,RGB(255,255,255),RGB(0,0,0));
			}			
		}
		else if(iKqua>1)
		{
			// Có nhiều kết quả -> hiển thị hộp thoại
			AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
			DoModal();

			return 1;
		}
		else lit_tccv.SetRowColors(CLRow,RGB(255,255,255),RGB(255,0,0));

		return 0;
	}
	catch(...){_s(L"#TC237: Lỗi tìm kiếm tiêu chuẩn.");}

	return 0;
}


void Dlg_all_tieuchuan::f_Load_DL()
{
	TRY
	{
		ObjConn.OpenAccessDB(pathMDB,L"",L"");
		CRecordset* Recset = ObjConn.Execute(SqlString);
		int dem=0;
		CString sGan[3];
		while( !Recset->IsEOF() )
		{
			Recset->GetFieldValue(L"MaTC",sGan[0]);
			Recset->GetFieldValue(L"TenTC",sGan[1]);
			Recset->GetFieldValue(L"LoaiTC",sGan[2]);

			m_List.InsertItem(dem,L"",0);
			m_List.SetItemText(dem,1,sGan[0]);
			m_List.SetItemText(dem,2,sGan[1]);
			m_List.SetItemText(dem,3,sGan[2]);
			dem++;

			Recset->MoveNext();
		}
		Recset->Close();
		ObjConn.CloseDatabase();

		// Kết quả tìm kiếm
		sGan[0].Format(L"Kết quả tìm kiếm (%d tiêu chuẩn)",dem);
		m_Group.SetWindowText(sGan[0]);
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[82] Lỗi tìm kiếm dữ liệu tiêu chuẩn")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_all_tieuchuan::OnBnClickedB1()
{
	try
	{
		// Xóa toàn bộ dữ liệu trên List Control
		m_List.DeleteAllItems();
		ASSERT(m_List.GetItemCount() == 0);

		// Xác định từ khóa
		CString val=L"";
		m_Txt.GetWindowTextW(val);
		val.Replace(L"'",L"");
		val.Trim();

		// Truy vấn & load dữ liệu
		if(val==L"") SqlString = L"SELECT * FROM tbTentieuchuan ORDER BY MaTC ASC;";
		else
		{
			f_Tack_tukhoa(val);
			f_Tao_truyvan();
		}

		f_Load_DL();
	}
	catch(...){}
}


void Dlg_all_tieuchuan::OnBnClickedB2()
{
	TRY
	{
		// Số ngẫu nhiên theo thời gian
		CString val=L"";
		time_t now = time(0);
		tm *ltm = localtime(&now);
		val.Format(L"%d%d",ltm->tm_min,ltm->tm_sec);
		int _nRand = _wtoi(val);

		// Kiểm tra dữ liệu có trong CSDL không?
		CString sTvan=L"",sDem=L"",jNew=L"";
		jNew.Format(L"TCNW.%d%d%d",_nRand,rand()%10,rand()%10);
		sTvan.Format(L"SELECT COUNT(*) AS iDem FROM tbTentieuchuan WHERE MaTC = '%s';",jNew);
		
		// Kết nối SQL
		ObjConn.OpenAccessDB(pathMDB,L"",L"");		
		CRecordset* Recset = ObjConn.Execute(sTvan);
		Recset->GetFieldValue(L"iDem",sDem);
		Recset->Close();

		while (_wtoi(sDem)>0)
		{
			sDem=L"";
			jNew.Format(L"TCNW.%d%d%d",_nRand,rand()%10,rand()%10);
			sTvan.Format(L"SELECT COUNT(*) AS iDem FROM tbTentieuchuan WHERE MaTC = '%s';",jNew);
			Recset = ObjConn.Execute(sTvan);
			Recset->GetFieldValue(L"iDem",sDem);
			Recset->Close();
		}
		ObjConn.CloseDatabase();

		// Mã được sinh ngẫu nhiên là jNew
		m_Matc.SetWindowText(jNew);
		m_Tentc.SetWindowText(L"");
		m_Loaitc.SetWindowText(L"");
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[41] Lỗi thêm mới tiêu chuẩn")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_all_tieuchuan::OnBnClickedB4()
{
	TRY
	{
		CString s1=L"",s2=L"",s3=L"";
		m_Matc.GetWindowTextW(s1); s1.Trim();
		if(s1==L"")
		{
			GotoDlgCtrl(GetDlgItem(ALTC_T2));
			return;
		}

		// Tên & loại tiêu chuẩn
		m_Tentc.GetWindowTextW(s2);
		m_Loaitc.GetWindowTextW(s3);
		s2.Replace(L"'",L""); s2.Trim();
		s3.Replace(L"'",L""); s3.Trim();

		int nItem=-1;
		int nCount = m_List.GetItemCount();

		POSITION pos = m_List.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			for (int i = 0; i < nCount; i++)
			if (m_List.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
			{
				nItem=i;
				break;
			}
		}
		if(nItem==-1) nItem = 0;

		// Kiểm tra dữ liệu có trong CSDL không?
		CString sTvan=L"",sDem=L"";
		sTvan.Format(L"SELECT COUNT(*) AS iDem FROM tbTentieuchuan WHERE MaTC = '%s';",s1);
		
		// Kết nối SQL
		ObjConn.OpenAccessDB(pathMDB,L"",L"");		
		CRecordset* Recset = ObjConn.Execute(sTvan);
		Recset->GetFieldValue(L"iDem",sDem);
		Recset->Close();

		if (_wtoi(sDem)>0)
		{
			int n0 = _yn(L"Dữ liệu đã tồn tại. Bạn có muốn thay thế bằng dữ liệu mới?",0);
			if(n0==7)
			{
				ObjConn.CloseDatabase();
				return;
			}

			// Tồn tại tiêu chuẩn --> Update
			sTvan.Format(L"UPDATE tbTentieuchuan "
				L"SET TenTC='%s', LoaiTC='%s' WHERE MaTC='%s';",s2,s3,s1);
			ObjConn.ExecuteRB(sTvan);
		}
		else
		{
			sTvan.Format(L"INSERT INTO tbTentieuchuan (MaTC,TenTC,LoaiTC) "
				L"VALUES ('%s','%s','%s');",s1,s2,s3);
			ObjConn.ExecuteRB(sTvan);

			// Thêm mới hoàn toàn --> Chèn dòng vào list
			m_List.InsertItem(nItem,L"",0);
		}

		// Đổ dữ liệu vào list
		m_List.SetItemText(nItem,0,L"");
		m_List.SetItemText(nItem,1,s1);
		m_List.SetItemText(nItem,2,s2);
		m_List.SetItemText(nItem,3,s3);

		ObjConn.CloseDatabase();
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[21] Lỗi thêm mới dữ liệu tiêu chuẩn")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_all_tieuchuan::OnBnClickedB3()
{
	try
	{
		if(qThemTC==1) return;

		POSITION pos=m_List.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			int nCount = m_List.GetItemCount();
			for(int i = 0; i < nCount; i++)		
			{
				if (m_List.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					if(_pcaltc==0)
					{
						// Đổ dữ liệu lên bảng công việc
						if(iTang>0) lit_tccv.InsertItem(iVtTC+iTang,L"",0);
						lit_tccv.SetItemText(iVtTC+iTang,0,L"1");
						lit_tccv.SetItemText(iVtTC+iTang,1,m_List.GetItemText(i,1));
						lit_tccv.SetItemText(iVtTC+iTang,2,m_List.GetItemText(i,2));
						lit_tccv.SetRowColors(iVtTC+iTang,RGB(255,255,255),RGB(0,0,255));
					}
					else
					{
						// Đổ dữ liệu lên bảng vật liệu
						if(iTang>0) lit_tcvl.InsertItem(iVtTC+iTang,L"",0);
						lit_tcvl.SetItemText(iVtTC+iTang,0,L"1");
						lit_tcvl.SetItemText(iVtTC+iTang,1,m_List.GetItemText(i,1));
						lit_tcvl.SetItemText(iVtTC+iTang,2,m_List.GetItemText(i,2));
						lit_tcvl.SetRowColors(iVtTC+iTang,RGB(255,255,255),RGB(0,0,255));
					}
					
					break;
				}
			}

			if(iClose==0) OnOK();
			else iTang++;
		}
	}
	catch(...){_s(L"#QL12.75: Lỗi áp dụng tiêu chuẩn.");}
}


void Dlg_all_tieuchuan::OnBnClickedB5()
{
	f_SaveWidth();
	OnCancel();
}


void Dlg_all_tieuchuan::OnBnClickedCk1()
{
	CButton *m_ctlCheck = (CButton*) GetDlgItem(ALTC_CK1);
	if(m_ctlCheck->GetCheck() == BST_UNCHECKED) iClose=0;
	else iClose=1;
}


void Dlg_all_tieuchuan::OnNMClickL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	sDownTC=L"",sFullTC=L"";
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	int nItem = pNMListView->iItem;
	if(nItem>=0)
	{
		iTCclick = nItem;
		sDownTC = m_List.GetItemText(iTCclick,1);
		sFullTC = m_List.GetItemText(iTCclick,2);

		m_Matc.SetWindowText(m_List.GetItemText(nItem,1));
		m_Tentc.SetWindowText(m_List.GetItemText(nItem,2));
		m_Loaitc.SetWindowText(m_List.GetItemText(nItem,3));
	}
}


void Dlg_all_tieuchuan::OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	OnBnClickedB3();
}


void Dlg_all_tieuchuan::OnContextMenu(CWnd* /*pWnd*/, CPoint point)
{
	try
	{
		if(iTCclick==-1) return;

		CMenu mnu2;
		mnu2.LoadMenu(IDR_MENU8);
		
		CListCtrl *pClist;
		pClist = reinterpret_cast<CListCtrl *>(GetDlgItem(ALTC_L1));

		CRect rectSubmitButton;
		pClist->GetWindowRect(&rectSubmitButton);

		CMenu *mnuP2 = mnu2.GetSubMenu(0);
		ASSERT(mnuP2);

		if( rectSubmitButton.PtInRect(point) )
			mnuP2->TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON, point.x, point.y, this);
	}
	catch(...){}
}


void Dlg_all_tieuchuan::OnNMRClickL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	sDownTC=L"",sFullTC=L"";
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	iTCclick = pNMListView->iItem;
	if(iTCclick>=0)
	{
		sDownTC = m_List.GetItemText(iTCclick,1);
		sFullTC = m_List.GetItemText(iTCclick,2);
	}
}


CString Dlg_all_tieuchuan::_GetMaTieuchuan()
{
	try
	{
		iTCclick=-1;
		CString szval=L"";
		POSITION pos = m_List.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			for (int i = 0; i < m_List.GetItemCount(); i++)
			if (m_List.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
			{
				iTCclick=i;
				szval = m_List.GetItemText(i,1);
				return szval;
			}
		}		
	}
	catch(...){return L"";}
	return L"";
}



void Dlg_all_tieuchuan::OnTi40039()
{
	sDownTC = _GetMaTieuchuan();
	f_DocFileTC(sDownTC, 1);
}


void Dlg_all_tieuchuan::OnTi40040()
{
	sDownTC = _GetMaTieuchuan();
	f_SearchGoogle(L"",sDownTC,0);
}