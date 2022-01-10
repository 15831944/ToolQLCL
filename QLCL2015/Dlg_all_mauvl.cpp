#include "Dlg_all_mauvl.h"

Dlg_all_mauvl::Dlg_all_mauvl(CWnd* pParent /*=NULL*/)
	: CDialog(Dlg_all_mauvl::IDD, pParent)
{
	
}

void Dlg_all_mauvl::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, ALMAU_T1, m_Txt);
	DDX_Control(pDX, ALMAU_T2, m_MaLm);
	DDX_Control(pDX, ALMAU_T3, m_Pplm);
	DDX_Control(pDX, ALMAU_T4, m_Tslm);
	DDX_Control(pDX, ALMAU_T5, m_Dvi);
	DDX_Control(pDX, ALMAU_L1, m_List);
	DDX_Control(pDX, ALMAU_G1, m_Group);
	DDX_Control(pDX, ALMAU_B3, m_btn[3]);	
}

BEGIN_MESSAGE_MAP(Dlg_all_mauvl, CDialog)
	ON_BN_CLICKED(ALMAU_B1, &Dlg_all_mauvl::OnBnClickedB1)
	ON_BN_CLICKED(ALMAU_B2, &Dlg_all_mauvl::OnBnClickedB2)
	ON_BN_CLICKED(ALMAU_B4, &Dlg_all_mauvl::OnBnClickedB4)
	ON_BN_CLICKED(ALMAU_B3, &Dlg_all_mauvl::OnBnClickedB3)
	ON_BN_CLICKED(ALMAU_B5, &Dlg_all_mauvl::OnBnClickedB5)
	ON_NOTIFY(LVN_KEYDOWN, ALMAU_L1, &Dlg_all_mauvl::OnLvnKeydownL1)
	ON_BN_CLICKED(ALMAU_CK1, &Dlg_all_mauvl::OnBnClickedCk1)
	ON_NOTIFY(NM_CLICK, ALMAU_L1, &Dlg_all_mauvl::OnNMClickL1)
	ON_NOTIFY(NM_DBLCLK, ALMAU_L1, &Dlg_all_mauvl::OnNMDblclkL1)
	ON_WM_SIZE()
	ON_WM_SIZING()	
END_MESSAGE_MAP()

BEGIN_EASYSIZE_MAP(Dlg_all_mauvl)
	EASYSIZE(ALMAU_S1,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(ALMAU_T1,ES_BORDER,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(ALMAU_B1,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)

	EASYSIZE(ALMAU_G1,ES_BORDER,ES_BORDER,ES_BORDER,ES_BORDER,0)
	EASYSIZE(ALMAU_L1,ES_BORDER,ES_BORDER,ES_BORDER,ES_BORDER,0)

	EASYSIZE(ALMAU_B2,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)
	EASYSIZE(ALMAU_T2,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)
	EASYSIZE(ALMAU_T3,ES_BORDER,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(ALMAU_T4,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(ALMAU_T5,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(ALMAU_B4,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)

	EASYSIZE(ALMAU_CK1,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)
	EASYSIZE(ALMAU_B3,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(ALMAU_B5,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
END_EASYSIZE_MAP

BOOL Dlg_all_mauvl::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	iClose = 0,iTang = 0;	
	if(jTMVL==0) m_btn[3].EnableWindow(1);

	f_Load_ToolTip();
	m_Txt.SetWindowText(xl_tukhoa);
	m_Txt.SetCueBanner(L"Nhập từ khóa tìm kiếm (Ví dụ: Đóng và ép cọc,...)");
	m_MaLm.SetCueBanner(L"Mã mẫu VL");
	m_Pplm.SetCueBanner(L"Nhập phương pháp lấy mẫu...");
	m_Tslm.SetCueBanner(L"Tần suất LM");
	m_Dvi.SetCueBanner(L"Đơn vị tính");

	f_StyleList();
	f_Load_DL();

	INIT_EASYSIZE;

	return TRUE; 
}

// Bắt sự kiện click Enter
BOOL Dlg_all_mauvl::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if (pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(ALMAU_T1))
	{
		OnBnClickedB1();
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(ALMAU_T2))
	{
		GotoDlgCtrl(GetDlgItem(ALMAU_T3));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(ALMAU_T3))
	{
		GotoDlgCtrl(GetDlgItem(ALMAU_T4));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(ALMAU_T4))
	{
		GotoDlgCtrl(GetDlgItem(ALMAU_T5));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(ALMAU_T5))
	{
		GotoDlgCtrl(GetDlgItem(ALMAU_T1));
		return TRUE;
	}
	else if (pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(ALMAU_L1))
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

	return FALSE;
}

void Dlg_all_mauvl::OnSize(UINT nType, int cx, int cy) 
{
	CDialog::OnSize(nType, cx, cy);
	UPDATE_EASYSIZE;
}

void Dlg_all_mauvl::OnSizing(UINT fwSide, LPRECT pRect) 
{
	CDialog::OnSizing(fwSide, pRect);
	EASYSIZE_MINSIZE(320,240,fwSide,pRect);
}


void Dlg_all_mauvl::f_Load_ToolTip()
{
	CListCtrl * pls1 = (CListCtrl *)GetDlgItem(ALMAU_L1);
	CButton * btn1 = (CButton *)GetDlgItem(ALMAU_B2);
	CButton * btn2 = (CButton *)GetDlgItem(ALMAU_B4);
	
	m_ToolTips.Create(this);
	m_ToolTips.SetMaxTipWidth(600);
	m_ToolTips.SetDelayTime(900);

	m_ToolTips.AddTool(pls1, L"Kích đúp để áp dụng dữ liệu. Giữ Shift/Ctrl để chọn nhiều");
	m_ToolTips.AddTool(btn1, L"Tạo mới dữ liệu.");
	m_ToolTips.AddTool(btn2, L"Cập nhật dữ liệu chỉnh sửa.");
}


void Dlg_all_mauvl::f_StyleList()
{
	m_List.InsertColumn(0,L"",LVCFMT_LEFT,0);
	m_List.InsertColumn(1,L"Mã MVL",LVCFMT_LEFT,_wL5[0]);
	m_List.InsertColumn(2,L"Phương pháp lấy mẫu vật liệu",LVCFMT_LEFT,_wL5[1]);
	m_List.InsertColumn(3,L"Tần suất lấy mẫu",LVCFMT_LEFT,_wL5[2]);
	m_List.InsertColumn(4,L"Đơn vị tính",LVCFMT_LEFT,_wL5[3]);
	m_List.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);
}


void Dlg_all_mauvl::f_SaveWidth()
{
	for (int i = 0; i < 4; i++) _wL5[i] = m_List.GetColumnWidth(i+1);
}


void Dlg_all_mauvl::f_Tack_tukhoa(CString sKey)
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


void Dlg_all_mauvl::f_Tao_truyvan()
{
	try
	{
		CString val=L"";
		if(iKey==1)
		{
			SqlString.Format(L"SELECT * FROM tbMau "
			L"WHERE (MaLayMau LIKE '%s%s%s' OR Phuongphaplaymau LIKE '%s%s%s' "
			L"OR TansuatLaymau LIKE '%s%s%s' OR Donvi LIKE '%s%s%s');"
			,L"%",pKey[1],L"%",L"%",pKey[1],L"%",L"%",pKey[1],L"%",L"%",pKey[1],L"%");
		}
		else
		{
			SqlString.Format(L"SELECT * FROM tbMau " 
			L"WHERE ((MaLayMau LIKE '%s%s%s' OR Phuongphaplaymau LIKE '%s%s%s' "
			L"OR TansuatLaymau LIKE '%s%s%s' OR Donvi LIKE '%s%s%s')"
			,L"%",pKey[1],L"%",L"%",pKey[1],L"%",L"%",pKey[1],L"%",L"%",pKey[1],L"%");

			for (int i = 2; i <= iKey; i++)
			{
				val.Format(
					L" AND (MaLayMau LIKE '%s%s%s' OR Phuongphaplaymau LIKE '%s%s%s' "
					L"OR TansuatLaymau LIKE '%s%s%s' OR Donvi LIKE '%s%s%s')"
					,L"%",pKey[i],L"%",L"%",pKey[i],L"%",L"%",pKey[i],L"%",L"%",pKey[i],L"%");
				SqlString+=val;
			}
			SqlString+= L");";
		}
	}
	catch(...){}
}


void Dlg_all_mauvl::OnLvnKeydownL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);
	if(pLVKeyDow->wVKey == VK_SPACE) OnBnClickedB3();
	else if(pLVKeyDow->wVKey == VK_DELETE) f_Xoa_mauvl();

	*pResult = 0;
}


void Dlg_all_mauvl::f_Xoa_mauvl()
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
			if(dem==1) val = L"Bạn có chắc chắn sẽ xóa mẫu vật liệu này trong dữ liệu?";
			else val = L"Bạn có chắc chắn sẽ xóa những mẫu vật liệu này trong dữ liệu?";
			
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
						sTvan.Format(L"DELETE FROM tbMau WHERE MaLayMau='%s';",val);
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
		AfxMessageBox(_T("[224] Lỗi xóa dữ liệu mẫu vật liệu")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_all_mauvl::f_Truy_van_DL()
{
	TRY
	{
		if(getPathQLCL()==L"") return;
		if(connectDsn(pathMDB)==-1) return;
		ObjConn.OpenAccessDB(pathMDB,L"",L"");
		CRecordset* Recset = ObjConn.Execute(SqlString);
		iKqua=0;
		while( !Recset->IsEOF() )
		{
			iKqua++;
			if(iKqua>1) break;
			Recset->GetFieldValue(L"MaLayMau",XLXD[1].CDG[0]);
			Recset->GetFieldValue(L"Phuongphaplaymau",XLXD[1].CDG[1]);
			Recset->GetFieldValue(L"TansuatLaymau",XLXD[1].CDG[2]);
			Recset->GetFieldValue(L"Donvi",XLXD[1].CDG[3]);
			Recset->MoveNext();
		}
		Recset->Close();
		ObjConn.CloseDatabase();
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[1] Lỗi tìm kiếm dữ liệu mẫu vật liệu")+e->m_strError);
	}
	END_CATCH;
}


int Dlg_all_mauvl::f_Tim_kiem_MVL()
{
	try
	{
		// Tạo truy vấn tra cứu công việc		
		if(xl_tukhoa==L"") SqlString = L"SELECT * FROM tbMau ORDER BY MaLayMau ASC;";
		else
		{
			f_Tack_tukhoa(xl_tukhoa);
			f_Tao_truyvan();
		}		

		// Truy vấn dữ liệu
		f_Truy_van_DL();
		if(iKqua==1)
		{
			lit_ppvl.SetItemText(CLRow,1,XLXD[1].CDG[0]);
			lit_ppvl.SetItemText(CLRow,2,XLXD[1].CDG[1]);
			lit_ppvl.SetItemText(CLRow,3,XLXD[1].CDG[2]);
			lit_ppvl.SetItemText(CLRow,4,XLXD[1].CDG[3]);
			lit_ppvl.SetRowColors(CLRow,RGB(255,255,255),RGB(0,0,0));
		}
		else if(iKqua>1)
		{
			// Có nhiều kết quả -> hiển thị hộp thoại
			AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
			DoModal();

			return 1;
		}
		else lit_ppvl.SetRowColors(CLRow,RGB(255,255,255),RGB(255,0,0));

		return 0;
	}
	catch(...){_s(L"#TC237: Lỗi tìm kiếm mẫu vật liệu.");}

	return 0;
}


void Dlg_all_mauvl::f_Load_DL()
{
	TRY
	{
		// Kết nối SQL
		ObjConn.OpenAccessDB(pathMDB,L"",L"");
		CRecordset* Recset = ObjConn.Execute(SqlString);
		int dem=0;
		CString sGan[5];
		while( !Recset->IsEOF() )
		{
			Recset->GetFieldValue(L"MaLayMau",sGan[0]);
			Recset->GetFieldValue(L"Phuongphaplaymau",sGan[1]);
			Recset->GetFieldValue(L"TansuatLaymau",sGan[2]);
			Recset->GetFieldValue(L"Donvi",sGan[3]);

			m_List.InsertItem(dem,L"",0);
			m_List.SetItemText(dem,1,sGan[0]);
			m_List.SetItemText(dem,2,sGan[1]);
			m_List.SetItemText(dem,3,sGan[2]);
			m_List.SetItemText(dem,4,sGan[3]);
			dem++;

			Recset->MoveNext();
		}
		Recset->Close();
		ObjConn.CloseDatabase();

		// Kết quả tìm kiếm
		sGan[0].Format(L"Kết quả tìm kiếm (%d mẫu vật liệu)",dem);
		m_Group.SetWindowText(sGan[0]);
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[82] Lỗi tìm kiếm dữ liệu mẫu vật liệu")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_all_mauvl::OnBnClickedB1()
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
		if(val==L"") SqlString = L"SELECT * FROM tbMau ORDER BY MaLayMau ASC;";
		else
		{
			f_Tack_tukhoa(val);
			f_Tao_truyvan();
		}

		f_Load_DL();
	}
	catch(...){}
}


void Dlg_all_mauvl::OnBnClickedB2()
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
		jNew.Format(L"VL.%d%d%d",_nRand,rand()%10,rand()%10);
		sTvan.Format(L"SELECT COUNT(*) AS iDem FROM tbMau WHERE MaLayMau = '%s';",jNew);
		
		ObjConn.OpenAccessDB(pathMDB,L"",L"");		
		CRecordset* Recset = ObjConn.Execute(sTvan);
		Recset->GetFieldValue(L"iDem",sDem);
		Recset->Close();

		while (_wtoi(sDem)>0)
		{
			sDem=L"";
			jNew.Format(L"VLNW.%d%d%d",_nRand,rand()%10,rand()%10);
			sTvan.Format(L"SELECT COUNT(*) AS iDem FROM tbMau WHERE MaLayMau = '%s';",jNew);
			Recset = ObjConn.Execute(sTvan);
			Recset->GetFieldValue(L"iDem",sDem);
			Recset->Close();
		}

		ObjConn.CloseDatabase();

		m_MaLm.SetWindowText(jNew);
		m_Pplm.SetWindowText(L"");
		m_Tslm.SetWindowText(L"");
		m_Dvi.SetWindowText(L"");
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[41] Lỗi thêm mới mẫu vật liệu")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_all_mauvl::OnBnClickedB4()
{
	TRY
	{
		CString s1=L"",s2=L"",s3=L"",s4=L"";
		m_MaLm.GetWindowTextW(s1); s1.Trim();
		if(s1==L"")
		{
			GotoDlgCtrl(GetDlgItem(ALMAU_T2));
			return;
		}

		// Tên & loại tiêu chuẩn
		m_Pplm.GetWindowTextW(s2);
		m_Tslm.GetWindowTextW(s3);
		m_Dvi.GetWindowTextW(s4);
		s2.Replace(L"'",L""); s2.Trim();
		s3.Replace(L"'",L""); s3.Trim();
		s4.Replace(L"'",L""); s4.Trim();

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
		sTvan.Format(L"SELECT COUNT(*) AS iDem FROM tbMau WHERE MaLayMau = '%s';",s1);
		
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
			sTvan.Format(L"UPDATE tbMau "
				L"SET Phuongphaplaymau='%s', TansuatLaymau='%s', Donvi='%s' WHERE MaLayMau='%s';",s2,s3,s4,s1);
			ObjConn.ExecuteRB(sTvan);
		}
		else
		{
			sTvan.Format(L"INSERT INTO tbMau (MaLayMau,Phuongphaplaymau,TansuatLaymau,Donvi) "
				L"VALUES ('%s','%s','%s','%s');",s1,s2,s3,s4);
			ObjConn.ExecuteRB(sTvan);

			// Thêm mới hoàn toàn --> Chèn dòng vào list
			m_List.InsertItem(nItem,L"",0);
		}

		// Đổ dữ liệu vào list
		m_List.SetItemText(nItem,0,L"");
		m_List.SetItemText(nItem,1,s1);
		m_List.SetItemText(nItem,2,s2);
		m_List.SetItemText(nItem,3,s3);
		m_List.SetItemText(nItem,4,s4);

		ObjConn.CloseDatabase();
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("[21] Lỗi thêm mới dữ liệu")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_all_mauvl::OnBnClickedB3()
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
					if(iTang>0) lit_ppvl.InsertItem(CLRow+iTang,L"",0);

					lit_ppvl.SetItemText(CLRow+iTang,0,L"1");
					for (int j = 1; j <= 4; j++)
						lit_ppvl.SetItemText(CLRow+iTang,j,m_List.GetItemText(i,j));

					lit_ppvl.SetRowColors(CLRow+iTang,RGB(255,255,255),RGB(0,0,255));
					break;
				}
			}

			if(iClose==0) OnOK();
			else iTang++;
		}
	}
	catch(...){_s(L"#QL12.75: Lỗi áp dụng mẫu vật liệu.");}
}


void Dlg_all_mauvl::OnBnClickedB5()
{
	f_SaveWidth();
	OnCancel();
}


void Dlg_all_mauvl::OnBnClickedCk1()
{
	CButton *m_ctlCheck = (CButton*) GetDlgItem(ALMAU_CK1);
	if(m_ctlCheck->GetCheck() == BST_UNCHECKED) iClose=0;
	else iClose=1;
}


void Dlg_all_mauvl::OnNMClickL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	int nItem = pNMListView->iItem;
	if(nItem>=0)
	{
		m_MaLm.SetWindowText(m_List.GetItemText(nItem,1));
		m_Pplm.SetWindowText(m_List.GetItemText(nItem,2));
		m_Tslm.SetWindowText(m_List.GetItemText(nItem,3));
		m_Dvi.SetWindowText(m_List.GetItemText(nItem,4));
	}
}


void Dlg_all_mauvl::OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	OnBnClickedB3();
}
