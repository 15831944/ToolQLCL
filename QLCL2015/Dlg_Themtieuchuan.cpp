#include "stdafx.h"
#include "Dlg_Themtieuchuan.h"
#include "Dlg_all_tieuchuan.h"

// Dlg_Themtieuchuan dialog
IMPLEMENT_DYNAMIC(Dlg_Themtieuchuan, CDialogEx)


Dlg_Themtieuchuan::Dlg_Themtieuchuan(CWnd* pParent /*=NULL*/)
	: CDialogEx(Dlg_Themtieuchuan::IDD, pParent)
{

}


Dlg_Themtieuchuan::~Dlg_Themtieuchuan()
{
}

void Dlg_Themtieuchuan::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, ADNTC_T1, m_txtTieuchuan);
	DDX_Control(pDX, ADNTC_T2, m_txtCongviec);

	DDX_Control(pDX, ADNTC_L1, m_listTC);
	DDX_Control(pDX, ADNTC_L2, m_listThem);
	DDX_Control(pDX, ADNTC_L3, m_listCongviec);
	DDX_Control(pDX, ADNTC_L4, m_listViewTC);

	DDX_Control(pDX, ADNTC_B4, m_btnKqua);
	DDX_Control(pDX, ADNTC_CK1, m_check);
	DDX_Control(pDX, ADNTC_CK2, m_cksltc);
	
	DDX_Control(pDX, ADNTC_CB1, m_cbbox);
	DDX_Control(pDX, ADNTC_STTITLE, m_title);
	DDX_Control(pDX, ADNTC_PRS1, m_progress);

	DDX_Control(pDX, ADNTC_PRE, m_btnPre);
	DDX_Control(pDX, ADNTC_NXT, m_btnNxt);
	DDX_Control(pDX, ADNTC_CBBFL, m_cbbFile);
}


BEGIN_MESSAGE_MAP(Dlg_Themtieuchuan, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_SETCURSOR()
	ON_WM_CONTEXTMENU()
	ON_BN_CLICKED(ADNTC_B1, &Dlg_Themtieuchuan::OnBnClickedB1)
	ON_BN_CLICKED(ADNTC_B2, &Dlg_Themtieuchuan::OnBnClickedB2)
	ON_BN_CLICKED(ADNTC_B3, &Dlg_Themtieuchuan::OnBnClickedB3)
	ON_BN_CLICKED(ADNTC_CK1, &Dlg_Themtieuchuan::OnBnClickedCk1)
	ON_BN_CLICKED(ADNTC_B4, &Dlg_Themtieuchuan::OnBnClickedB4)
	ON_NOTIFY(NM_CLICK, ADNTC_L2, &Dlg_Themtieuchuan::OnNMClickL2)
	ON_NOTIFY(NM_RCLICK, ADNTC_L2, &Dlg_Themtieuchuan::OnNMRClickL2)
	ON_BN_CLICKED(ADNTC_B7, &Dlg_Themtieuchuan::OnBnClickedB7)
	ON_BN_CLICKED(ADNTC_B8, &Dlg_Themtieuchuan::OnBnClickedB8)
	ON_NOTIFY(NM_CLICK, ADNTC_L3, &Dlg_Themtieuchuan::OnNMClickL3)
	ON_NOTIFY(NM_CLICK, ADNTC_L1, &Dlg_Themtieuchuan::OnNMClickL1)
	ON_EN_SETFOCUS(ADNTC_T1, &Dlg_Themtieuchuan::OnEnSetfocusT1)
	ON_NOTIFY(NM_CLICK, ADNTC_L4, &Dlg_Themtieuchuan::OnNMClickL4)	
	ON_NOTIFY(NM_RCLICK, ADNTC_L1, &Dlg_Themtieuchuan::OnNMRClickL1)
	ON_NOTIFY(NM_RCLICK, ADNTC_L3, &Dlg_Themtieuchuan::OnNMRClickL3)
	ON_NOTIFY(NM_RCLICK, ADNTC_L4, &Dlg_Themtieuchuan::OnNMRClickL4)
	ON_COMMAND(ID_TH40029, &Dlg_Themtieuchuan::OnTh40029)
	ON_COMMAND(ID_TH40030, &Dlg_Themtieuchuan::OnTh40030)
	ON_COMMAND(ID_TH40031, &Dlg_Themtieuchuan::OnTh40031)
	ON_COMMAND(ID_TH40032, &Dlg_Themtieuchuan::OnTh40032)
	ON_COMMAND(ID_TH40033, &Dlg_Themtieuchuan::OnTh40033)
	ON_COMMAND(ID_TH40034, &Dlg_Themtieuchuan::OnTh40034)
	ON_BN_CLICKED(ADNTC_B5, &Dlg_Themtieuchuan::OnBnClickedB5)
	ON_BN_CLICKED(ADNTC_B6, &Dlg_Themtieuchuan::OnBnClickedB6)
	ON_BN_CLICKED(ADNTC_CK2, &Dlg_Themtieuchuan::OnBnClickedCk2)
	ON_COMMAND(ID_TH40041, &Dlg_Themtieuchuan::OnTh40041)
	ON_COMMAND(ID_TH40042, &Dlg_Themtieuchuan::OnTh40042)
	ON_NOTIFY(NM_DBLCLK, ADNTC_L1, &Dlg_Themtieuchuan::OnNMDblclkL1)
	ON_CBN_SELCHANGE(ADNTC_CB1, &Dlg_Themtieuchuan::OnCbnSelchangeCb1)
	ON_BN_CLICKED(ADNTC_PRE, &Dlg_Themtieuchuan::OnBnClickedPre)
	ON_BN_CLICKED(ADNTC_NXT, &Dlg_Themtieuchuan::OnBnClickedNxt)
	ON_CBN_SELCHANGE(ADNTC_CBBFL, &Dlg_Themtieuchuan::OnCbnSelchangeCbbfl)
END_MESSAGE_MAP()


// Dlg_Themtieuchuan message handlers
BOOL Dlg_Themtieuchuan::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	iTCclick=-1;
	sDownTC=L"",sFullTC=L"";

	szlMaTC=L"";
	nItem=-1,nListClick=1;
	tslkq=0,lanshow=1,ibuocnhay=200;
	m_txtTieuchuan.SetCueBanner(L"Nhập từ khóa tìm kiếm tiêu chuẩn");
	m_txtCongviec.SetCueBanner(L"Nhập từ khóa tìm kiếm công việc");

	m_cbbox.AddString(L"Công việc");
	m_cbbox.AddString(L"Vật liệu");
	if(iCbbThemtc!=1) m_cbbox.SetCurSel(iCbbThemtc);

	m_progress.AlignText(DT_CENTER);
	m_progress.SetBarColor(GGREEN);
	m_progress.SetBarBkColor(WHITE);
	m_progress.SetTextColor(BLACK);
	m_progress.SetTextBkColor(WHITE);

	m_title.SubclassDlgItem(ADNTC_STTITLE,this);
	m_title.SetTextColor(BLUE);

	_SetStyleList();

	getPathQLCL();
	f_Danh_sachfile();
	
	if(connectDsn(pathMDB)==-1) return TRUE;
	_TruyvanSQL();

	//if(sFiledulieu[1].Find(L"\\")>0) sFullPath = sFiledulieu[1];
	//else sFullPath.Format(L"%s\\%s",sDDanfile,sFiledulieu[1]);

	sFullPath = sFpathdulieu[1];
	if(connectDsn(sFullPath)==-1) return TRUE;

	this->SetWindowTextW(L"Thêm nhanh nhóm tiêu chuẩn | " + sFullPath);
	_LoadlistFile();

	return TRUE; 
}


BOOL Dlg_Themtieuchuan::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if(iMes==WM_KEYDOWN)
	{
		if(iWPar==VK_RETURN)
		{
			if(pWndWithFocus==GetDlgItem(ADNTC_T1))
			{
				_TruyvanSQL();
				return TRUE;
			}
			else if(pWndWithFocus==GetDlgItem(ADNTC_T2))
			{
				_TruyvanSQLCongviec();
				return TRUE;
			}
		}
		else if(iWPar==VK_TAB)
		{
			if(pWndWithFocus==GetDlgItem(ADNTC_T1))
			{
				GotoDlgCtrl(GetDlgItem(ADNTC_T2));
				return TRUE;
			}
			else if(pWndWithFocus==GetDlgItem(ADNTC_T2))
			{
				GotoDlgCtrl(GetDlgItem(ADNTC_T1));
				return TRUE;
			}
		}
		else if(iWPar==VK_ESCAPE)
		{
			_OnExit();
			return TRUE;
		}
	}
	else if(iMes == WM_CHAR)
	{
		TCHAR chr = static_cast<TCHAR>(pMsg->wParam);
		if(pWndWithFocus == GetDlgItem(ADNTC_L1) 
			|| pWndWithFocus == GetDlgItem(ADNTC_L2) 
			|| pWndWithFocus == GetDlgItem(ADNTC_L4))
		{
			switch (chr)
			{
				case 0x04:
				{
					OnTh40041();
					break;
				}
				case 0x07:
				{
					OnTh40042();
					break;
				}

				default:
					break;
			}
		}
	}

	return FALSE;
}


void Dlg_Themtieuchuan::OnSysCommand(UINT nID, LPARAM lParam)
{
	if(nID == SC_CLOSE) _OnExit();
	else CDialog::OnSysCommand(nID, lParam);
}


BOOL Dlg_Themtieuchuan::OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message)
{
	if (pWnd == GetDlgItem(ADNTC_B1) || pWnd == GetDlgItem(ADNTC_B2) 
		|| pWnd == GetDlgItem(ADNTC_B3) || pWnd == GetDlgItem(ADNTC_B4) 
		|| pWnd == GetDlgItem(ADNTC_B5) || pWnd == GetDlgItem(ADNTC_B6) 
		|| pWnd == GetDlgItem(ADNTC_B7) || pWnd == GetDlgItem(ADNTC_B8))
	{
		::SetCursor(AfxGetApp()->LoadStandardCursor(IDC_HAND));
		return TRUE;
	} 
	return CDialog::OnSetCursor(pWnd, nHitTest, message);
}


void Dlg_Themtieuchuan::f_Danh_sachfile()
{
	try
	{
		CString spath = GIOR(iRowDLGoc,iColumnDLGoc,prTS,L"");
		CString szfolder = GIOR(iRowDLGoc+1,iColumnDLGoc,prTS,L"");
		if(spath.Right(1)!=L"\\") spath+=L"\\";
		sDDanfile = spath + szfolder;

		slfileDL=0;
		CString szval=L"", scheck=L"";
		for (int i = iRowDLGoc+1; i < iRowDLGoc+100; i++)
		{
			szval = GIOR(i,iColumnDLGoc+2,prTS,L"");
			if(szval==L"") break;

			scheck = GIOR(i,iColumnDLGoc+3,prTS,L"");
			if(_wtoi(scheck)==1)
			{
				if(szval.Find(L"\\")==-1) scheck.Format(L"%s\\%s",sDDanfile,szval);
				else
				{
					scheck = szval;
					szval = _TackTenFile(szval,1);
				}
			
				if(_FileExists(scheck,0)==1)
				{
					slfileDL++;
					sFiledulieu[slfileDL] = szval;
					sFpathdulieu[slfileDL] = scheck;
				}
			}		
		}

		if(slfileDL==0)
		{
			_s(L"Không tồn tại tệp tin dữ liệu. Vui lòng kiểm tra lại đường dẫn.");
			return;
		}
	}
	catch(...){}
}


void Dlg_Themtieuchuan::_LoadlistFile()
{
	try
	{
		int ivt=0,ilen=0;
		CString szval=L"";
		for (int i = 1; i <= slfileDL; i++)
		{
			/*if(sFiledulieu[i].Find(L"\\")==-1) szval = sFiledulieu[i];
			else
			{
				ilen = sFiledulieu[i].GetLength();
				for (int j = ilen-1; j > 0; j--)
				{
					if(sFiledulieu[i].Mid(j,1)==L"\\")
					{
						szval = sFiledulieu[i].Right(ilen-j-1);
						break;
					}
				}
			}*/

			m_cbbFile.AddString(sFiledulieu[i]);
			if(sFullPath==sFpathdulieu[i]) ivt=i-1;
		}

		m_cbbFile.SetCurSel(ivt);

		if(slfileDL==1)
		{
			m_btnPre.EnableWindow(0);
			m_btnNxt.EnableWindow(0);
			m_cbbFile.EnableWindow(0);
		}
	}
	catch(...){}
}

void Dlg_Themtieuchuan::_OnExit()
{
	CDialog::OnCancel();
}


void Dlg_Themtieuchuan::_TruyvanSQL()
{
	try
	{
		CString szval=L"";
		m_txtTieuchuan.GetWindowTextW(szval);
		szval.Replace(L"'",L"");
		szval.Trim();

		if(szval==L"") SqlString = L"SELECT * FROM tbTentieuchuan WHERE MaTC <> '' ORDER BY MaTC ASC;";
		else
		{
			_TackTukhoa(szval,L" ",L"+");

			if(iKey==1)
			{
				SqlString.Format(
					L"SELECT * FROM tbTentieuchuan WHERE (MaTC LIKE '%s%s%s' OR TenTC LIKE '%s%s%s')",
					L"%",pKey[1],L"%",L"%",pKey[1],L"%");
			}
			else
			{
				SqlString.Format(
					L"SELECT * FROM tbTentieuchuan WHERE ((MaTC LIKE '%s%s%s' OR TenTC LIKE '%s%s%s')",
					L"%",pKey[1],L"%",L"%",pKey[1],L"%");

				szval=L"";
				for (int i = 2; i <= iKey; i++)
				{
					szval.Format(
						L" AND (MaTC LIKE '%s%s%s' OR TenTC LIKE '%s%s%s')"
						,L"%",pKey[i],L"%",L"%",pKey[i],L"%");
					SqlString+=szval;
				}

				SqlString+= L")";
			}

			SqlString+= L" ORDER BY MaTC ASC;";
		}

		_LoadAllTC();
	}
	catch(...){}
}


void Dlg_Themtieuchuan::_TruyvanSQLCongviec()
{
	try
	{
		CString sTbl=L"tbDonGia";
		CString sFd[2]={L"MHDG",L"TCV"};
		if(iCbbThemtc==1)
		{
			sTbl=L"tbTuDienVatTu";
			sFd[0]=L"MSVT";
			sFd[1]=L"TVT";
		}

		CString szval=L"";
		m_txtCongviec.GetWindowTextW(szval);
		szval.Replace(L"'",L"");
		szval.Trim();

		if(szval==L"")
		{
			//SqlString.Format(L"SELECT * FROM %s WHERE %s <> '' ORDER BY %s ASC;",sTbl,sFd[0],sFd[0]);
			SqlString.Format(L"SELECT * FROM %s;",sTbl);
		}
		else
		{
			_TackTukhoa(szval,L" ",L"+");

			if(iKey==1)
			{
				SqlString.Format(
					L"SELECT * FROM %s WHERE (%s LIKE '%s%s%s' OR %s LIKE '%s%s%s')",
					sTbl,sFd[0],L"%",pKey[1],L"%",sFd[1],L"%",pKey[1],L"%");
			}
			else
			{
				SqlString.Format(
					L"SELECT * FROM %s WHERE ((%s LIKE '%s%s%s' OR %s LIKE '%s%s%s')",
					sTbl,sFd[0],L"%",pKey[1],L"%",sFd[1],L"%",pKey[1],L"%");

				szval=L"";
				for (int i = 2; i <= iKey; i++)
				{
					szval.Format(
						L" AND (%s LIKE '%s%s%s' OR %s LIKE '%s%s%s')"
						,sFd[0],L"%",pKey[i],L"%",sFd[1],L"%",pKey[i],L"%");
					SqlString+=szval;
				}

				SqlString+= L")";
			}
		}

		_LoadAllCongviec();
	}
	catch(...){}
}


void Dlg_Themtieuchuan::_SetStyleList()
{
	m_listTC.InsertColumn(0,L"",LVCFMT_LEFT,22);
	m_listTC.InsertColumn(1,L"Mã TC",LVCFMT_LEFT,0);
	m_listTC.InsertColumn(2,L"Tên tiêu chuẩn",LVCFMT_LEFT,500);
	m_listTC.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP | LVS_EX_CHECKBOXES);

	m_listThem.InsertColumn(0,L"",LVCFMT_LEFT,22);
	m_listThem.InsertColumn(1,L"Mã TC",LVCFMT_LEFT,0);
	m_listThem.InsertColumn(2,L"Tên tiêu chuẩn",LVCFMT_LEFT,500);
	m_listThem.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP | LVS_EX_CHECKBOXES);

	m_listCongviec.InsertColumn(0,L"",LVCFMT_LEFT,22);
	m_listCongviec.InsertColumn(1,L"Mã số",LVCFMT_LEFT,80);
	m_listCongviec.InsertColumn(2,L"Tên công việc/ vật liệu",LVCFMT_LEFT,500);
	m_listCongviec.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP | LVS_EX_CHECKBOXES);
	m_listCongviec.SetColumnColors(0,WHITE,BLUE);
	m_listCongviec.SetColumnSorting(0, CListCtrlEx::Auto, CListCtrlEx::StringNumber);

	m_listViewTC.InsertColumn(0,L"",LVCFMT_LEFT,22);
	m_listViewTC.InsertColumn(1,L"Mã TC",LVCFMT_LEFT,0);	
	m_listViewTC.InsertColumn(2,L"Tên tiêu chuẩn",LVCFMT_LEFT,500);	
	m_listViewTC.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP | LVS_EX_CHECKBOXES);
}


void Dlg_Themtieuchuan::_LoadAllTC()
{
	TRY
	{
		if((int)m_listTC.GetItemCount()>0)
		{
			m_listTC.DeleteAllItems();
			ASSERT(m_listTC.GetItemCount() == 0);
		}		

		ObjConn2.OpenAccessDB(pathMDB, L"", L"");
		CRecordset* Recset = ObjConn2.Execute(SqlString);
		
		long cong = 0;
		CString szval[2];
		while( !Recset->IsEOF() )
		{
			Recset->GetFieldValue(L"MaTC",szval[0]);
			Recset->GetFieldValue(L"TenTC",szval[1]);

			m_listTC.InsertItem(cong,L"",0);
			m_listTC.SetItemText(cong,1,szval[0]);
			m_listTC.SetItemText(cong,2,szval[1]);

			cong++;
			Recset->MoveNext();
		}
		Recset->Close();
		ObjConn2.CloseDatabase();

		if(cong>0) GotoDlgCtrl(GetDlgItem(ADNTC_L1));
		else GotoDlgCtrl(GetDlgItem(ADNTC_T1));
	}
	CATCH(CDBException,e){}
	END_CATCH;
}


void Dlg_Themtieuchuan::_LoadAllCongviec()
{
	TRY
	{
		CString sTbl=L"tbDonGia";
		CString sFd[2]={L"MHDG",L"TCV"};
		if(iCbbThemtc==1)
		{
			sTbl=L"tbTuDienVatTu";
			sFd[0]=L"MSVT";
			sFd[1]=L"TVT";
		}

		m_check.SetCheck(0);
		if((int)m_listCongviec.GetItemCount()>0)
		{
			m_listCongviec.DeleteAllItems();
			ASSERT(m_listCongviec.GetItemCount() == 0);
		}

		ObjConn.OpenAccessDB(sFullPath, L"", L"");
		CRecordset* Recset = ObjConn.Execute(SqlString);
		
		tslkq=0;
		CRecordset* RCS2;
		CString struyvan=L"";
		while( !Recset->IsEOF() )
		{
			tslkq++;
			Recset->GetFieldValue(sFd[0],XLXD[tslkq].CDG[0]);			
			Recset->GetFieldValue(sFd[1],XLXD[tslkq].CDG[1]);

			Recset->MoveNext();
		}
		Recset->Close();
		ObjConn.CloseDatabase();

		lanshow=1;
		int slload=ibuocnhay;
		if(tslkq<=slload)
		{
			slload=tslkq;
			m_btnKqua.EnableWindow(0);
		}
		else m_btnKqua.EnableWindow(1);

		for (int i = 0; i < slload; i++)
		{
			m_listCongviec.InsertItem(i,L"",0);
			m_listCongviec.SetItemText(i,1,XLXD[i+1].CDG[0]);
			m_listCongviec.SetItemText(i,2,XLXD[i+1].CDG[1]);
		}

		// Kết quả tìm kiếm
		CString szval=L"";
		szval.Format(L"Hiển thị thêm (%d / %d kết quả)",slload,tslkq);
		m_btnKqua.SetWindowText(szval);

		if(slload>0) GotoDlgCtrl(GetDlgItem(ADNTC_L3));
		else GotoDlgCtrl(GetDlgItem(ADNTC_T2));
	}
	CATCH(CDBException,e){}
	END_CATCH;
}


void Dlg_Themtieuchuan::OnBnClickedB1()
{
	int count = m_listTC.GetItemCount();
	if(count==0) return;

	if(nItem==-1) nItem = m_listThem.GetItemCount();	

	CString szMa=L"",szTen=L"";
	for (int i = count-1; i >= 0; i--)
	{
		if((int)m_listTC.GetCheck(i)==1)
		{
			szMa = m_listTC.GetItemText(i,1);
			szTen = m_listTC.GetItemText(i,2);

			if(_CheckItemTC(szMa)==0)
			{
				m_listThem.InsertItem(nItem,L"",0);
				m_listThem.SetItemText(nItem,1,szMa);
				m_listThem.SetItemText(nItem,2,szTen);
				m_listThem.SetCheck(nItem,1);
			}			
		}
	}
}


int Dlg_Themtieuchuan::_CheckItemTC(CString szCheck)
{
	int count = m_listThem.GetItemCount();
	if(count==0) return 0;

	CString szval=L"";
	for (int i = 0; i < count; i++)
	{
		szval = m_listThem.GetItemText(i,1);
		if(szCheck==szval) return 1;
	}

	return 0;
}


void Dlg_Themtieuchuan::OnBnClickedB2()
{
	int count = m_listThem.GetItemCount();
	if(count==0) return;

	for (int i = count-1; i >= 0; i--)
		if((int)m_listThem.GetCheck(i)==1) m_listThem.DeleteItem(i);
}


void Dlg_Themtieuchuan::OnBnClickedB3()
{
	try
	{
		CString sTbl=L"tbMaCV_Tieuchuan";
		CString sFd=L"MaCV";
		if(iCbbThemtc==1)
		{
			sTbl=L"tbMaVL_Tieuchuan";
			sFd=L"MaVL";
		}

		int count = m_listThem.GetItemCount();
		if(count==0) return;

		int sltc=0;
		CString sMaTC[200];
		for (int i = 0; i < count ; i++)
		{
			if((int)m_listThem.GetCheck(i)==1)
			{
				sMaTC[sltc] = m_listThem.GetItemText(i,1);
				sltc++;
				if(sltc==199) break;
			}
		}

		if(sltc==0)
		{
			_s(L"Vui lòng tích chọn tiêu chuẩn cần thêm.");
			return;
		}

		count = m_listCongviec.GetItemCount();
		if(count==0) return;

		int dem=0;
		for (int i = 0; i < count ; i++)
			if((int)m_listCongviec.GetCheck(i)==1) dem++;

		if(dem==0)
		{
			_s(L"Vui lòng tích chọn công việc sẽ thêm tiêu chuẩn.");
			return;
		}

		if(_yn(L"Bạn có chắc thêm các tiêu chuẩn vào các công việc được chọn?",0)!=6) return;

		ObjConn2.OpenAccessDB(pathMDB, L"", L"");
		CRecordset* Recset;

		m_title.ShowWindow(0);
		m_progress.ShowWindow(1);
		m_progress.SetRange(0,dem);
		m_progress.SetWindowText(L"Đang gán tiêu chuẩn: ");
		m_progress.SetPos(0);

		int cong=0,gt=0;
		CString szval=L"",sdem=L"";
		for (int i = 0; i < count; i++)
		{
			if((int)m_listCongviec.GetCheck(i)==1)
			{
				cong++;
				m_progress.SetPos(cong);
				szval = m_listCongviec.GetItemText(i,1);
				SqlString.Format(L"SELECT COUNT(*) AS iDem FROM %s WHERE %s LIKE '%s';",sTbl,sFd,szval);
				Recset = ObjConn2.Execute(SqlString);
				Recset->GetFieldValue(L"iDem",sdem);
				Recset->Close();
				gt = _wtoi(sdem);				

				for (int j = 0; j < sltc; j++)
				{
					// Kiểm tra tiêu chuẩn đã tồn tại trong công việc đó chưa?
					SqlString.Format(
						L"SELECT COUNT(*) AS iDem FROM %s "
						L"WHERE (%s LIKE '%s' AND MaTC LIKE '%s');",sTbl,sFd,szval,sMaTC[j]);

					Recset = ObjConn2.Execute(SqlString);
					Recset->GetFieldValue(L"iDem",sdem);
					Recset->Close();

					if(_wtoi(sdem)==0)
					{
						gt++;
						if(gt<=9) sdem.Format(L"0%d",gt);
						else sdem.Format(L"%d",gt);

						SqlString.Format(L"INSERT INTO %s (ID,%s,MaTC) "
							L"VALUES ('%s','%s','%s');",sTbl,sFd,sdem,szval,sMaTC[j]);
						ObjConn2.ExecuteRB(SqlString);
					}
				}
			}
		}		
		
		m_progress.ShowWindow(0);
		m_title.ShowWindow(1);
		m_title.SetWindowText(L"Hoàn tất thêm mới tiêu chuẩn!");
		ObjConn2.CloseDatabase();
	}
	catch(...){}
}


void Dlg_Themtieuchuan::OnBnClickedCk1()
{
	int count = m_listCongviec.GetItemCount();
	if(count==0) return;

	int check = m_check.GetCheck();
	for (int i = 0; i < count; i++) m_listCongviec.SetCheck(i,check);
}


void Dlg_Themtieuchuan::OnBnClickedB4()
{
	try
	{
		int nCount = m_listCongviec.GetItemCount();

		lanshow++;
		int slload=ibuocnhay*lanshow;
		if(tslkq<=slload)
		{
			slload=tslkq;
			m_btnKqua.EnableWindow(0);
		}
		else m_btnKqua.EnableWindow(1);

		int check = m_check.GetCheck();
		for (int i = ibuocnhay*(lanshow-1); i < slload; i++)
		{
			m_listCongviec.InsertItem(i,L"",0);
			m_listCongviec.SetItemText(i,1,XLXD[i+1].CDG[0]);
			m_listCongviec.SetItemText(i,2,XLXD[i+1].CDG[1]);
			m_listCongviec.SetCheck(i,check);
		}

		RECT r;
		m_listCongviec.GetItemRect(0, &r, LVIR_BOUNDS);
		int a = r.bottom - r.top;

		CRect rectCtrl;
		m_listCongviec.GetWindowRect(&rectCtrl);
		int b = rectCtrl.Height();
	
		int gt=(int)(b/a)-5;
		m_listCongviec.EnsureVisible(nCount+gt, FALSE);

		CString szval=L"";
		szval.Format(L"Hiển thị thêm (%d / %d kết quả)",slload,tslkq);
		m_btnKqua.SetWindowText(szval);
	}
	catch(...){}
}


void Dlg_Themtieuchuan::OnNMClickL2(NMHDR *pNMHDR, LRESULT *pResult)
{
	nListClick=2;
	m_title.SetWindowText(L"Bạn có thể thay đổi vị trí thứ tự các tiêu chuẩn");
	
	sDownTC=L"",sFullTC=L"";
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	nItem = pNMListView->iItem;

	if(nItem>=0)
	{
		iTCclick = nItem;
		sDownTC = m_listThem.GetItemText(iTCclick,1);
		sFullTC = m_listThem.GetItemText(iTCclick,2);
	}

	*pResult = 0;
}


void Dlg_Themtieuchuan::OnNMRClickL2(NMHDR *pNMHDR, LRESULT *pResult)
{
	nListClick=2;
	sDownTC=L"",sFullTC=L"";
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	nItem = pNMListView->iItem;
	if(nItem>=0)
	{
		iTCclick = nItem;
		sDownTC = m_listThem.GetItemText(iTCclick,1);
		sFullTC = m_listThem.GetItemText(iTCclick,2);
	}
}


void Dlg_Themtieuchuan::OnBnClickedB7()
{
	_TruyvanSQL();
}


void Dlg_Themtieuchuan::OnBnClickedB8()
{
	_TruyvanSQLCongviec();
}


void Dlg_Themtieuchuan::OnNMClickL3(NMHDR *pNMHDR, LRESULT *pResult)
{
	iTCclick=-1;
	sDownTC=L"",sFullTC=L"";

	nListClick=3;
	m_title.SetWindowText(L"Sử dụng phím Shift hoặc Ctrl để chọn nhiều, sau đó nhấn chuột phải để tích chọn");

	if((int)m_listViewTC.GetItemCount()>0)
	{
		m_listViewTC.DeleteAllItems();
		ASSERT(m_listViewTC.GetItemCount() == 0);
	}

	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	int nRow = pNMListView->iItem;
	if(nRow>=0) _ViewTC(nRow);
	*pResult = 0;
}


void Dlg_Themtieuchuan::_ViewTC(int irow)
{
	try
	{
		CString sTbl=L"tbMaCV_Tieuchuan";
		CString sFd=L"MaCV";
		if(iCbbThemtc==1)
		{
			sTbl=L"tbMaVL_Tieuchuan";
			sFd=L"MaVL";
		}

		szlMaTC = m_listCongviec.GetItemText(irow,1);
		if(szlMaTC==L"") return;

		SqlString.Format(L"SELECT * FROM %s WHERE %s LIKE '%s' ORDER BY ID ASC;",sTbl,sFd,szlMaTC);

		ObjConn2.OpenAccessDB(pathMDB, L"", L"");
		CRecordset* Recset = ObjConn2.Execute(SqlString);
		
		int dem=0;
		CString szval=L"";
		CRecordset* RCS;
		CString gt[3]={L"",L"",L""};
		while( !Recset->IsEOF() )
		{
			Recset->GetFieldValue(sFd,gt[0]);
			Recset->GetFieldValue(L"MaTC",gt[1]);

			if(gt[1]!=L"")
			{
				szval.Format(L"SELECT * FROM tbTentieuchuan WHERE MaTC LIKE '%s';",gt[1]);
				RCS = ObjConn2.Execute(szval);

				while( !RCS->IsEOF() )
				{
					RCS->GetFieldValue(L"TenTC",gt[2]);
					if(gt[2]==L"") gt[2] = L"Chưa tồn tại tên tiêu chuẩn";

					m_listViewTC.InsertItem(dem,L"",0);
					m_listViewTC.SetItemText(dem,1,gt[1]);
					m_listViewTC.SetItemText(dem,2,gt[2]);

					dem++;

					break;
				}
				RCS->Close();
			}
			Recset->MoveNext();
		}
		Recset->Close();
		ObjConn2.CloseDatabase();

	}
	catch(...){}
}


void Dlg_Themtieuchuan::OnNMClickL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	nListClick=1;
	m_title.SetWindowText(L"Tích chọn các tiêu chuẩn và nhấn vào nút 'Thêm' để đưa vào danh sách các tiêu chuẩn sẽ sử dụng");

	sDownTC=L"",sFullTC=L"";
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	iTCclick = pNMListView->iItem;
	if(iTCclick>=0)
	{
		sDownTC = m_listTC.GetItemText(iTCclick,1);
		sFullTC = m_listTC.GetItemText(iTCclick,2);
	}
}


void Dlg_Themtieuchuan::OnEnSetfocusT1()
{
	m_title.SetWindowText(L"Để thêm mới tiêu chuẩn, bạn có thể kích chuột phải vào list bên dưới");
}


void Dlg_Themtieuchuan::OnNMClickL4(NMHDR *pNMHDR, LRESULT *pResult)
{
	nListClick=4;
	m_title.SetWindowText(L"Bạn có thể kích chuột phải để xóa các tiêu chuẩn");

	sDownTC=L"",sFullTC=L"";
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	iTCclick = pNMListView->iItem;
	if(iTCclick>=0)
	{
		sDownTC = m_listViewTC.GetItemText(iTCclick,1);
		sFullTC = m_listViewTC.GetItemText(iTCclick,2);
	}
}


void Dlg_Themtieuchuan::OnContextMenu(CWnd* /*pWnd*/, CPoint point)
{
	try
	{
		CMenu mnu1;
		mnu1.LoadMenu(IDR_MENU4);

		CListCtrl *pClist;

		if(nListClick==1) pClist = reinterpret_cast<CListCtrl *>(GetDlgItem(ADNTC_L1));
		else if(nListClick==2) pClist = reinterpret_cast<CListCtrl *>(GetDlgItem(ADNTC_L2));
		else if(nListClick==3) pClist = reinterpret_cast<CListCtrl *>(GetDlgItem(ADNTC_L3));
		else if(nListClick==4) pClist = reinterpret_cast<CListCtrl *>(GetDlgItem(ADNTC_L4));

		CRect rectSubmitButton;
		pClist->GetWindowRect(&rectSubmitButton);

		CMenu *mnuP1 = mnu1.GetSubMenu(0);

		if(nListClick==1) mnuP1->EnableMenuItem(ID_TH40034, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
		else if(nListClick>1)
		{
			mnuP1->EnableMenuItem(ID_TH40031, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			mnuP1->EnableMenuItem(ID_TH40033, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			if(nListClick!=2 && nListClick!=3) mnuP1->EnableMenuItem(ID_TH40034, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
		}

		if(nListClick==3 || iTCclick==-1)
		{
			mnuP1->EnableMenuItem(ID_TH40041, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			mnuP1->EnableMenuItem(ID_TH40042, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
		}
		
		ASSERT(mnuP1);
		if(rectSubmitButton.PtInRect(point))
			mnuP1->TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON, point.x, point.y, this);
	}
	catch(...){}
}


void Dlg_Themtieuchuan::OnNMRClickL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	nListClick=1;	
	sDownTC=L"",sFullTC=L"";
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	iTCclick = pNMListView->iItem;
	if(iTCclick>=0)
	{
		sDownTC = m_listTC.GetItemText(iTCclick,1);
		sFullTC = m_listTC.GetItemText(iTCclick,2);
	}
}


void Dlg_Themtieuchuan::OnNMRClickL3(NMHDR *pNMHDR, LRESULT *pResult)
{
	iTCclick=-1;
	sDownTC=L"",sFullTC=L"";

	nListClick=3;
	*pResult = 0;
}


void Dlg_Themtieuchuan::OnNMRClickL4(NMHDR *pNMHDR, LRESULT *pResult)
{
	nListClick=4;	
	sDownTC=L"",sFullTC=L"";
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	iTCclick = pNMListView->iItem;
	if(iTCclick>=0)
	{
		sDownTC = m_listViewTC.GetItemText(iTCclick,1);
		sFullTC = m_listViewTC.GetItemText(iTCclick,2);
	}
}


void Dlg_Themtieuchuan::OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	nListClick=1;
	
	sDownTC=L"",sFullTC=L"";
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	iTCclick = pNMListView->iItem;
	if(iTCclick>=0)
	{
		sDownTC = m_listTC.GetItemText(iTCclick,1);
		sFullTC = m_listTC.GetItemText(iTCclick,2);
	}

	CString szMa = m_listTC.GetItemText(iTCclick,1);
	CString szTen = m_listTC.GetItemText(iTCclick,2);

	int ncount = m_listThem.GetItemCount();
	m_listThem.InsertItem(ncount,L"",0);
	m_listThem.SetItemText(ncount,1,szMa);
	m_listThem.SetItemText(ncount,2,szTen);
	m_listThem.SetCheck(ncount,1);
}



void Dlg_Themtieuchuan::_SetCheckList(int check)
{
	try
	{
		if(nListClick==1)
		{
			if(m_listTC.GetFirstSelectedItemPosition()!=NULL)
				for (int i = 0; i < m_listTC.GetItemCount(); i++)
					if (m_listTC.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED) m_listTC.SetCheck(i,check);
		}
		else if(nListClick==2)
		{
			if(m_listThem.GetFirstSelectedItemPosition()!=NULL)
				for (int i = 0; i < m_listThem.GetItemCount(); i++)
					if (m_listThem.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED) m_listThem.SetCheck(i,check);
		}
		else if(nListClick==3)
		{
			if(m_listCongviec.GetFirstSelectedItemPosition()!=NULL)
				for (int i = 0; i < m_listCongviec.GetItemCount(); i++)
					if (m_listCongviec.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED) m_listCongviec.SetCheck(i,check);
		}
		else if(nListClick==4)
		{
			if(m_listViewTC.GetFirstSelectedItemPosition()!=NULL)
				for (int i = 0; i < m_listViewTC.GetItemCount(); i++)
					if (m_listViewTC.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED) m_listViewTC.SetCheck(i,check);
		}
	}
	catch(...){}
}


void Dlg_Themtieuchuan::OnTh40029()
{
	_SetCheckList(1);
}


void Dlg_Themtieuchuan::OnTh40030()
{
	_SetCheckList(0);
}


void Dlg_Themtieuchuan::OnTh40031()
{
	qThemTC=1;
	SqlString = L"SELECT * FROM tbTentieuchuan ORDER BY MaTC ASC;";
	AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
	Dlg_all_tieuchuan _dlg;
	_dlg.DoModal();
}


void Dlg_Themtieuchuan::OnTh40032()
{
	if(nListClick==1)
	{
		if(_yn(L"Các tiêu chuẩn sẽ bị xóa hoàn toàn trong CSDL. Bạn có chắc chắn xóa?",0)!=6) return;
		_XoaTieuchuan();
	}
	else if(nListClick==2) OnBnClickedB2();
	else if(nListClick==3)
	{
		if(_yn(L"Tất cả tiêu chuẩn sẽ bị xóa khỏi công việc được chọn trong CSDL. Bạn có chắc chắn xóa?",0)!=6) return;
		_XoaALLTCOfCongviec();
	}
	else if(nListClick==4)
	{
		if(_yn(L"Các tiêu chuẩn sẽ bị xóa khỏi công việc trong CSDL. Bạn có chắc chắn xóa?",0)!=6) return;
		_XoaTCOfCongviec();
	}
}


void Dlg_Themtieuchuan::OnTh40033()
{
	OnBnClickedB1();
}


void Dlg_Themtieuchuan::OnTh40034()
{
	OnBnClickedB3();
}


void Dlg_Themtieuchuan::OnBnClickedB5()
{
	try
	{
		if(m_listThem.GetFirstSelectedItemPosition()!=NULL)
		{
			int nItem=0;
			for (int i = 0; i < m_listThem.GetItemCount(); i++)
				if (m_listThem.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					nItem=i;
					break;
				}
			
			if(nItem>0)
			{
				CString szval=L"";
				int iChk0 = m_listThem.GetCheck(nItem);
				int iChk1 = m_listThem.GetCheck(nItem-1);

				m_listThem.SetCheck(nItem,iChk1);
				m_listThem.SetCheck(nItem-1,iChk0);

				szval = m_listThem.GetItemText(nItem,1);
				m_listThem.SetItemText(nItem,1,m_listThem.GetItemText(nItem-1,1));
				m_listThem.SetItemText(nItem-1,1,szval);

				szval = m_listThem.GetItemText(nItem,2);
				m_listThem.SetItemText(nItem,2,m_listThem.GetItemText(nItem-1,2));
				m_listThem.SetItemText(nItem-1,2,szval);	

				m_listThem.SetItemState(nItem,0,LVIS_SELECTED);
				m_listThem.SetItemState(nItem-1,LVIS_SELECTED,LVIS_SELECTED);
			}
		}
	}
	catch(...){}
}


void Dlg_Themtieuchuan::OnBnClickedB6()
{
	try
	{
		if(m_listThem.GetFirstSelectedItemPosition()!=NULL)
		{
			int nItem=0;
			int count = m_listThem.GetItemCount();
			for (int i = 0; i < count; i++)
				if (m_listThem.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					nItem=i;
					break;
				}

			if(nItem<count-1)
			{
				CString szval=L"";
				int iChk0 = m_listThem.GetCheck(nItem);
				int iChk1 = m_listThem.GetCheck(nItem+1);

				m_listThem.SetCheck(nItem,iChk1);
				m_listThem.SetCheck(nItem+1,iChk0);

				szval = m_listThem.GetItemText(nItem,1);
				m_listThem.SetItemText(nItem,1,m_listThem.GetItemText(nItem+1,1));
				m_listThem.SetItemText(nItem+1,1,szval);

				szval = m_listThem.GetItemText(nItem,2);
				m_listThem.SetItemText(nItem,2,m_listThem.GetItemText(nItem+1,2));
				m_listThem.SetItemText(nItem+1,2,szval);

				m_listThem.SetItemState(nItem,0,LVIS_SELECTED);
				m_listThem.SetItemState(nItem+1,LVIS_SELECTED,LVIS_SELECTED);
			}
		}
	}
	catch(...){}
}


void Dlg_Themtieuchuan::_XoaTieuchuan()
{
	try
	{
		int count = m_listTC.GetItemCount();
		if(count==0) return;

		int iktr=-1;
		for (int i = 0; i < count ; i++)
		{
			if((int)m_listTC.GetCheck(i)==1)
			{
				iktr=i;
				break;
			}
		}

		if(iktr==-1) return;

		CString szval=L"";
		ObjConn2.OpenAccessDB(pathMDB, L"", L"");		
		for (int i = count-1; i >= iktr; i--)
		{
			if((int)m_listTC.GetCheck(i)==1)
			{
				szval = m_listTC.GetItemText(i,1);
				SqlString.Format(L"DELETE FROM tbTentieuchuan WHERE MaTC LIKE '%s';",szval);

				ObjConn2.ExecuteRB(SqlString);
				m_listTC.DeleteItem(i);
			}			
		}

		ObjConn2.CloseDatabase();
	}
	catch(...){}
}


void Dlg_Themtieuchuan::_XoaTCOfCongviec()
{
	try
	{
		CString sTbl=L"tbMaCV_Tieuchuan";
		CString sFd=L"MaCV";
		if(iCbbThemtc==1)
		{
			sTbl=L"tbMaVL_Tieuchuan";
			sFd=L"MaVL";
		}

		int count = m_listViewTC.GetItemCount();
		if(count==0) return;

		int iktr=-1;
		for (int i = 0; i < count ; i++)
		{
			if((int)m_listViewTC.GetCheck(i)==1)
			{
				iktr=i;
				break;
			}
		}

		if(iktr==-1) return;

		CString szval=L"";
		ObjConn2.OpenAccessDB(pathMDB, L"", L"");
		
		for (int i = count-1; i >= iktr; i--)
		{
			if((int)m_listViewTC.GetCheck(i)==1)
			{
				szval = m_listViewTC.GetItemText(i,1);
				SqlString.Format(L"DELETE FROM %s WHERE (%s LIKE '%s' AND MaTC LIKE '%s');",sTbl,sFd,szlMaTC,szval);
				ObjConn2.ExecuteRB(SqlString);
				m_listViewTC.DeleteItem(i);
			}			
		}

		ObjConn2.CloseDatabase();
	}
	catch(...){}
}


void Dlg_Themtieuchuan::_XoaALLTCOfCongviec()
{
	try
	{
		CString sTbl=L"tbMaCV_Tieuchuan";
		CString sFd=L"MaCV";
		if(iCbbThemtc==1)
		{
			sTbl=L"tbMaVL_Tieuchuan";
			sFd=L"MaVL";
		}

		int count = m_listCongviec.GetItemCount();
		if(count==0) return;

		int iktr=-1;
		for (int i = 0; i < count ; i++)
		{
			if((int)m_listCongviec.GetCheck(i)==1)
			{
				iktr=i;
				break;
			}
		}

		if(iktr==-1) return;
		
		CString szval=L"";
		ObjConn2.OpenAccessDB(pathMDB, L"", L"");

		for (int i = count-1; i >= iktr; i--)
		{
			if((int)m_listCongviec.GetCheck(i)==1)
			{
				szval = m_listCongviec.GetItemText(i,1);
				SqlString.Format(L"DELETE FROM %s WHERE %s LIKE '%s';",sTbl,sFd,szval);
				ObjConn2.ExecuteRB(SqlString);
			}			
		}

		ObjConn2.CloseDatabase();
	}
	catch(...){}
}


void Dlg_Themtieuchuan::OnBnClickedCk2()
{
	try
	{
		CString sTbl=L"tbMaCV_Tieuchuan";
		CString sFd=L"MaCV";
		if(iCbbThemtc==1)
		{
			sTbl=L"tbMaVL_Tieuchuan";
			sFd=L"MaVL";
		}

		if((int)m_cksltc.GetCheck()==0)
		{
			m_listCongviec.SetColumnWidth(0,22);
			return;
		}

		m_listCongviec.SetColumnWidth(0,40);
		int count = m_listCongviec.GetItemCount();
		if(count==0) return;		

		ObjConn2.OpenAccessDB(pathMDB, L"", L"");
		CRecordset* Recset;

		m_title.ShowWindow(0);
		m_progress.ShowWindow(1);
		m_progress.SetRange(0,count);
		m_progress.SetWindowText(L"Đang xác định số lượng tiêu chuẩn: ");
		m_progress.SetPos(0);

		CString szval=L"";
		for (int i = 0; i < count; i++)
		{
			szval = m_listCongviec.GetItemText(i,0);
			if(szval==L"")
			{
				szval = m_listCongviec.GetItemText(i,1);
				SqlString.Format(L"SELECT COUNT(*) AS iDem FROM %s WHERE %s LIKE '%s';",sTbl,sFd,szval);
				Recset = ObjConn2.Execute(SqlString);
				Recset->GetFieldValue(L"iDem",szval);
				Recset->Close();
				m_listCongviec.SetItemText(i,0,szval);
			}			
		}

		m_progress.ShowWindow(0);
		m_title.ShowWindow(1);
		m_title.SetWindowText(L"Bạn có thể kích vào cột tiêu đề để sắp xếp số lượng tiêu chuẩn theo thứ tự.");
		ObjConn2.CloseDatabase();
	}
	catch(...){}	
}


CString Dlg_Themtieuchuan::_GetMaTieuchuan()
{
	try
	{
		iTCclick=-1;
		CString szval=L"";
		POSITION pos;

		if(nListClick==1)
		{
			pos = m_listTC.GetFirstSelectedItemPosition();
			if(pos!=NULL)
			{
				for (int i = 0; i < m_listTC.GetItemCount(); i++)
				if (m_listTC.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					iTCclick=i;
					szval = m_listTC.GetItemText(i,1);
					return szval;
				}
			}
		}
		else if(nListClick==2)
		{
			pos = m_listThem.GetFirstSelectedItemPosition();
			if(pos!=NULL)
			{
				for (int i = 0; i < m_listThem.GetItemCount(); i++)
				if (m_listThem.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					iTCclick=i;
					szval = m_listThem.GetItemText(i,1);
					return szval;
				}
			}
		}
		else if(nListClick==4)
		{
			pos = m_listViewTC.GetFirstSelectedItemPosition();
			if(pos!=NULL)
			{
				for (int i = 0; i < m_listViewTC.GetItemCount(); i++)
				if (m_listViewTC.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					iTCclick=i;
					szval = m_listViewTC.GetItemText(i,1);
					return szval;
				}
			}
		}				
	}
	catch(...){return L"";}
	return L"";
}


void Dlg_Themtieuchuan::OnTh40041()
{
	sDownTC = _GetMaTieuchuan();
	f_DocFileTC(sDownTC, 1);
}


void Dlg_Themtieuchuan::OnTh40042()
{
	sDownTC = _GetMaTieuchuan();
	f_SearchGoogle(L"",sDownTC,0);
}

void Dlg_Themtieuchuan::OnCbnSelchangeCb1()
{
	iCbbThemtc = m_cbbox.GetCurSel();
	_TruyvanSQLCongviec();	
}


void Dlg_Themtieuchuan::OnBnClickedPre()
{
	int num = m_cbbFile.GetCurSel();
	num--;

	if(num<0) num = slfileDL-1;
	m_cbbFile.SetCurSel(num);

	//if(sFiledulieu[num+1].Find(L"\\")>0) sFullPath = sFiledulieu[num+1];
	//else sFullPath.Format(L"%s\\%s",sDDanfile,sFiledulieu[num+1]);

	sFullPath = sFpathdulieu[num+1];
	this->SetWindowTextW(L"Thêm nhanh nhóm tiêu chuẩn | " + sFullPath);	

	OnBnClickedB8();
}

void Dlg_Themtieuchuan::OnBnClickedNxt()
{
	int num = m_cbbFile.GetCurSel();
	num++;

	if(num>=slfileDL) num=0;
	m_cbbFile.SetCurSel(num);

	//if(sFiledulieu[num+1].Find(L"\\")>0) sFullPath = sFiledulieu[num+1];
	//else sFullPath.Format(L"%s\\%s",sDDanfile,sFiledulieu[num+1]);

	sFullPath = sFpathdulieu[num+1];
	this->SetWindowTextW(L"Thêm nhanh nhóm tiêu chuẩn | " + sFullPath);

	OnBnClickedB8();
}

void Dlg_Themtieuchuan::OnCbnSelchangeCbbfl()
{
	int num = m_cbbFile.GetCurSel();
	if(num==-1) return;	
	
	//if(sFiledulieu[num+1].Find(L"\\")>0) sFullPath = sFiledulieu[num+1];
	//else sFullPath.Format(L"%s\\%s",sDDanfile,sFiledulieu[num+1]);

	sFullPath = sFpathdulieu[num+1];
	this->SetWindowTextW(L"Thêm nhanh nhóm tiêu chuẩn | " + sFullPath);

	OnBnClickedB8();
}
