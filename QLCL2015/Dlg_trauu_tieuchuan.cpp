#include "Dlg_trauu_tieuchuan.h"
#include "Dlg_all_tieuchuan.h"
#include "Dlg_Themtieuchuan.h"

// Class tra cứu phụ lục vữa
Dlg_trauu_tieuchuan::Dlg_trauu_tieuchuan(CWnd* pParent /*=NULL*/)
	: CDialog(Dlg_trauu_tieuchuan::IDD, pParent)
{
	
}

void Dlg_trauu_tieuchuan::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, TCTC_T1, txtKey);
	DDX_Control(pDX, TCTC_T2, txtLTC);
	DDX_Control(pDX, TCTC_L1, list1);
	DDX_Control(pDX, TCTC_G1, sttic);
}

BEGIN_MESSAGE_MAP(Dlg_trauu_tieuchuan, CDialog)
	ON_WM_SIZE()
	ON_WM_SIZING()
	ON_WM_SYSCOMMAND()
	ON_WM_CONTEXTMENU()
	ON_BN_CLICKED(TCTC_B2, &Dlg_trauu_tieuchuan::OnBnClickedB2)
	ON_BN_CLICKED(TCTC_B1, &Dlg_trauu_tieuchuan::OnBnClickedB1)
	ON_BN_CLICKED(TCTC_B3, &Dlg_trauu_tieuchuan::OnBnClickedB3)
	ON_NOTIFY(NM_DBLCLK, TCTC_L1, &Dlg_trauu_tieuchuan::OnNMDblclkL1)
	ON_NOTIFY(LVN_KEYDOWN, TCTC_L1, &Dlg_trauu_tieuchuan::OnLvnKeydownL1)
	ON_BN_CLICKED(TCTC_B4, &Dlg_trauu_tieuchuan::OnBnClickedB4)
	ON_BN_CLICKED(TCTC_B5, &Dlg_trauu_tieuchuan::OnBnClickedB5)	
	ON_NOTIFY(NM_CLICK, TCTC_L1, &Dlg_trauu_tieuchuan::OnNMClickL1)
	ON_NOTIFY(NM_RCLICK, TCTC_L1, &Dlg_trauu_tieuchuan::OnNMRClickL1)
	ON_COMMAND(ID_TI40039_7, &Dlg_trauu_tieuchuan::OnTi40039)
	ON_COMMAND(ID_TI40040_7, &Dlg_trauu_tieuchuan::OnTi40040)
	ON_NOTIFY(LVN_ENDSCROLL, TCTC_L1, &Dlg_trauu_tieuchuan::OnLvnEndScrollL1)
END_MESSAGE_MAP()

BEGIN_EASYSIZE_MAP(Dlg_trauu_tieuchuan)
	// EASYSIZE(control,left,top,right,bottom,options)
	// options = 1 -> không thay đổi, =0 có thay đổi
	// ES_BORDER = ghim (thay đổi theo)
	// ES_KEEPSIZE = giữ nguyên ban đầu

	EASYSIZE(TCTC_S1,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)	
	EASYSIZE(TCTC_T1,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(TCTC_T2,ES_BORDER,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(TCTC_B1,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)

	EASYSIZE(TCTC_L1,ES_BORDER,ES_BORDER,ES_BORDER,ES_BORDER,0)
	EASYSIZE(TCTC_G1,ES_BORDER,ES_BORDER,ES_BORDER,ES_BORDER,0)

	EASYSIZE(TCTC_B2,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(TCTC_B3,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(TCTC_B4,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(TCTC_B5,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
END_EASYSIZE_MAP

BOOL Dlg_trauu_tieuchuan::OnInitDialog()
{
CDialog::OnInitDialog();
	// Load Icon dự toán
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	iStopload=1;
	tslkq=0,lanshow=1,ibuocnhay=50;

	iTCclick=-1;
	sDownTC=L"",sFullTC=L"";

	// Từ khóa hiển thị hướng dẫn
	txtKey.SetCueBanner(_T("Nhập từ khóa tìm kiếm ..."), TRUE);
	txtLTC.SetCueBanner(_T("Nhập loại tiêu chuẩn ..."), TRUE);
	txtKey.SetWindowText(xl_tukhoa);

	f_Style_tieuchuan();
	f_Load_list_tieuchuan();

	INIT_EASYSIZE;

	return TRUE; 
}

// Bắt sự kiện click Enter
BOOL Dlg_trauu_tieuchuan::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if (pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(TCTC_L1))
	{
		OnBnClickedB2();
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(TCTC_T1))
	{
		OnBnClickedB1();
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(TCTC_T2))
	{
		OnBnClickedB1();
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		OnBnClickedB3();
		return TRUE; 
	}
	else if(pMsg->message == WM_CHAR)
	{
		TCHAR chr = static_cast<TCHAR>(pMsg->wParam);
		if(pWndWithFocus == GetDlgItem(TCTC_L1))
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

void Dlg_trauu_tieuchuan::OnSysCommand(UINT nID, LPARAM lParam)
{
	if(nID == SC_CLOSE) OnBnClickedB3();
	else CDialog::OnSysCommand(nID, lParam);
}

void Dlg_trauu_tieuchuan::OnSize(UINT nType, int cx, int cy) 
{
	CDialog::OnSize(nType, cx, cy);
	UPDATE_EASYSIZE;
}

void Dlg_trauu_tieuchuan::OnSizing(UINT fwSide, LPRECT pRect) 
{
	CDialog::OnSizing(fwSide, pRect);
	EASYSIZE_MINSIZE(560,250,fwSide,pRect);	// chiều rộng + chiều cao
}


void Dlg_trauu_tieuchuan::f_Style_tieuchuan()
{
	// Định dạng kiểu List Control
	list1.InsertColumn(0,L"Mã tiêu chuẩn",LVCFMT_LEFT,_tctc0);
	list1.InsertColumn(1,L"Tên tiêu chuẩn",LVCFMT_LEFT,_tctc1);
	list1.InsertColumn(2,L"Loại tiêu chuẩn",LVCFMT_LEFT,_tctc2);
	list1.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);
}

void Dlg_trauu_tieuchuan::f_delete_list()
{
	list1.DeleteAllItems();
	ASSERT(list1.GetItemCount() == 0);
}

void Dlg_trauu_tieuchuan::f_Load_list_tieuchuan()
{
	TRY
	{
		tslkq=0;
		if(getPathQLCL()==L"") return;
		if(connectDsn(pathMDB)==-1) return;
		ObjConn.OpenAccessDB(pathMDB, L"", L"");
		CRecordset* Recset = ObjConn.Execute(SqlString);
		while( !Recset->IsEOF() )
		{
			tslkq++;
			Recset->GetFieldValue(L"MaTC",XLXD[tslkq].CDG[0]);
			Recset->GetFieldValue(L"TenTC",XLXD[tslkq].CDG[1]);
			Recset->GetFieldValue(L"LoaiTC",XLXD[tslkq].CDG[2]);			
			Recset->MoveNext();
		}
		Recset->Close();
		ObjConn.CloseDatabase();

		lanshow=1;
		int iz=ibuocnhay;
		if(tslkq<=iz)
		{
			iz=tslkq;
			iStopload=0;
		}
		else iStopload=1;

		for (int i = 0; i < iz; i++)
		{
			list1.InsertItem(i,XLXD[i+1].CDG[0],0);
			list1.SetItemText(i,1,XLXD[i+1].CDG[1]);
			list1.SetItemText(i,2,XLXD[i+1].CDG[2]);
		}

		// Kết quả tìm kiếm
		msg.Format(L"Kết quả tìm kiếm (%d tiêu chuẩn)",tslkq);
		sttic.SetWindowText(msg);

		if(tslkq>0) GotoDlgCtrl(GetDlgItem(TCTC_L1));
		else
		{
			int _iL1 = txtKey.LineLength(txtKey.LineIndex(0));
			if(_iL1>0) GotoDlgCtrl(GetDlgItem(TCTC_T1));
			else GotoDlgCtrl(GetDlgItem(TCTC_T2));
		}

	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(L"Lỗi #QL4.672: Lỗi cơ sở dữ liệu tiêu chuẩn."+e->m_strError);
	}
	END_CATCH;
}


void Dlg_trauu_tieuchuan::xl_timkiem_tieuchuan()
{
	try
	{
		CoInitialize(NULL);
		xl.GetActiveObject(L"Excel.Application");
		xl->Visible = true;
		xl->EnableCancelKey = XlEnableCancelKey(FALSE);

		pWb = xl->GetActiveWorkbook();
		pSheet = pWb->ActiveSheet;
		pRange = pSheet->Cells;
		_code = pSheet->CodeName;

		// Kiểm tra sheet hợp lệ
		int iBegin=8;
		if(_code == (_bstr_t)L"shHSNTVL" || _code == (_bstr_t)L"shHSNTCV" 
			|| _code == (_bstr_t)L"shHSNTGD" || _code == (_bstr_t)L"shNhomTC")
		{
			// Vị trí activate
			iRow = pSheet->Application->ActiveCell->Row;
			iColumn = pSheet->Application->ActiveCell->Column;
			cotHyperlink = 0;

			if(_code == (_bstr_t)L"shHSNTVL")
			{
				i1 = getColumn(pSheet,"DMVL_STT");
				i2 = getColumn(pSheet,"DMVL_TC");
				i3 = getColumn(pSheet,"DMVL_TC2");
			}
			else if(_code == (_bstr_t)L"shHSNTCV")
			{
				i1 = getColumn(pSheet,"DMBB_STT");
				i2 = getColumn(pSheet,"DMBB_TC");
				i3 = getColumn(pSheet,"DMBB_TC2");
			}
			else if(_code == (_bstr_t)L"shHSNTGD")
			{
				i1 = getColumn(pSheet,"DMGD_STT");
				i2 = getColumn(pSheet,"DMGD_TC");
				i3 = getColumn(pSheet,"DMGD_MATC");
			}
			else
			{
				// shNhomTC --> Nhóm tiêu chuẩn (09.02.2018)				
				i1 = getColumn(pSheet,"NHTC_STT");
				i2 = getColumn(pSheet,"NHTC_TC");
				i3 = getColumn(pSheet,"NHTC_MA");
				cotHyperlink = getColumn(pSheet, "NHTC_OPEN");

				iBegin=2;
				xde = 50+(int)pRange->SpecialCells(xlCellTypeLastCell)->GetRow();
			}

			// Xác định vị trí cmt END
			if(_code != (_bstr_t)L"shNhomTC") xde = FindComment(pSheet,i1,"END");
			
			// Kiểm tra vị trí hợp lệ
			if(iRow>=iBegin && iRow<xde && (iColumn==i2 || iColumn==i3))
			{
				// Từ khóa
				xl_tukhoa = pRange->GetItem(iRow,iColumn);

				// Xóa các ký tự đặc biệt
				xl_tukhoa.Replace(L"'",L"");
				xl_tukhoa.Replace(L"\\",L"");

				if(xl_tukhoa.Left(1)==L"<" || xl_tukhoa.Find(L"ích phải")>=0 
					|| xl_tukhoa.Find(L"ích chuột")>=0) xl_tukhoa=L"";

				// Kiểm tra xem từ khóa có giá trị không?
				CString ktra =  xl_tukhoa;
				ktra.Replace(L" ",L"");
				if(ktra.GetLength()>0)
				{
					f_Tack_tu_khoa(xl_tukhoa);
					f_xacdinh_SQL(L"");
				}
				else SqlString.Format(L"SELECT * FROM tbTentieuchuan WHERE MaTC<>'' ORDER BY MaTC ASC;");

				// Kiểm tra dữ liệu
				iKiemtra = xl_kiemtra_dulieu();

				/*if(iKiemtra==1)
				{
					if(iRow+3 > xde)
						pSheet->GetRange(pRange->GetItem(xde,1),pRange->GetItem(xde+3,1))->EntireRow->Insert(xlShiftDown);

					pRange->PutItem(iRow,i2,(_bstr_t)val2);
					if(i3>0) pRange->PutItem(iRow,i3,(_bstr_t)val1);
				}
				else */if(iKiemtra>=1)
				{
					AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
					xl->PutScreenUpdating(VARIANT_TRUE);
					DoModal();
					xl->PutScreenUpdating(VARIANT_FALSE);
				}
			}
		}

		CoUninitialize();
	}
	catch(_com_error & error){ _s(L"#QL4.66: Lỗi tra cứu tiêu chuẩn.");}
}


long Dlg_trauu_tieuchuan::xl_kiemtra_dulieu()
{
	long cong = 0;
	TRY
	{
		if(getPathQLCL()==L"") return 0;
		if(connectDsn(pathMDB)==-1) return 0;
		ObjConn.OpenAccessDB(pathMDB, L"", L"");
		CRecordset *Recset = ObjConn.Execute(SqlString);
		while( !Recset->IsEOF() )
		{
			Recset->GetFieldValue(L"MaTC",val1);
			Recset->GetFieldValue(L"TenTC",val2);
			Recset->GetFieldValue(L"LoaiTC",val3);
			cong++;

			break;		// tạm thời khi có dữ liệu iKiemtra >= 1 --> hiện hộp thoại luôn

			Recset->MoveNext();
		}
		Recset->Close();
		ObjConn.CloseDatabase();
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(L"Lỗi #QL4.671: Lỗi cơ sở dữ liệu tiêu chuẩn."+e->m_strError);
	}
	END_CATCH;

	return cong;
}



void Dlg_trauu_tieuchuan::xl_xacdinh_sheet(int sl)
{
	try
	{
		// 06/08/2020
		// Giới hạn số lượng download =1
		sl = 1;

		xl->EnableCancelKey = XlEnableCancelKey(FALSE);
		xl->PutScreenUpdating(VARIANT_FALSE);

		// Kiểm tra xem số lượng chèn có vượt quá số thứ tự tiếp theo hay vị trí END không?
		CString f1,f2;
		int tang = iRow+1;
		f1 = pRange->GetItem(tang,i1);
		f2 = pRange->GetItem(tang,i2);

		while (f1==L"" && f2==L"" && tang<xde)
		{
			tang++;
			f1 = pRange->GetItem(tang,i1);
			f2 = pRange->GetItem(tang,i2);
		}

		int k=iRow+sl-tang;
		if(k>0)
		{
			if(tang==xde) k = k+3;
			PRS=pSheet->GetRange(pRange->GetItem(tang,1),pRange->GetItem(tang+k-1,1));
			PRS->EntireRow->Insert(xlShiftDown);
		}

		// Đổ dữ liệu mã tiêu chuẩn
		for (int i = 1; i <=sl; i++)
		{
			pRange->PutItem(iRow+i-1,i2,(_bstr_t)_nametc[i]);
				
			// Kiểm tra xem cột bên cạnh có giá trị chưa? Chưa có thì đổ thêm ký tự " "
			msg = pRange->GetItem(iRow+i-1,i2+1);
			if(msg=L"") pRange->PutItem(iRow+i-1,i2+1,(_bstr_t)L" ");
			if(i3>0) pRange->PutItem(iRow+i-1,i3,(_bstr_t)_matc[i]);

			// Download và gán link tiêu chuẩn
			if (cotHyperlink > 0)
			{
				CString szLink = f_DocFileTC(_matc[i], 0);
				if (szLink != L"")
				{
					PRS = pRange->GetItem(iRow + i - 1, cotHyperlink);
					_xlSetHyperlink(pSheet, PRS, szLink, L"Mở file",
						xlNone, RGB(0, 0, 255), L"Arial", 12, true, true);
				}
			}
		}
	}
	catch(_com_error & error){_s(L"#QL4.69: Lỗi xác định sheet.");}
}


void Dlg_trauu_tieuchuan::f_Tack_tu_khoa(CString tukhoa)
{
	// Truyền giá trị từ khóa tại text-box vào biến txt_timkiem
	txt_timkiem  = new wchar_t[SIZE_LINE];
	wcscpy(txt_timkiem,tukhoa);

	// Gán giá trị cho các từ khóa được chia nhỏ
	for(int i=1;i<=10;i++) wcscpy(pKey[i],L"");

	// Tách chuỗi thành các từ khóa nhỏ pKey[i]
	wchar_t * tg = new wchar_t[SIZE_LINE];
	iKey=0;	// số lượng từ khóa
	tg = wcstok (txt_timkiem,L"+");
	while (tg != NULL)
		{
			iKey++;
			wcscpy(pKey[iKey],tg);
			tg = wcstok(NULL,L"+");
		}

	// Xóa khoảng trắng trước và sau chuỗi
	int k;
	for(int i=1;i<=iKey;i++)
	{
		if(wcscmp(pKey[i],L" ")==0)	wcscpy(pKey[i],L"");
		else
		{
			// xóa trước
			k=0;	
			while(wcscmp(tachkytu_(pKey[i],k+1),L" ")==0 && wcslen(pKey[i])>0) k++;
			wcscpy(pKey[i],&pKey[i][k]);

			// xóa sau
			k=wcslen(pKey[i]);	
			while(wcscmp(tachkytu_(pKey[i],k),L" ")==0 && k>=0) k--;
			if(k>=0 && k<wcslen(pKey[i])) wcscpy(pKey[i],tachchuoi_(pKey[i],k,0));
		}
	}
}



void Dlg_trauu_tieuchuan::f_Tack_loai_tieuchuan(CString tukhoa)
{
	// Truyền giá trị từ khóa tại text-box vào biến txt_timkiem
	txt_timkiem  = new wchar_t[SIZE_LINE];
	wcscpy(txt_timkiem,tukhoa);

	// Gán giá trị cho các từ khóa được chia nhỏ
	for(int i=1;i<=10;i++) wcscpy(pLTC[i],L"");

	// Tách chuỗi thành các từ khóa nhỏ pLTC[i]
	wchar_t * tg = new wchar_t[SIZE_LINE];
	iLTC=0;	// số lượng từ khóa
	tg = wcstok (txt_timkiem,L"+");
	while (tg != NULL)
		{
			iLTC++;
			wcscpy(pLTC[iLTC],tg);
			tg = wcstok(NULL,L"+");
		}

	// Xóa khoảng trắng trước và sau chuỗi
	int k;
	for(int i=1;i<=iLTC;i++)
	{
		if(wcscmp(pLTC[i],L" ")==0)	wcscpy(pLTC[i],L"");
		else
		{
			// xóa trước
			k=0;	
			while(wcscmp(tachkytu_(pLTC[i],k+1),L" ")==0 && wcslen(pLTC[i])>0) k++;
			wcscpy(pLTC[i],&pLTC[i][k]);

			// xóa sau
			k=wcslen(pLTC[i]);	
			while(wcscmp(tachkytu_(pLTC[i],k),L" ")==0 && k>=0) k--;
			if(k>=0 && k<wcslen(pLTC[i])) wcscpy(pLTC[i],tachchuoi_(pLTC[i],k,0));
		}
	}
}


CString Dlg_trauu_tieuchuan::f_xacdinh_sqlLTC()
{
	CString _gtri = L"";
	for (int i = 1; i <= iLTC; i++) _gtri.Format(L"%sAND LoaiTC LIKE '%s%s%s' ",_gtri,L"%",pLTC[i],L"%");
	return _gtri;
}


void Dlg_trauu_tieuchuan::f_xacdinh_SQL(CString _sql2)
{
	CString szval=L"";
	SqlString.Format(L"SELECT * FROM tbTentieuchuan "
		L"WHERE (MaTC LIKE '%s%s%s' OR TenTC LIKE '%s%s%s' OR Ghichu LIKE '%s%s%s') ",
		L"%",pKey[1],L"%",L"%",pKey[1],L"%",L"%",pKey[1],L"%");

	if(iKey>1)
	{
		msg = L"";
		for (int i = 2; i <= iKey; i++)
		{
			szval = msg;
			msg.Format(L"%sAND (MaTC LIKE '%s%s%s' OR TenTC LIKE '%s%s%s' OR Ghichu LIKE '%s%s%s') ",
				szval,L"%",pKey[i],L"%",L"%",pKey[i],L"%",L"%",pKey[i],L"%");
		}

		SqlString = SqlString+msg;
	}

	szval = SqlString;
	SqlString.Format(L"%s AND MaTC <> '' %s",szval,_sql2);
	SqlString+= L" ORDER BY MaTC ASC;";
}


void Dlg_trauu_tieuchuan::OnBnClickedB2()
{
	try
	{
		// Hàm check bản quyền có điều kiện đếm
		if(f_Mod_check()!=1) return;

		int dem=0;
		POSITION pos=list1.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			for(int i=0;i<list1.GetItemCount();i++)		
			{
				if (list1.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					dem++;
					_matc[dem] =  list1.GetItemText(i,0);
					_nametc[dem] =  list1.GetItemText(i,1);
				}
			}

			xl_xacdinh_sheet(dem);
			CDialog::OnOK();
		}
	}
	catch(_com_error & error){}
}


void Dlg_trauu_tieuchuan::OnBnClickedB1()
{
	txtLTC.GetWindowTextW(xl_LTC);
	xl_LTC.Trim();
	xl_LTC.Replace(L"'",L"");

	CString sqlLTC=L"";
	if(xl_LTC.GetLength()>0)
	{
		f_Tack_loai_tieuchuan(xl_LTC);
		sqlLTC = f_xacdinh_sqlLTC();
	}

	txtKey.GetWindowTextW(xl_tukhoa);
	xl_tukhoa.Trim();
	xl_tukhoa.Replace(L"'",L"");

	if(xl_tukhoa.GetLength()>0)
	{
		f_Tack_tu_khoa(xl_tukhoa);
		f_xacdinh_SQL(sqlLTC);
	}
	else SqlString.Format(L"SELECT * FROM tbTentieuchuan WHERE MaTC <> '' %s ORDER BY MaTC ASC;",sqlLTC);
	
	f_delete_list();
	f_Load_list_tieuchuan();
}


void Dlg_trauu_tieuchuan::OnBnClickedB3()
{
	xl->EnableCancelKey = XlEnableCancelKey(FALSE);
	xl->PutScreenUpdating(VARIANT_FALSE);

	_tctc0 = list1.GetColumnWidth(0);
	_tctc1 = list1.GetColumnWidth(1);
	_tctc2 = list1.GetColumnWidth(2);

	CDialog::OnCancel();
}


void Dlg_trauu_tieuchuan::OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	OnBnClickedB2();
}


void Dlg_trauu_tieuchuan::OnLvnKeydownL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);
	if(pLVKeyDow->wVKey == VK_SPACE) OnBnClickedB2();

	*pResult = 0;
}


void Dlg_trauu_tieuchuan::OnBnClickedB4()
{
	qThemTC=1;
	AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
	Dlg_all_tieuchuan _dlg;
	_dlg.DoModal();
}


void Dlg_trauu_tieuchuan::OnBnClickedB5()
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
	Dlg_Themtieuchuan open_dlg;
	open_dlg.DoModal();
}


void Dlg_trauu_tieuchuan::OnContextMenu(CWnd* /*pWnd*/, CPoint point)
{
	try
	{
		if(iTCclick==-1) return;

		CMenu mnu2;
		mnu2.LoadMenu(IDR_MENU7);
		
		CListCtrl *pClist;
		pClist = reinterpret_cast<CListCtrl *>(GetDlgItem(TCTC_L1));

		CRect rectSubmitButton;
		pClist->GetWindowRect(&rectSubmitButton);

		CMenu *mnuP2 = mnu2.GetSubMenu(0);
		ASSERT(mnuP2);

		if( rectSubmitButton.PtInRect(point) )
			mnuP2->TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON, point.x, point.y, this);
	}
	catch(...){}
}


void Dlg_trauu_tieuchuan::OnNMClickL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	sDownTC=L"",sFullTC=L"";
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	iTCclick = pNMListView->iItem;
	if(iTCclick>=0)
	{
		sDownTC = list1.GetItemText(iTCclick,0);
		sFullTC = list1.GetItemText(iTCclick,1);
	}
}


void Dlg_trauu_tieuchuan::OnNMRClickL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	sDownTC=L"",sFullTC=L"";
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	iTCclick = pNMListView->iItem;
	if(iTCclick>=0)
	{
		sDownTC = list1.GetItemText(iTCclick,0);
		sFullTC = list1.GetItemText(iTCclick,1);
	}
}


CString Dlg_trauu_tieuchuan::_GetMaTieuchuan()
{
	try
	{
		iTCclick=-1;
		CString szval=L"";
		POSITION pos = list1.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			for (int i = 0; i < list1.GetItemCount(); i++)
			if (list1.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
			{
				iTCclick=i;
				szval = list1.GetItemText(i,0);
				return szval;
			}
		}		
	}
	catch(...){return L"";}
	return L"";
}


void Dlg_trauu_tieuchuan::OnTi40039()
{
	sDownTC = _GetMaTieuchuan();
	f_DocFileTC(sDownTC, 1);
}

void Dlg_trauu_tieuchuan::OnTi40040()
{
	sDownTC = _GetMaTieuchuan();
	f_SearchGoogle(L"",sDownTC,0);
}

void Dlg_trauu_tieuchuan::OnLvnEndScrollL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	try
	{
		if(iStopload==0) return;

		RECT r;
		CRect rectCtrl;
		list1.GetItemRect(0, &r, LVIR_BOUNDS);
		list1.GetWindowRect(&rectCtrl);
		int a = r.bottom - r.top;	
		int b = rectCtrl.Height();
		int topIndex = list1.GetTopIndex();
		int ncount = list1.GetItemCount();
		if(topIndex+(int)(b/a) >= ncount)
		{
			lanshow++;
			int iz=ibuocnhay*lanshow;
			if(tslkq<=iz)
			{
				iz=tslkq;
				iStopload=0;
			}
			else iStopload=1;

			for (int i = ibuocnhay*(lanshow-1); i < iz; i++)
			{
				list1.InsertItem(i,XLXD[i+1].CDG[0],0);
				list1.SetItemText(i,1,XLXD[i+1].CDG[1]);
				list1.SetItemText(i,2,XLXD[i+1].CDG[2]);
			}

			list1.EnsureVisible(ncount+(int)(b/a)-5, FALSE);
		}
	}
	catch(...){}
}
