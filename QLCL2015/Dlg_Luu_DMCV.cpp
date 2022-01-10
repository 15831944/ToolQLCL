// Dlg_Luu_DMCV message handlers
#include "Dlg_Luu_DMCV.h"

Dlg_Luu_DMCV::Dlg_Luu_DMCV() : CDialog(Dlg_Luu_DMCV::IDD)
{
	ObjConn.CloseDatabase();
}

void Dlg_Luu_DMCV::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, LDMCV_R1, _nRad1);
	DDX_Control(pDX, LDMCV_R2, _nRad2);
	DDX_Control(pDX, LDMCV_C1, _cbb0);
}

BEGIN_MESSAGE_MAP(Dlg_Luu_DMCV, CDialog)
	
	ON_BN_CLICKED(LDMCV_B1, &Dlg_Luu_DMCV::OnBnClickedB1)
	ON_BN_CLICKED(LDMCV_B2, &Dlg_Luu_DMCV::OnBnClickedB2)
	ON_BN_CLICKED(LDMCV_R1, &Dlg_Luu_DMCV::OnBnClickedR1)
	ON_BN_CLICKED(LDMCV_R2, &Dlg_Luu_DMCV::OnBnClickedR2)
END_MESSAGE_MAP()

BOOL Dlg_Luu_DMCV::OnInitDialog()
{
CDialog::OnInitDialog();
	// Load Icon dự toán
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	if(shDMCV==(_bstr_t)L"shHSNTCV")
	{
		_nRad1.SetWindowText(L"Lưu các công việc được lựa chọn");
		_nRad2.SetWindowText(L"Giai đoạn");
		msg = L"Lưu danh mục công việc";
	}
	else
	{
		_nRad1.SetWindowText(L"Lưu các vật liệu được lựa chọn");
		_nRad2.SetWindowText(L"Tất cả");
		msg = L"Lưu danh mục vật liệu";
	}

	this->SetWindowTextW(msg);

	_nChk0=1;
	_nRad1.SetCheck(TRUE);
	_cbb0.EnableWindow(FALSE);
	f_Load_giaidoan();

	return TRUE; 
}

// Bắt sự kiện click Enter
BOOL Dlg_Luu_DMCV::PreTranslateMessage(MSG* pMsg)
{
	 CWnd* pWndWithFocus = GetFocus();
	if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(LDMCV_B1))
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


void Dlg_Luu_DMCV::OnBnClickedB1()
{
	CDialog::OnOK();

	if(getPathQLCL()==L"") return;
	if(connectDsn(pathMDB)==-1) return;
	ObjConn.OpenAccessDB(pathMDB, L"", L"");

	int _sll=0;
	if(_nChk0==1)
	{
		// Chọn theo vùng
		// Xác định vùng được lựa chọn
		PRS = psDMCV->Application->Selection;
		_bstr_t _bsArr = PRS->GetAddress(1,1,XlReferenceStyle::xlR1C1);
		int _nSelection = _fGetSelection((LPCTSTR)_bsArr);
		if(_nSelection>0)
		{
			for (int i = 1; i <= _nSelection; i++)
			{
				int _pos = _arrSLS[i].Find(L"-");
				if(_pos==-1)
				{
					// Chỉ có 1 lựa chọn
					_sll+=f_Luu_1_danhmuc(_wtoi(_arrSLS[i]));
				}
				else
				{
					// Có nhiều lựa chọn
					int jlen = _arrSLS[i].GetLength();
					int jbd = _wtoi(_arrSLS[i].Left(_pos));
					int jkt = _wtoi(_arrSLS[i].Right(jlen-_pos-1));
					for (int j = jbd; j <=jkt; j++) _sll+=f_Luu_1_danhmuc(j);
				}
			}
		}
	}
	else
	{
		// Chọn theo hạng mục
		int ncount = _cbb0.GetCurSel();
		int _ibd=8,_ikt=xde;
		if(ncount>0)
		{
			_ibd = nvtri[ncount];
			_ikt = nvtri[ncount+1];
		}

		for (int i = _ibd; i < _ikt; i++) _sll+=f_Luu_1_danhmuc(i);
	}
	ObjConn.CloseDatabase();

	if(_sll==0) _s(L"Không có dữ liệu được lưu. \nVùng lựa chọn chưa đúng, hoặc dữ liệu đã tồn tại trong CSDL.");
	else _s(L"Dữ liệu được lưu thành công.");
}


int Dlg_Luu_DMCV::f_Luu_1_danhmuc(int _vtri)
{
	try
	{
		// Kiểm tra vị trí có thỏa mãn điều kiện lưu không?
		CString _vma,val1,val2;
		val1 = prDMCV->GetItem(_vtri,_ncot[1]);
		_vma = prDMCV->GetItem(_vtri,_ncot[2]);
		if(_wtoi(val1)==0 || _vma==L"") return 0;

		// Kiểm tra dữ liệu có trong CSDL không?
		CString sDem=L"";
		if(shDMCV==(_bstr_t)L"shHSNTCV") SqlString.Format(L"SELECT COUNT(*) AS iDem FROM tbMaCV_loaiCV WHERE MaCV LIKE '%s';",_vma);
		else SqlString.Format(L"SELECT COUNT(*) AS iDem FROM tbMaVL_tenVL WHERE MaVL LIKE '%s';",_vma);
		CRecordset* Recset = ObjConn.Execute(SqlString);
		Recset->GetFieldValue(L"iDem",sDem);
		Recset->Close();
		if(_wtoi(sDem)>0) return 0;

		// Lưu mã vật liệu, mã công việc
		val1 = prDMCV->GetItem(_vtri,_ncot[3]);		
		if(shDMCV==(_bstr_t)L"shHSNTCV") msg.Format(L"INSERT INTO tbMaCV_loaiCV (MaCV,LoaiCV) VALUES ('%s','%s');",_vma,val1);
		else msg.Format(L"INSERT INTO tbMaVL_tenVL (MaVL,TenVL) VALUES ('%s','%s');",_vma,val1);
		ObjConn.ExecuteRB(msg);

		// Xác định khoảng chứa các tiêu chuẩn
		int _icuoi;
		for (_icuoi = _vtri+1; _icuoi <= xde; _icuoi++)
		{
			val1 = prDMCV->GetItem(_icuoi,_ncot[1]);
			val2 = prDMCV->GetItem(_icuoi,_ncot[2]);
			if(_wtoi(val1)>0 && val2!=L"") break;
		}

		// Lưu mã tiêu chuẩn
		for (int i = _vtri; i < _icuoi; i++)
		{
			val1 = prDMCV->GetItem(i,_ncot[4]);
			val1.Trim();
			if(val1!=L"")
			{
				if(shDMCV==(_bstr_t)L"shHSNTCV") msg.Format(L"INSERT INTO tbMaCV_Tieuchuan (MaCV,MaTC) VALUES ('%s','%s');",_vma,val1);
				else msg.Format(L"INSERT INTO tbMaVL_Tieuchuan (MaVL,MaTC) VALUES ('%s','%s');",_vma,val1);
				ObjConn.ExecuteRB(msg);
			}
		}
	}
	catch(_com_error & error){}

	return 1;
}


void Dlg_Luu_DMCV::OnBnClickedB2()
{
	OnCancel();
}


void Dlg_Luu_DMCV::OnBnClickedR1()
{
	_nChk0=1;
	_cbb0.EnableWindow(FALSE);
}


void Dlg_Luu_DMCV::OnBnClickedR2()
{
	_nChk0=2;
	_cbb0.EnableWindow(TRUE);
}


void Dlg_Luu_DMCV::f_Load_giaidoan()
{
	try
	{
		// Xác định cột & END
		if(shDMCV==(_bstr_t)L"shHSNTCV")
		{
			_ncot[1] = getColumn(psDMCV,"DMBB_STT");
			_ncot[2] = getColumn(psDMCV,"DMBB_MACV");
			_ncot[3] = getColumn(psDMCV,"DMBB_ND");
			_ncot[4] = getColumn(psDMCV,"DMBB_TC2");
		}
		else
		{
			_ncot[1] = getColumn(psDMCV,"DMVL_STT");
			_ncot[2] = getColumn(psDMCV,"DMVL_MAVL");
			_ncot[3] = getColumn(psDMCV,"DMVL_ND");
			_ncot[4] = getColumn(psDMCV,"DMVL_TC2");
		}

		xde = FindComment(psDMCV,_ncot[1],"END");
		CString val=L"";
		
		if(shDMCV==(_bstr_t)L"shHSNTCV")
		{
			int dem=0;
			_cbb0.AddString(L"Tất cả giai đoạn");
			for (int i = 8; i < xde; i++)
			{
				val = prDMCV->GetItem(i,_ncot[2]);
				val.Trim();
				if(val.MakeUpper()==L"GD")
				{
					dem++;
					nvtri[dem]=i;
					val = prDMCV->GetItem(i,_ncot[3]);
					_cbb0.AddString(val);
				}
			}

			dem++;
			nvtri[dem]=xde;
		}
		else _cbb0.AddString(L"Tất cả công việc");

		_cbb0.SetCurSel(0);
	}
	catch(_com_error & error){}
}



// Tra cứu lại công tác (open)
void Dlg_Luu_DMCV:: f_kiemtra_sheet()
{
	try
	{
		CoInitialize(NULL);
		xl.GetActiveObject(L"Excel.Application");
		xl->Visible = true;
		xl->EnableCancelKey = XlEnableCancelKey(FALSE);

		pWb = xl->GetActiveWorkbook();
		psDMCV = pWb->ActiveSheet;
		prDMCV = psDMCV->Cells;
		shDMCV = psDMCV->CodeName;

		if(shDMCV==(_bstr_t)L"shHSNTCV" || shDMCV==(_bstr_t)L"shHSNTVL")
		{
			AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
			DoModal();
		}
		else _s(L"Sử dụng chức năng này tại sheet danh mục vật liệu hoặc danh mục công việc.");
	
	CoUninitialize();
	}
	catch(_com_error & error){}
}
