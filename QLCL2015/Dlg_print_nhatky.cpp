// Dlg_print_nhatky message handlers
#include "Dlg_print_nhatky.h"
#include "Dg_In_ho_so.h"
#include "Dlg_Setup_Print.h"
#include "cls_main.h"
#include "Dlg_luu_Hso.h"

Dlg_print_nhatky::Dlg_print_nhatky() : CDialog(Dlg_print_nhatky::IDD)
{
}

void Dlg_print_nhatky::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, NKY_T1, txt1);
	DDX_Control(pDX, NKY_D1, dt1);
	DDX_Control(pDX, NKY_D2, dt2);
	DDX_Control(pDX, NKY_R1, rd1);
	DDX_Control(pDX, NKY_R2, rd2);
	DDX_Control(pDX, NKY_R3, rd3);
	DDX_Control(pDX, NKY_R4, rd4);
	DDX_Control(pDX, NKY_R5, rd5);
	DDX_Control(pDX, NKY_R6, rd6);
	DDX_Control(pDX, NKY_S1, stt1);
	DDX_Control(pDX, NKY_CK1, fdkem);
}

BEGIN_MESSAGE_MAP(Dlg_print_nhatky, CDialog)
	ON_BN_CLICKED(NKY_B2, &Dlg_print_nhatky::OnBnClickedB2)
	ON_BN_CLICKED(NKY_B3, &Dlg_print_nhatky::OnBnClickedB3)
	ON_BN_CLICKED(NKY_B1, &Dlg_print_nhatky::OnBnClickedB1)
	ON_BN_CLICKED(NKY_R1, &Dlg_print_nhatky::OnBnClickedR1)
	ON_BN_CLICKED(NKY_R2, &Dlg_print_nhatky::OnBnClickedR2)
	ON_BN_CLICKED(NKY_CK1, &Dlg_print_nhatky::OnBnClickedCk1)
	ON_EN_KILLFOCUS(NKY_T1, &Dlg_print_nhatky::OnEnKillfocusT1)
	ON_BN_CLICKED(NKY_B4, &Dlg_print_nhatky::OnBnClickedB4)
	ON_BN_CLICKED(NKY_B5, &Dlg_print_nhatky::OnBnClickedB5)
	ON_BN_CLICKED(NKY_R3, &Dlg_print_nhatky::OnBnClickedR3)
	ON_BN_CLICKED(NKY_R4, &Dlg_print_nhatky::OnBnClickedR4)
	ON_BN_CLICKED(NKY_R5, &Dlg_print_nhatky::OnBnClickedR5)
	ON_BN_CLICKED(NKY_R6, &Dlg_print_nhatky::OnBnClickedR6)
END_MESSAGE_MAP()

BOOL Dlg_print_nhatky::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	_ird=0;
	_irp=2;
	rd1.SetCheck(1);
	rd4.SetCheck(1);
	dt1.EnableWindow(0);
	dt2.EnableWindow(0);
	fdkem.SetCheck(1);
	_idkem=1;

	txt1.SetWindowText(L"1");
	stt1.SetWindowText(f_ngaythang());

	if(iSetup==0) 
	{
		Dg_In_ho_so _dlg;
		_dlg.f_thietlap_giatri();
		iSetup=1;
	}

	return TRUE; 
}

// Bắt sự kiện click Enter
BOOL Dlg_print_nhatky::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(NKY_B2))
	{
		OnBnClickedB2();
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		/*try
		{
			CoInitialize(NULL);
			xl.GetActiveObject(L"Excel.Application");
			xl->Visible = true;
			xl->EnableCancelKey = XlEnableCancelKey(FALSE);
			CoUninitialize();
		}
		catch(_com_error & error){}*/

		CDialog::OnCancel();
		return TRUE; 
	}

	return FALSE;
}


CString Dlg_print_nhatky::f_ngaythang()
{
	msg = L"";

	try
	{
		/*CoInitialize(NULL);
		xl.GetActiveObject(L"Excel.Application");
		xl->Visible = true;
		xl->EnableCancelKey = XlEnableCancelKey(FALSE);*/

		// Định nghĩa các sheet
		shNDNK = name_sheet("shMauNKY4");
		psNDNK=xl->Sheets->GetItem(shNDNK);
		prNDNK = psNDNK->Cells;

		// Kiểm tra ẩn hiện sheet
		int iVisible = psNDNK->GetVisible();
		if (iVisible!=-1) psNDNK->PutVisible(XlSheetVisibility::xlSheetVisible);
		psNDNK->Activate();

		// Xác định ngày bắt đầu-kết thúc (toàn bộ)
		ngbd = getVCell(psNDNK,L"HK_NBD");
		ngkt = getVCell(psNDNK,L"HK_NKT");
		msg.Format(L"(Từ ngày %s đến %s)",ngbd,ngkt);	

		// Set giá trị ngày tháng datetime control
		CString val = getVCell(psNDNK,L"HK_NGAY");
		f_set_datetime(val);

		//CoUninitialize();
	}
	catch(_com_error & error){_s(L"#QL4.351: Lỗi in nội dung nhật ký.");}
	return msg;
}


void Dlg_print_nhatky::f_set_datetime(CString _sTime)
{
	try
	{
		int jnam = _wtoi(_sTime.Right(4));
		if(jnam==0) jnam=2017;

		int jthang = _wtoi(_sTime.Mid(3,2));
		if(jthang==0) jthang=1;

		int jngay = _wtoi(_sTime.Left(2));
		if(jngay==0) jngay=1;

		CDateTimeCtrl* pCtrl1 = (CDateTimeCtrl*) GetDlgItem(NKY_D1);
		CDateTimeCtrl* pCtrl2 = (CDateTimeCtrl*) GetDlgItem(NKY_D2);
		CTime _ctt(jnam, jthang, jngay, 0, 0, 0);
		VERIFY(pCtrl1->SetTime(&_ctt));
		VERIFY(pCtrl2->SetTime(&_ctt));
	}
	catch(...){}
}

void Dlg_print_nhatky::OnBnClickedB2()
{
	try
	{
		// Hàm check bản quyền
		if(f_Check_ban_quyen()!=1) return;

		_iptrang=0;
		_idkem = fdkem.GetCheck();
		if(_idkem==1) _iptrang = _irp;
		
		_bstr_t _prter = xl->Application->GetActivePrinter();
		//if(f_Check_printer(_prter)==1) return;     //<-- Tạm thời comment test in (09.08.2017)

		/*CoInitialize(NULL);
		xl.GetActiveObject(L"Excel.Application");
		xl->Visible = true;
		xl->EnableCancelKey = XlEnableCancelKey(FALSE);*/

		// Định nghĩa các sheet
		shNDNK = name_sheet("shMauNKY4");
		psNDNK=xl->Sheets->GetItem(shNDNK);
		prNDNK = psNDNK->Cells;

		// Số bản in
		iLen = txt1.LineLength(txt1.LineIndex(0));
		lp = numin.GetBuffer(iLen);
		txt1.GetLine(0, lp,iLen);
		numin.ReleaseBuffer();

		_iNgay=10;
		BOOLEAN fx = TRUE;
	
		if(_ird==1)
		{
			// Xác định ngày bắt đầu & kết thúc (lựa chọn)
			GetDlgItemText(NKY_D1,ngay1);
			GetDlgItemText(NKY_D2,ngay2);
			_iNgay = ngay1.GetLength();

			// So sánh ngày tháng
			int f0 = compare_date(_iNgay,ngay1,ngay2);
			int f1 = compare_date(_iNgay,ngbd,ngay1);
			int f2 = compare_date(_iNgay,ngay2,ngkt);

			if(f0==1 && f1==1 && f2==1) 
			{
				fx = TRUE;
				ngbd = ngay1;
				ngkt = ngay2;
			}
			else fx = FALSE;
		}

		if(fx == TRUE)
		{
			fx = FALSE;
			msg = L"";
			
			if(ngbd==ngkt) msg.Format(L"Bạn có chắc chắn in nhật ký ngày %s?",ngbd);
			else
			{
				if(_ird==0) msg.Format(L"Bạn có chắc chắn in TOÀN BỘ NHẬT KÝ \nTừ ngày %s đến %s?",ngbd,ngkt);
				else msg.Format(L"Bạn có chắc chắn in nhật ký từ ngày %s đến %s?",ngbd,ngkt);
			}	
				
			msg.Format(L"%s \n-Máy in: %s",msg,(LPCTSTR)_prter);

			int result = MessageBox(msg,_T("Thông báo"), MB_YESNO | MB_ICONQUESTION);
			if (result == 6) fx = TRUE;
			else fx = FALSE;

			if(fx == TRUE)
			{
				CDialog::OnOK();
				xl->PutScreenUpdating(VARIANT_FALSE);

				psNDNK->GetRange(prNDNK->GetItem(2,99),prNDNK->GetItem(3,99))->PutNumberFormat("dd/mm/yyyy");
				psNDNK->GetRange(prNDNK->GetItem(2,100),prNDNK->GetItem(3,100))->PutNumberFormat(L"General");

				f_ktra_date(ngbd,prNDNK,2,99);
				f_ktra_date(ngkt,prNDNK,3,99);

				int i1 = getRow(psNDNK,"HK_NBD");
				int i2 = getColumn(psNDNK,"HK_NBD");

				int luuHKN1 = getRow(psNDNK,"HK_NGAY"); 
				int luuHKN2 = getColumn(psNDNK,"HK_NGAY");				

				msg.Format(L"=%s-%s",GAR(2,99,prNDNK,0),GAR(i1,i2,prNDNK,0));
				prNDNK->PutItem(2,100,(_bstr_t)msg);
				msg.Format(L"=%s-%s",GAR(3,99,prNDNK,0),GAR(2,99,prNDNK,0));
				prNDNK->PutItem(3,100,(_bstr_t)msg);

				CString _ibdau = prNDNK->GetItem(2,100);
				CString _itong = prNDNK->GetItem(3,100);

				msg = prNDNK->GetItem(2,100);
				i1 = getRow(psNDNK,"HK_CONG");
				i2 = getColumn(psNDNK,"HK_CONG"); 
				prNDNK->PutItem(i1,i2,(_bstr_t)msg);
				psNDNK->GetRange(prNDNK->GetItem(i1,i2),prNDNK->GetItem(i1,i2))->NumberFormat = L"General";
				psNDNK->GetRange(prNDNK->GetItem(2,99),prNDNK->GetItem(3,100))->ClearContents();

				int _tang = _wtoi(_ibdau);
				int _tong = _wtoi(_itong)+1;

				// Xóa bỏ các thiết lập vùng in
				//psNDNK->ResetAllPageBreaks();
				//xl->ActiveWindow->View = xlPageBreakPreview;

				// Kiểu in đứng nếu là biên bản thường
				psNDNK->PageSetup->Orientation = xlPortrait;

				// Thiết lập chế độ in
				_psp = psNDNK->PageSetup;
				f_Setup_Print(psNDNK,1);

				CString mss=L"";
				int _itrang,_ingang,_idoc;
				int xde,dem,_fsum,page_in,_ptram=0;
				if(_wtoi(numin)>0)
				{
					// Đường dẫn nhật ký
					pathNK = GIOR(luuHKN1,luuHKN2+2,prNDNK,L"");
					if(pathNK==L"") pathNK.Format(L"%sDulieu\\Nhatky.mdb",_fGPathFolder(L"ToolQLCL.dll"));
					if(connectDsn(pathNK)==-1) return;
					ObjConn.OpenAccessDB(pathNK,L"",L"");

					for (int t = 1; t <= _wtoi(numin); t++)
					{
						_itrang = _myPrint.ifirst;
						_ingang=1,_idoc=1;
						for (int i = 1; i <=_tong; i++)
						{
							if(i==_tong) 
								mss.Format(
								L"Đang tiến hành in dữ liệu (bản %d). Vui lòng đợi trong giây lát...(99%s)",t,L"%");
							else
							{
								_ptram = (int)(100*i/_tong);
								mss.Format(
									L"Đang tiến hành in dữ liệu (bản %d). Vui lòng đợi trong giây lát...(%d%s)",t,_ptram,L"%");
							}
							
							xl->PutStatusBar((_bstr_t)mss);

							// Tiến hành in
							msg.Format(L"%d",_tang+i);
							prNDNK->PutItem(i1,i2,(_bstr_t)msg);

							cls_main call_hk;
							call_hk.f_xuat_nhat_ky();							

							// Xác định lại ngày hiển thị
							_ngaythang = prNDNK->GetItem(luuHKN1,luuHKN2);

							// 11.06.2018
							// Kiểm tra có bỏ qua phần in không?
							msg = GIOR(luuHKN1+3,luuHKN2,prNDNK,L"");
							if(msg.Left(6)==L"Bỏ qua") continue;

							// Vị trí cuối cùng
							xde = FindComment(psNDNK,1,"END2");
							dem = 0;
							_fsum=0;
							for (int j = 1; j < xde; j++)
							{
								dem = psNDNK->GetRange(prNDNK->GetItem(j,1),prNDNK->GetItem(j,1))->RowHeight;
								_fsum+=dem;
							}
							page_in = (int)(_fsum*0.35/297)+5;

							// Đánh số trang
							if(_myPrint.inumpage==1)
							{
								if(_myPrint.ibreaks==1) _itrang = _myPrint.ifirst;
								else
								{
									psNDNK->PageSetup->FirstPageNumber = _itrang;
									_ingang = psNDNK->HPageBreaks->GetCount()+1;
									_idoc = psNDNK->VPageBreaks->GetCount()+1;
									_islpage = _ingang*_idoc;
									_itrang+=_islpage;
								}

								if(_myPrint.isecpage==0) psNDNK->PageSetup->RightFooter = (_bstr_t)L"&P";
								else psNDNK->PageSetup->CenterFooter = (_bstr_t)L"&P";
							}
			
							// Bổ sung phần xét vùng in (11.03.2016)
							_bstr_t _bsPr = psNDNK->PageSetup->GetPrintArea();
							CString _csPr = (LPCTSTR)_bsPr;
							_csPr.Trim();
							int ibdau=1;
							int ikthuc=5;
							if(_csPr!=L"")
							{
								_csPr.Replace(L"$",L"");
								int _pos = _csPr.Find(L":");
								if(_pos>0)
								{
									CString _cpr1 = _csPr.Left(_pos);
									CString _cpr2 = _csPr.Right(_csPr.GetLength()-_pos-1);

									// Cột bắt đầu
									PRS = psNDNK->GetRange((_bstr_t)_cpr1);
									ibdau = PRS->GetColumn();

									// Cột kết thúc
									PRS = psNDNK->GetRange((_bstr_t)_cpr2);
									ikthuc = PRS->GetColumn();
								}
							}

							// In hồ sơ
							psNDNK->GetRange(
								prNDNK->GetItem(1,ibdau),
								prNDNK->GetItem(xde,ikthuc))->PrintOut(xlPrintSheetEnd,_fsum,1);

							// Bổ sung lưu hình ảnh nhật ký (23.12.2016)
							f_in_sheetpic(psNDNK);
						}
					}

					ObjConn.CloseDatabase();

					_bstr_t shHAnh = name_sheet("shHinhAnh");
					_WorksheetPtr psHAnh = xl->Sheets->GetItem(shHAnh);
					int iVisible = psHAnh->GetVisible();
					if (iVisible!=2) psHAnh->PutVisible(XlSheetVisibility::xlSheetVeryHidden);

					iVisible = psNDNK->GetVisible();
					if (iVisible!=-1) psNDNK->PutVisible(XlSheetVisibility::xlSheetVisible);
					psNDNK->Activate();
				}
				else _s(L"Số bản in không phù hợp.");
			
			//xl->PutScreenUpdating(VARIANT_TRUE);
		}
	}
	else _s(L"Kiểm tra lại thời gian in nhật ký.");
	
	xl->PutStatusBar((_bstr_t)L"Ready");
	//CoUninitialize();
	}
	catch(_com_error & error){_s(L"#QL4.351: Lỗi in nội dung nhật ký.");}
}


void Dlg_print_nhatky::f_Load_data()
{
	TRY
	{
		CRecordset* Recset;
		CString val0=L"",val1=L"",val2=L"";
		SqlString.Format(L"SELECT * FROM tbAnh WHERE ngayghi = '%s' ORDER BY maanh ASC;",_ngaythang);
		Recset = ObjConn.Execute(SqlString);
		int slg=0,sdu=0;

		grPic=0;
		while( !Recset->IsEOF() )
		{
			Recset->GetFieldValue(L"hinhanh",val1);			
			val1.Trim();
			if(val1!=L"")
			{
				slg++;
				sdu = slg%_iptrang;
				if(sdu==1) grPic++;
				if(sdu==0) sdu= _iptrang;

				// Xác định kích thước hình ảnh (iWThuc x iHThuc)
				_fGetImage(val1);

				Mpic[grPic]._spt[sdu] = val1;
				Mpic[grPic].iWt[sdu] = iWThuc;
				Mpic[grPic].iHt[sdu] = iHThuc;

				// Ngày tháng
				Mpic[grPic]._sNgay = _ngaythang;

				// Nội dung
				Recset->GetFieldValue(L"noidung",val2);				
				Mpic[grPic]._sNdung[sdu] = val2;
			}

			Recset->MoveNext();
		}
		Recset->Close();

		// Gán NULL cho các giá trị ảnh không tồn tại
		if(sdu%_iptrang>0) for (int i = sdu+1; i <= _iptrang; i++) Mpic[grPic]._spt[i]=L"";
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("Lỗi tải dữ liệu hình ảnh")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_print_nhatky::f_in_sheetpic(_WorksheetPtr ps1)
{
	try
	{
		// Load dữ liệu xác định số lượng hình ảnh
		f_Load_data();
		if(grPic==0) return;

		_bstr_t shHAnh = name_sheet("shHinhAnh");
		_WorksheetPtr psHAnh = xl->Sheets->GetItem(shHAnh);
		RangePtr prHAnh = psHAnh->Cells;

		int iVisible = psHAnh->GetVisible();
		if (iVisible!=-1) psHAnh->PutVisible(XlSheetVisibility::xlSheetVisible);
		psHAnh->Activate();
		psHAnh->PageSetup->PutOrientation(xlPortrait);

		int dem=0,jgtMax=0,jWmax=0;
		CString val=L"";

///////// In ảnh cho từng nhóm

		if(_iptrang==2 || _iptrang==3)
		{
			// Cột
			jgtMax = 635-120*_iptrang;			// 275 và 395
			PRS = prHAnh->GetItem(1,1);
			PRS->PutColumnWidth(95);
			jWmax = PRS->GetWidth();

			// Tiêu đề			
			PRS = psHAnh->GetRange(prHAnh->GetItem(1,1),prHAnh->GetItem(2*_iptrang,1));
			PRS->PutVerticalAlignment(xlTop);
			PRS->PutWrapText(1);
		}
		else if(_iptrang==4)
		{
			// Cột
			jgtMax = 395;
			PRS = prHAnh->GetItem(1,1);
			PRS->PutColumnWidth(47);
			jWmax = PRS->GetWidth();

			// Cột 2
			PRS = prHAnh->GetItem(1,2);
			PRS->PutColumnWidth(47);

			// Tiêu đề
			PRS = psHAnh->GetRange(prHAnh->GetItem(1,1),prHAnh->GetItem(4,2));
			PRS->PutVerticalAlignment(xlTop);
			PRS->PutWrapText(1);
		}

		if(grPic>0)
		for (int i = 1; i <= grPic; i++)
		{
			dem=0;
			int _gcout = psHAnh->Shapes->Count;
			if(_gcout>=1)
			{
				// Xóa tất cả ảnh & control
				for (int j = _gcout; j >=1; j--) psHAnh->GetShapes()->Item(j)->Delete();

				// Xóa tất cả dữ liệu
				prHAnh->ClearContents();
			}

			if(_iptrang==2 || _iptrang==3)
			{
				for (int j = 1; j <= _iptrang; j++)
				{
					if(Mpic[i]._spt[j]!=L"")
					{
						dem++;
						val.Format(L"Ngày %s: %s",Mpic[i]._sNgay,Mpic[i]._sNdung[j]);
						prHAnh->PutItem(dem,1,(_bstr_t)val);
						PRS = prHAnh->GetItem(dem,1);
						PRS->EntireRow->Rows->AutoFit();
						int nChg = PRS->GetRowHeight();
						nChg+=5;
						PRS->PutRowHeight(nChg);

						dem++;
						if(Mpic[i].iWt[j]>0 && Mpic[i].iHt[j]>0)
						{
							int y1 = jgtMax-nChg;
							int x1 = (int)(y1*Mpic[i].iWt[j]/Mpic[i].iHt[j]);

							// Kiểm tra x1 có vượt quá jWmax không?
							if(x1>=jWmax)
							{
								x1 = jWmax-10;
								y1 = (int)(x1*Mpic[i].iHt[j]/Mpic[i].iWt[j]);
							}

							// Chèn ảnh vào biên bản & điều chỉnh kích thước theo (x1,y1) = (rộng,cao) đơn vị: pixcel
							int b1 = (int)((jWmax-x1)/2);
							PRS = prHAnh->GetItem(dem,1);
							PRS->PutRowHeight(y1+5);
							psHAnh->Shapes->AddPicture(
								(_bstr_t)Mpic[i]._spt[j],Office::msoTrue,Office::msoTrue,
								b1,PRS->Top,(float)x1,(float)y1);
						}
						else
						{
							val.Format(L"Vui lòng kiểm tra đường dẫn: %s",Mpic[i]._spt[j]);
							prHAnh->PutItem(dem,1,(_bstr_t)val);
						}
					}
				}
			}
			else if(_iptrang==4)
			{
				// Chạy từ trái qua phải, từ trên xuống dưới
				int cot=1,hang=2;
				for (int j = 1; j <= 4; j++)
				{
					if(j==1 || j==3) cot=1;
					else cot=2;

					if(j==1 || j==2) hang=1;
					else hang=3;					

					if(Mpic[i]._spt[j]!=L"")
					{
						val.Format(L"Ngày %s: %s",Mpic[i]._sNgay,Mpic[i]._sNdung[j]);
						prHAnh->PutItem(hang,cot,(_bstr_t)val);
					}
				}

				int nChg = 45;
				for (int j = 1; j <= 4; j++)
				{
					if(j==1 || j==3)
					{
						cot=1;
						PRS = prHAnh->GetItem(j,1);
						PRS->EntireRow->Rows->AutoFit();
						nChg = PRS->GetRowHeight();
						nChg+=5;
						PRS->PutRowHeight(nChg);
					}
					else cot=2;

					if(j==1 || j==2) hang=2;
					else hang=4;

					if(Mpic[i]._spt[j]!=L"")
					{
						if(Mpic[i].iWt[j]>0 && Mpic[i].iHt[j]>0)
						{
							int y1 = jgtMax-nChg;
							int x1 = (int)(y1*Mpic[i].iWt[j]/Mpic[i].iHt[j]);

							// Kiểm tra x1 có vượt quá jWmax không?
							if(x1>=jWmax)
							{
								x1 = jWmax-10;
								y1 = (int)(x1*Mpic[i].iHt[j]/Mpic[i].iWt[j]);
							}

							// PRS->Left : Căn trái, PRS->Top : Căn trên, Căn giữa: b1 = (int)((jWmax-x1)/2)+jWmax*(cot-1);
							// Chèn ảnh vào biên bản & điều chỉnh kích thước theo (x1,y1) = (rộng,cao) đơn vị: pixcel
							int b1 = (int)((jWmax-x1)/2)+jWmax*(cot-1);
							PRS = prHAnh->GetItem(hang,cot);
							PRS->PutRowHeight(y1+5);
							psHAnh->Shapes->AddPicture(
								(_bstr_t)Mpic[i]._spt[j],Office::msoTrue,Office::msoTrue,
								b1,PRS->Top,(float)x1,(float)y1);
						}
						else
						{
							val.Format(L"Vui lòng kiểm tra đường dẫn: %s",Mpic[i]._spt[j]);
							prHAnh->PutItem(hang,cot,(_bstr_t)val);
						}						
					}					
				}
			}

			// In hình ảnh
			psHAnh->PrintOut(xlPrintSheetEnd,100,1);
		}
		
//////////////////////////////////////////////////////////////////////

	}
	catch(...){_s(L"error-854 (2)");}
}


void Dlg_print_nhatky::OnBnClickedB3()
{
	try
	{
		// Hàm check bản quyền
		if(f_Check_ban_quyen()!=1) return;

		/*CoInitialize(NULL);
		xl.GetActiveObject(L"Excel.Application");
		xl->Visible = true;
		xl->EnableCancelKey = XlEnableCancelKey(FALSE);*/

		// Định nghĩa các sheet
		shNDNK = name_sheet("shMauNKY4");
		psNDNK=xl->Sheets->GetItem(shNDNK);
		prNDNK = psNDNK->Cells;

		BOOLEAN fx = TRUE;
		_iNgay=10;
		if(_ird==1)
		{
			// Xác định ngày bắt đầu & kết thúc (lựa chọn)
			GetDlgItemText(NKY_D1,ngay1);
			GetDlgItemText(NKY_D2,ngay2);
			_iNgay = ngay1.GetLength();

			// So sánh ngày tháng
			int f0 = compare_date(_iNgay,ngay1,ngay2);
			int f1 = compare_date(_iNgay,ngbd,ngay1);
			int f2 = compare_date(_iNgay,ngay2,ngkt);

			if(f0==1 && f1==1 && f2==1) 
			{
				fx = TRUE;
				ngbd = ngay1;
				ngkt = ngay2;
			}
			else fx = FALSE;
		}

		if(fx == TRUE)
		{
			psNDNK->GetRange(prNDNK->GetItem(2,98),prNDNK->GetItem(3,98))->PutNumberFormat("dd/mm/yyyy");
			psNDNK->GetRange(prNDNK->GetItem(2,99),prNDNK->GetItem(3,99))->PutNumberFormat("General");

			f_ktra_date(ngbd,prNDNK,2,98);
			f_ktra_date(ngkt,prNDNK,3,98);

			int i1 = getRow(psNDNK,"HK_NBD");
			int i2 = getColumn(psNDNK,"HK_NBD");

			msg.Format(L"=%s-%s",GAR(2,98,prNDNK,0),GAR(i1,i2,prNDNK,0));
			prNDNK->PutItem(2,99,(_bstr_t)msg);
			msg.Format(L"=%s-%s",GAR(3,98,prNDNK,0),GAR(2,98,prNDNK,0));
			prNDNK->PutItem(3,99,(_bstr_t)msg);

			CString _ibdau = prNDNK->GetItem(2,99);
			CString _itong = prNDNK->GetItem(3,99);

			msg = prNDNK->GetItem(2,99);
			i1 = getRow(psNDNK,"HK_CONG");
			i2 = getColumn(psNDNK,"HK_CONG"); 
			prNDNK->PutItem(i1,i2,(_bstr_t)msg);
			psNDNK->GetRange(prNDNK->GetItem(i1,i2),prNDNK->GetItem(i1,i2))->PutNumberFormat("General");
			psNDNK->GetRange(prNDNK->GetItem(2,98),prNDNK->GetItem(3,99))->ClearContents();

			int _tang = _wtoi(_ibdau);
			int _tong = _wtoi(_itong)+1;

			// Xóa bỏ các thiết lập vùng in
			psNDNK->ResetAllPageBreaks();
			//xl->ActiveWindow->View = xlPageBreakPreview;

			// Kiểu in đứng nếu là biên bản thường
			psNDNK->PageSetup->Orientation = xlPortrait;

			// Thiết lập chế độ in
			_psp = psNDNK->PageSetup;
			f_Setup_Print(psNDNK,1);

			erp = 0;
			int xde;
			for (int i = 1; i <=_tong; i++)
			{
				if(erp==1) return;

				// Tiến hành in
				msg.Format(L"%d",_tang+i);
				prNDNK->PutItem(i1,i2,(_bstr_t)msg);

				cls_main call_hk;
				call_hk.f_xuat_nhat_ky();

				xde = FindComment(psNDNK,1,"END2");

				// Xem trước nhật ký
				psNDNK->GetRange(prNDNK->GetItem(1,1),prNDNK->GetItem(xde,5))->PrintPreview();

				int result = MessageBox(L"Bạn có muốn dừng chế độ xem trước?",L"Thông báo", MB_YESNO | MB_ICONQUESTION);
				if (result == 6) erp = 1;
			}

			CDialog::ActivateTopParent();
		}
		else _s(L"Kiểm tra lại thời gian in nhật ký.");
	
		xl->PutStatusBar((_bstr_t)L"Ready");
		//CoUninitialize();
	}
	catch(_com_error & error){_s(L"#QL4.352: Lỗi xem trước nội dung nhật ký.");}
}


void Dlg_print_nhatky::OnBnClickedB1()
{
	iDlog=1;

	// Hiển thị hộp thoại tùy chọn
	AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
	Dlg_Setup_Print _dlg;
	_dlg.DoModal();

	CDialog::ActivateTopParent();
}


void Dlg_print_nhatky::OnBnClickedR1()
{
	_ird=0;
	dt1.EnableWindow(0);
	dt2.EnableWindow(0);
}


void Dlg_print_nhatky::OnBnClickedR2()
{
	_ird=1;
	dt1.EnableWindow(1);
	dt2.EnableWindow(1);
}


void Dlg_print_nhatky::OnBnClickedCk1()
{
	CButton *m_ctlCheck = (CButton*) GetDlgItem(NKY_CK1);
	if(m_ctlCheck->GetCheck() == BST_UNCHECKED) f_set_enable(0);
	else f_set_enable(1);
}


void Dlg_print_nhatky::f_set_enable(int type)
{
	rd3.EnableWindow(type);
	rd4.EnableWindow(type);
	rd5.EnableWindow(type);
	rd6.EnableWindow(type);
}

void Dlg_print_nhatky::OnEnKillfocusT1()
{
	iLen = txt1.LineLength(txt1.LineIndex(0));
	lp = numin.GetBuffer(iLen);
	txt1.GetLine(0, lp,iLen);
	numin.ReleaseBuffer();

	if(_wtoi(numin)<=0) txt1.SetWindowText(L"1");
	else if(_wtoi(numin)>100) txt1.SetWindowText(L"100");
	else
	{
		msg.Format(L"%d",_wtoi(numin));
		txt1.SetWindowText(msg);
	}
}


void Dlg_print_nhatky::OnBnClickedB4()
{
	try
	{
		// Lựa chọn máy in
		xl->StatusBar = (_bstr_t)L"Đang tiến hành chọn máy in. Vui lòng đợi trong giây lát..";
		xl->Application->Dialogs->GetItem(xlDialogPrinterSetup)->Show();
	
		//MessageBox(L"Bạn hãy lựa chọn hồ sơ và tiến hành công việc in ấn."
		//	L"\nNếu bạn không lựa chọn máy in, phần mềm sẽ lấy chế độ in mặc định.",L"Chú ý", MB_OK | MB_ICONINFORMATION);

		CDialog::ActivateTopParent();
		xl->StatusBar = (_bstr_t)L"Ready";
	}
	catch(_com_error & error){_s(L"#QL38: Lỗi lựa chọn máy in.");}
 }


void Dlg_print_nhatky::OnBnClickedB5()
{
	// Hàm check bản quyền
	if(f_Check_ban_quyen()!=1) return;

	_iptrang=0;
	_idkem = fdkem.GetCheck();
	if(_idkem==1) _iptrang = _irp;

	_iNgay=10;
	BOOLEAN fx = TRUE;

	if(_ird==1)
	{
		// Xác định ngày bắt đầu & kết thúc (lựa chọn)
		GetDlgItemText(NKY_D1,ngay1);
		GetDlgItemText(NKY_D2,ngay2);
		_iNgay = ngay1.GetLength();

		// So sánh ngày tháng
		int f0 = compare_date(_iNgay,ngay1,ngay2);
		int f1 = compare_date(_iNgay,ngbd,ngay1);
		int f2 = compare_date(_iNgay,ngay2,ngkt);

		if(f0==1 && f1==1 && f2==1) 
		{
			fx = TRUE;
			ngbd = ngay1;
			ngkt = ngay2;
		}
		else fx = FALSE;
	}

	if(fx == TRUE)
	{
		fx = FALSE;
		msg = L"";
		if(ngbd==ngkt) msg.Format(L"Bạn có chắc chắn lưu nhật ký ngày %s?",ngbd);
		else
		{
			if(_ird==0) msg.Format(L"Bạn có chắc chắn lưu TOÀN BỘ NHẬT KÝ \nTừ ngày %s đến ngày %s?",ngbd,ngkt);
			else msg.Format(L"Bạn có chắc chắn lưu nhật ký từ ngày %s đến %s?",ngbd,ngkt);
		}	

		int result = MessageBox(msg,_T("Thông báo"), MB_YESNO | MB_ICONQUESTION);
		if (result == 6) fx = TRUE;
		else fx = FALSE;

		if(fx == TRUE)
		{
			iLuuType=1;
			OnOK();

			// Hiển thị hộp thoại lưu hồ sơ
			AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
			Dlg_luu_Hso _dlg;
			_dlg.DoModal();
		}		
	}
	else _s(L"Vui lòng kiểm tra lại ngày nhập.");
}

void Dlg_print_nhatky::OnBnClickedR3()
{
	_irp=1;
}


void Dlg_print_nhatky::OnBnClickedR4()
{
	_irp=2;
}


void Dlg_print_nhatky::OnBnClickedR5()
{
	_irp=3;
}


void Dlg_print_nhatky::OnBnClickedR6()
{
	_irp=4;
}
