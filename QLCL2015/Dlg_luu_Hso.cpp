// Dlg_luu_Hso message handlers
#include "Dlg_luu_Hso.h"
#include "FolderDlg.h"
#include "cls_main.h"

Dlg_luu_Hso::Dlg_luu_Hso() : CDialog(Dlg_luu_Hso::IDD)
{
}

void Dlg_luu_Hso::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, LUUHS1_T1, Txt1);
	DDX_Control(pDX, LUUHS1_R1, Crd1);
	DDX_Control(pDX, LUUHS1_R2, Crd2);
	DDX_Control(pDX, LUUHS1_CK1, Chk0);
}

BEGIN_MESSAGE_MAP(Dlg_luu_Hso, CDialog)
	ON_BN_CLICKED(LUUHS1_R1, &Dlg_luu_Hso::OnBnClickedR1)
	ON_BN_CLICKED(LUUHS1_R2, &Dlg_luu_Hso::OnBnClickedR2)
	ON_BN_CLICKED(LUUHS1_CK1, &Dlg_luu_Hso::OnBnClickedCk1)
	ON_BN_CLICKED(LUUHS1_B1, &Dlg_luu_Hso::OnBnClickedB1)
	ON_BN_CLICKED(LUUHS1_B2, &Dlg_luu_Hso::OnBnClickedB2)
	ON_BN_CLICKED(LUUHS1_B3, &Dlg_luu_Hso::OnBnClickedB3)
END_MESSAGE_MAP()

BOOL Dlg_luu_Hso::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	iType=0;
	iChk=0;
	Crd1.SetCheck(1);
	if(iLuuType==1) Chk0.EnableWindow(0);

	Txt1.SetCueBanner(L"Browse...",TRUE);
	f_Load_ToolTip();

	if(sNKLuu==L"") sNKLuu.Format(L"%s%s",_fGPathFolder(L"ToolQLCL.dll"),L"Dulieu");
	Txt1.SetWindowText(sNKLuu);
	sPathF = sNKLuu;

	// Lấy thời gian hiện tại
	f_GetTimeNow();

	return TRUE; 
}


// Bắt sự kiện click Enter
BOOL Dlg_luu_Hso::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
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
	else if( pMsg->message == WM_LBUTTONDOWN ||
		pMsg->message == WM_LBUTTONUP ||
		pMsg->message == WM_MOUSEMOVE )
	{
		m_ToolTips.RelayEvent(pMsg);
		return CDialog::PreTranslateMessage(pMsg);
	}

	return FALSE;
}


void Dlg_luu_Hso::f_Load_ToolTip()
{
	CButton * btn1 = (CButton *)GetDlgItem(LUUHS1_R1);
	CButton * btn2 = (CButton *)GetDlgItem(LUUHS1_R2);
	CButton * btn3 = (CButton *)GetDlgItem(LUUHS1_CK1);

	m_ToolTips.Create(this);
	m_ToolTips.SetMaxTipWidth(600);
	m_ToolTips.SetDelayTime(900);

	m_ToolTips.AddTool(btn1, L"Định dạng kỹ thuật in ấn (XPS)."
		L"\nĐược tích hợp sẵn trong tất cả máy tính cá nhân.");

	m_ToolTips.AddTool(btn2, L"Định dạng tài liệu di động (PDF)."
		L"\n Cần phải cài đặt chương trình hỗ trợ đọc định dạng PDF.");

	m_ToolTips.AddTool(btn3, L"Tích chọn nếu bạn muốn phân chi tiết thư mục cho từng công việc.");
}


void Dlg_luu_Hso::f_GetTimeNow()
{
	time_t currentTime;
	struct tm *localTime;

	time( &currentTime );
	localTime = localtime( &currentTime );

	int iNow[6];
	iNow[0]  = (localTime->tm_year)%10;
	iNow[1]  = localTime->tm_mon + 1;
	iNow[2]  = localTime->tm_mday;
	iNow[3]  = localTime->tm_hour;
	iNow[4]  = localTime->tm_min;

	CString sL[6] = {L"",L"",L"",L"",L"",L""};
	for (int i = 1; i < 5; i++)
	{
		if(iNow[i]<10) sL[i].Format(L"0%d",iNow[i]);
		else sL[i].Format(L"%d",iNow[i]);
	}

	sRDay.Format(L"%d%s%s",iNow[0],sL[1],sL[2]);
	sRTime.Format(L"%s%s",sL[3],sL[4]);
}


void Dlg_luu_Hso::OnBnClickedR1()
{
	iType=0;
}


void Dlg_luu_Hso::OnBnClickedR2()
{
	iType=1;
	MessageBox(L"Máy tính của bạn cần cài chương trình hỗ trợ đọc file PDF. "
		L"\nViệc lưu file PDF có thể diễn ra lâu hơn so với file XPS.",L"Lưu ý",MB_OK | MB_ICONWARNING);
}

void Dlg_luu_Hso::OnBnClickedCk1()
{
	CButton *m_ctlCheck = (CButton*) GetDlgItem(LUUHS1_CK1);
	if(m_ctlCheck->GetCheck() == BST_UNCHECKED) iChk=0;
	else iChk=1;
}


BOOL Dlg_luu_Hso::Create_Folder(CString directoryPath)
{
	_demIn=0;
	 CString directoryPathClone = directoryPath;
	 LPTSTR lpPath = directoryPath.GetBuffer(MAX_PATH);
	 if (!PathFileExists(lpPath))
	 {
		if (!CreateDirectory(lpPath, NULL) && (GetLastError() == ERROR_PATH_NOT_FOUND))
		{
			if (PathRemoveFileSpec(lpPath))
			{
				CString newPath = lpPath;
				Create_Folder(newPath);
			}
		}
	 }

	directoryPath.ReleaseBuffer(-1);
	return PathFileExists(directoryPathClone);
}


void Dlg_luu_Hso::OnBnClickedB1()
{
	try
	{
		_demIn=0;
		CString val=L"";
		Txt1.GetWindowTextW(val);
		val.Trim();
		if(val==L"") val = sNKLuu;
		sCreate = val;

		if(iLuuType==1)
		{
			val.Format(L"%s\\NK%s%s",sCreate,sRDay,sRTime);
			Create_Folder(val);
			sCreate= val;

			// Lưu nhật ký
			xl_luu_nhat_ky();

			// Hiển thị thư mục đường dẫn
			_s(L"Lưu dữ liệu thành công!");
			ShellExecute(NULL, L"open",sCreate, NULL, NULL, SW_SHOWMAXIMIZED);

			return;
		}

		val.Format(L"%s\\HS%s%s",sCreate,sRDay,sRTime);
		Create_Folder(val);
		sCreate= val;

		sDM=L"",sBB=L"",sVL=L"",sCV=L"",sGD=L"";
		sDM.Format(L"%s\\%s",sCreate,L"DanhMuc");
		sBB.Format(L"%s\\%s",sCreate,L"BienBan");
		sVL.Format(L"%s\\%s\\%s",sCreate,L"BienBan",L"1.VatLieu");
		sCV.Format(L"%s\\%s\\%s",sCreate,L"BienBan",L"2.CongViec");
		sGD.Format(L"%s\\%s\\%s",sCreate,L"BienBan",L"3.GiaiDoan");

		CDialog::OnOK();
		xl->PutScreenUpdating(VARIANT_FALSE);
		xl->PutStatusBar((_bstr_t)L"Đang tiến hành lưu hồ sơ. Vui lòng chờ trong giây lát...");

		// Gán giá trị chèn ảnh
		shTS = name_sheet("shTS");
		psTS = xl->Sheets->GetItem(shTS);
		prTS = psTS->Cells;
		prTS->PutItem(getRow(psTS,"TS_PIC"),getColumn(psTS,"TS_PIC"),_ihidepic);

		_itrang = _myPrint.ifirst;		 // Trang bắt đầu khởi tạo
		if(luuhsCB1==0)
		{
			Create_Folder(sDM);
			Create_Folder(sBB);

			// In toàn bộ danh mục
			for (int j = 1; j <=4; j++) f_luu_danh_muc_hs(j);

			// Lưu toàn bộ hồ sơ
			Luu_toan_bo_ho_so();
		}
		else if(luuhsCB1>0 && luuhsCB1<4)
		{
			// Tạo thư mục biên bản
			Create_Folder(sBB);

			// Tạo thư mục hồ sơ
			CString sP1 = L"";
			if(luuhsCB1==1) sP1=sVL;
			else if(luuhsCB1==2) sP1=sCV;
			else if(luuhsCB1==3) sP1=sGD;
			Create_Folder(sP1);

			// Lưu toàn bộ 1 bộ hồ sơ nào đó (all biên bản)
			for(int j=1;j<=iLuuKEY;j++) f_Kiem_tra_inum(sP1,iCViec[j],luuhsCB1,luuhsCB2);
		}
		else if(luuhsCB1==4)
		{
			Create_Folder(sDM);

			// In toàn bộ danh mục
			if(luuhsCB2==0) for (int j = 1; j <=4; j++) f_luu_danh_muc_hs(j);
			else f_luu_danh_muc_hs(luuhsCB2);
		}

		// Trả lại giá trị ban đầu cho ảnh
		prTS->PutItem(getRow(psTS,"TS_PIC"),getColumn(psTS,"TS_PIC"),iLuuMDP);
		xl->PutStatusBar((_bstr_t)L"Ready");
		
		_s(L"Lưu dữ liệu thành công.");		

		// Hiển thị thư mục đường dẫn
		ShellExecute(NULL, L"open",sCreate, NULL, NULL, SW_SHOWMAXIMIZED);
	}
	catch(...){_s(L"Lỗi lưu hồ sơ [23.9]");}
}

// Edit 02.08
void Dlg_luu_Hso::xl_luu_nhat_ky()
{
	try
	{
		CDialog::OnOK();
		xl->PutScreenUpdating(VARIANT_FALSE);

		// Định nghĩa các sheet
		_bstr_t shNDNK = name_sheet("shMauNKY4");
		_WorksheetPtr psNDNK = xl->Sheets->GetItem(shNDNK);
		RangePtr prNDNK = psNDNK->Cells;

		psNDNK->GetRange(prNDNK->GetItem(2,99),prNDNK->GetItem(3,99))->PutNumberFormat("dd/mm/yyyy");
		psNDNK->GetRange(prNDNK->GetItem(2,100),prNDNK->GetItem(3,100))->PutNumberFormat("General");

		f_ktra_date(ngbd,prNDNK,2,99);
		f_ktra_date(ngkt,prNDNK,3,99);

		int i1 = getRow(psNDNK,"HK_NBD");
		int i2 = getColumn(psNDNK,"HK_NBD");

		int luuHKN1 = getRow(psNDNK,"HK_NGAY"); 
		int luuHKN2 = getColumn(psNDNK,"HK_NGAY");		

		CString msg = L"";
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

		CString sDuoi = _duoiXPS, sPrt = _prXPS, mss=L"", sAutoName=L"";
		if(iType==1)
		{
			sDuoi = _duoiPDF;	
			sPrt = _prPDF;
		}

		int _islpage,_itrang,_ingang,_idoc;
		int xde,dem,_fsum,page_in;
		int _ptram=0;

		_itrang = _myPrint.ifirst;
		_ingang=1,_idoc=1;

		// Đường dẫn nhật ký
		if(getPathNHKY()==L"") return;
		if(connectDsn(pathNK)==-1) return;
		ObjConn.OpenAccessDB(pathNK,L"",L"");

		for (int i = 1; i <=_tong; i++)
		{
			if(i==_tong) mss.Format(L"Đang tiến hành lưu dữ liệu. Vui lòng đợi trong giây lát...(99%s)",L"%");
			else
			{
				_ptram = (int)(100*i/_tong);
				mss.Format(L"Đang tiến hành lưu dữ liệu. Vui lòng đợi trong giây lát...(%d%s)",_ptram,L"%");
			}			
			xl->PutStatusBar((_bstr_t)mss);

			// Tiến hành in
			msg.Format(L"%d",_tang+i);
			prNDNK->PutItem(i1,i2,(_bstr_t)msg);

			cls_main call_hk;
			call_hk.f_xuat_nhat_ky();

			// Xác định ngày tháng
			_ngaythang = prNDNK->GetItem(luuHKN1,luuHKN2);

			// 11.06.2018
			// Kiểm tra có bỏ qua phần in không?
			msg = GIOR(luuHKN1+3,luuHKN2,prNDNK,L"");
			if(msg.Left(6)==L"Bỏ qua") continue;

			xde = FindComment(psNDNK,1,"END2");
			dem = 0,_fsum=0;
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

			CString sVal = GIOR(luuHKN1,luuHKN2,prNDNK,L"");
			if(sVal!=L"") sVal.Replace(L"/",L"-");

			sAutoName = L"";
			if(i<10) sAutoName.Format(L"%s\\NK00%d-ngay-%s%s",sCreate,i,sVal,sDuoi);
			else if(i>=10 && i<100) sAutoName.Format(L"%s\\NK0%d-ngay-%s%s",sCreate,i,sVal,sDuoi);
			else sAutoName.Format(L"%s\\NK%d-ngay-%s%s",sCreate,i,sVal,sDuoi);

			// Tiến hành in đến từ đầu đến trang gần cuối cùng
			CString sPrg=L"";
			sPrg.Format(L"%s:%s",GAR(1,ibdau,prNDNK,0),GAR(xde,ikthuc,prNDNK,0));
			if(iType==1) f_SavePDF(sAutoName,sPrg);
			else
			psNDNK->GetRange(
				prNDNK->GetItem(1,ibdau),
				prNDNK->GetItem(xde,ikthuc))
				->PrintOut(1,vtMissing,1,false,(_bstr_t)sPrt,true,false,(_bstr_t)sAutoName);

			// Bổ sung lưu hình ảnh nhật ký (23.12.2016)
			if(_idkem==1)
			{
				sAutoName.Replace(sDuoi,L"");
				sPrg.Format(L"%s-image",sAutoName);
				f_luu_sheetpic(psNDNK,sPrg,sDuoi);
			}

			_bstr_t shHAnh = name_sheet("shHinhAnh");
			_WorksheetPtr psHAnh = xl->Sheets->GetItem(shHAnh);
			int iVisible = psHAnh->GetVisible();
			if (iVisible!=2) psHAnh->PutVisible(XlSheetVisibility::xlSheetVeryHidden);

			iVisible = psNDNK->GetVisible();
			if (iVisible!=-1) psNDNK->PutVisible(XlSheetVisibility::xlSheetVisible);
			psNDNK->Activate();
		}

		ObjConn.CloseDatabase();
		xl->PutStatusBar((_bstr_t)L"Ready");
	}
	catch(...){_s(L"#QL6.08: Lỗi lưu nhật ký.");}
}


void Dlg_luu_Hso::f_Load_data()
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


void Dlg_luu_Hso::f_luu_sheetpic(_WorksheetPtr ps1,CString sth,CString sduoi)
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

			// Lưu hình ảnh
			if(i<10) val.Format(L"%s-0%d%s",sth,i,sduoi);
			else val.Format(L"%s-%d%s",sth,i,sduoi);

			CString sPrintArea=L"";
			sPrintArea.Format(L"A1:A%d",dem);
			if(_itrang==4) sPrintArea.Format(L"A1:B%d",dem);

			if(iType==1) f_SavePDF(val,sPrintArea);
			else psHAnh->PrintOut(1,vtMissing,1,false,(_bstr_t)sPrintArea,true,false,(_bstr_t)val);
		}
		
//////////////////////////////////////////////////////////////////////

	}
	catch(...){_s(L"error-854");}
}


void Dlg_luu_Hso::f_luu_danh_muc_hs(int icbb)
{
	// icbb = lựa chọn danh mục cần in
	// Các danh mục
	if(icbb == 1) Luu_chitiet_danhmuc((_bstr_t)"Sheet6");
	else if(icbb == 2) Luu_chitiet_danhmuc((_bstr_t)"shHSNTVL");
	else if(icbb == 3) Luu_chitiet_danhmuc((_bstr_t)"shHSNTCV");
	else if(icbb == 4) Luu_chitiet_danhmuc((_bstr_t)"shHSNTGD");
}


void Dlg_luu_Hso::Luu_chitiet_danhmuc(_bstr_t sPrint)
{
	// sPrint = sheet hiện hành
	
	try
	{
		shPrint = name_sheet(sPrint);
		psPrint=xl->Sheets->GetItem(shPrint);
		prPrint = psPrint->Cells;

		// Kiểm tra ẩn hiện sheet
		int iVisible = psPrint->GetVisible();
		if (iVisible!=-1) return;

		CString sDuoi = _duoiXPS;
		CString sPrt = _prXPS;
		if(iType==1)
		{
			sDuoi = _duoiPDF;
			sPrt = _prPDF;
		}

		CString val=L"",sPrg=L"";
		int iCol1,iCol2;

		// Xác định cột đầu- cuối trang in
		if(sPrint == (_bstr_t)"Sheet6")
		{
			// Danh mục hồ sơ
			iCol1 = getColumn(psPrint,"DMHS_STT");
			iCol2 = getColumn(psPrint,"DMHS_TTHS");

			val.Format(L"%s\\%s",sDM,L"1-HoSo");
		}
		else if(sPrint == (_bstr_t)"shHSNTVL")
		{
			// Danh mục vật liệu
			iCol1 = getColumn(psPrint,"DMVL_STT");
			iCol2 = 1+getColumn(psPrint,"DMVL_KBB");

			val.Format(L"%s\\%s",sDM,L"2-VatLieu");
		}
		else if(sPrint == (_bstr_t)"shHSNTCV")
		{
			// Danh mục công việc
			iCol1 = getColumn(psPrint,"DMBB_STT");
			iCol2 = getColumn(psPrint,"DMBB_PS");

			val.Format(L"%s\\%s",sDM,L"3-CongViec");
		}
		else if(sPrint == (_bstr_t)"shHSNTGD")
		{
			// Danh mục giai đoạn
			iCol1 = getColumn(psPrint,"DMGD_STT");
			iCol2 = getColumn(psPrint,"DMGD_KBB");

			val.Format(L"%s\\%s",sDM,L"4-GiaiDoan");
		}

		int xde = FindComment(psPrint,iCol1,"END");

		// Kiểu in nằm nếu là biên bản danh mục
		psPrint->PageSetup->Orientation = xlLandscape;

		// Thiết lập chế độ in
		_psp = psPrint->PageSetup;
		f_Setup_Print(psPrint,0);

		// Full path
		sPathF = val + sDuoi;
		sPrg.Format(L"%s:%s",GAR(1,iCol1,prPrint,0),GAR(xde,iCol2,prPrint,0));
		if(iType==1) f_SavePDF(sPathF,sPrg);
		else
		psPrint->GetRange(
			prPrint->GetItem(1,iCol1),
			prPrint->GetItem(xde,iCol2))
			->PrintOut(1,vtMissing,1,false,(_bstr_t)sPrt,true,false,(_bstr_t)sPathF);
	}
	catch(_com_error & error){_s(L"#QL6.01: Lỗi in chi tiết danh mục.");}
}


void Dlg_luu_Hso::f_Kiem_tra_bsung(CString sPt,CString inum)
{
	//try
	//{
	//	// Kiểm tra xem có 1 hay nhiều công việc?
	//	int _pos = inum.Find(L"-");
	//	if(_pos>0)
	//	{
	//		// Lựa chọn một chuỗi công việc liên tiếp
	//		int ibd = _wtoi(inum.Left(_pos));
	//		int ikt = _wtoi(inum.Right(inum.GetLength()-_pos-1));

	//		if(ibd>ikt)
	//		{
	//			int itg = ibd;
	//			ibd = ikt;
	//			ikt = itg;
	//		}

	//		// Tách chuỗi thành các công việc nhỏ
	//		for (int i = 1; i <= slbsThuc; i++)
	//			for (int j = ibd; j <= ikt; j++)
	//			{
	//				if(_myPrint.ibreaks==1 && _myPrint.inextpage==1) _itrang = _myPrint.ifirst;
	//				Luu_chitiet_bienban(csInbs[i],1,j);
	//			}
	//	}
	//	else
	//	{
	//		// Chỉ có 1 công việc được lựa chọn
	//		// Kiểm tra điều kiện ngắt trang khi thay đổi công việc
	//		for (int i = 1; i <= slbsThuc; i++)
	//		{
	//			if(_myPrint.ibreaks==1 && _myPrint.inextpage==1) _itrang = _myPrint.ifirst;
	//			Luu_chitiet_bienban(csInbs[i],1,_wtoi(inum));
	//		}
	//	}
	//}
	//catch(_com_error & error){_s(L"#QL6.01- Lỗi in hồ sơ 12.");}
}


void Dlg_luu_Hso::f_Kiem_tra_inum(CString sPt,CString inum,int icb1, int icb2)
{
	// inum = 1 công việc/ hoặc chuỗi công việc
	// icb1 - giá trị combobox1 - hồ sơ ( trong hàm này, icb1 luôn >0)
	// icb2 - giá trị combobox1 - biên bản
	
	try
	{
		CString val=L"";
		int ibd=0,ikt=0;

		// Kiểm tra xem có 1 hay nhiều công việc?
		int _pos = inum.Find(L"-");
		if(_pos<=0)
		{
			ibd = _wtoi(inum);
			ikt = ibd;
		}
		else
		{
			ibd = _wtoi(inum.Left(_pos));
			ikt = _wtoi(inum.Right(inum.GetLength()-_pos-1));

			if(ibd>ikt)
			{
				int itg = ibd;
				ibd = ikt;
				ikt = itg;
			}
		}

		// Tách chuỗi thành các công việc nhỏ
		for (int i = ibd; i <= ikt; i++)
		{
			if(iChk==1)
			{
				val=L"";
				val.Format(L"%s\\%d",sPt,i);
				Create_Folder(val);
			}

			if(_myPrint.ibreaks==1 && _myPrint.inextpage==1) _itrang = _myPrint.ifirst;

			if(icb2>0)
			{
				// đây là in lẻ 1 biên bản
				if(icb1==1)
				{
					if(icb2==1)
					{
						for (int j = 1; j <= isxin[0]; j++)
							if(SXI_VL_cn[j].Left(12)==L"shHSNTVLLAYM") f_lua_chon_bienban(i,SXI_VL_cn[j],icb1,icb2);
					}
					else if(icb2==2)
					{
						for (int j = 1; j <= isxin[0]; j++)
							if(SXI_VL_cn[j].Left(12)==L"shHSNTVLNBVL") f_lua_chon_bienban(i,SXI_VL_cn[j],icb1,icb2);
					}
					else if(icb2==3)
					{
						for (int j = 1; j <= isxin[0]; j++)
							if(SXI_VL_cn[j].Left(11)==L"shHSNTVLYVL") f_lua_chon_bienban(i,SXI_VL_cn[j],icb1,icb2);
					}
					if(icb2==4)
					{
						for (int j = 1; j <= isxin[0]; j++)
							if(SXI_VL_cn[j].Left(11)==L"shHSNTVLNVL") f_lua_chon_bienban(i,SXI_VL_cn[j],icb1,icb2);
					}
				}
				else if(icb1==2)
				{
					if(icb2==1)
					{
						for (int j = 1; j <= isxin[1]; j++)
							if(SXI_CV_cn[j].Left(12)==L"shHSNTCVLMTN") f_lua_chon_bienban(i,SXI_CV_cn[j],icb1,icb2);
					}
					else if(icb2==2)
					{
						for (int j = 1; j <= isxin[1]; j++)
						{
							if(SXI_CV_cn[j].Left(12)==L"shHSNTCVNBCV" 
								&& SXI_CV_cn[j].Left(13)!=L"shHSNTCVNBCV1" 
								&& SXI_CV_cn[j].Left(13)!=L"shHSNTCVNBCV2") f_lua_chon_bienban(i,SXI_CV_cn[j],icb1,icb2);
						}
					}
					else if(icb2==3)
					{
						for (int j = 1; j <= isxin[1]; j++)
							if(SXI_CV_cn[j].Left(12)==L"shHSNTCVYNCV") f_lua_chon_bienban(i,SXI_CV_cn[j],icb1,icb2);
					}
					if(icb2==4)
					{
						for (int j = 1; j <= isxin[1]; j++)
							if(SXI_CV_cn[j].Left(12)==L"shHSNTCVNTCV") f_lua_chon_bienban(i,SXI_CV_cn[j],icb1,icb2);
					}
					else if(icb2==5)
					{
						for (int j = 1; j <= isxin[1]; j++)
							if(SXI_CV_cn[j].Left(13)==L"shHSNTCVNBCV1") f_lua_chon_bienban(i,SXI_CV_cn[j],icb1,icb2);
					}
					if(icb2==6)
					{
						for (int j = 1; j <= isxin[1]; j++)
							if(SXI_CV_cn[j].Left(13)==L"shHSNTCVNBCV2") f_lua_chon_bienban(i,SXI_CV_cn[j],icb1,icb2);
					}
				}
				else if(icb1==3)
				{
					if(icb2==1)
					{
						for (int j = 1; j <= isxin[2]; j++)
							if(SXI_GD_cn[j].Left(14)==L"shHSNTGDNTBGD1") f_lua_chon_bienban(i,SXI_GD_cn[j],icb1,icb2);
					}
					else if(icb2==2)
					{
						for (int j = 1; j <= isxin[2]; j++)
							if(SXI_GD_cn[j].Left(13)==L"shHSNTGDYCNT1") f_lua_chon_bienban(i,SXI_GD_cn[j],icb1,icb2);
					}
					else if(icb2==3)
					{
						for (int j = 1; j <= isxin[2]; j++)
							if(SXI_GD_cn[j].Left(13)==L"shHSNTGDNTGD1") f_lua_chon_bienban(i,SXI_GD_cn[j],icb1,icb2);
					}
				}
			}
			else
			{
				// đây là in toàn bộ biên bản
				if(icb1==1) for (int j = 1; j <= isxin[0]; j++) f_lua_chon_bienban(i,SXI_VL_cn[j],icb1,0);
				else if(icb1==2) for (int j = 1; j <= isxin[1]; j++) f_lua_chon_bienban(i,SXI_CV_cn[j],icb1,0);
				else if(icb1==3) for (int j = 1; j <= isxin[2]; j++) f_lua_chon_bienban(i,SXI_GD_cn[j],icb1,0);
			}
		}
	}
	catch(_com_error & error){}
}


// Hàm lưu toàn bộ hồ sơ
void Dlg_luu_Hso::Luu_toan_bo_ho_so()
{
	try
	{
		CString val=L"";
		//int _ptram=0;

//////////// Lưu hồ sơ vật liệu
		if(slcv_VL>0)
		{
			Create_Folder(sVL);
			for(int j=1;j<=slcv_VL;j++)
			{
				//_ptram = (int)(100*j/slcv_VL);
				//val.Format(L"Đang tiến hành lưu các biên bản vật liệu...(%d%s)",_ptram,L"%");
				//xl->PutStatusBar((_bstr_t)val);

				if(iChk==1)
				{
					val=L"";
					val.Format(L"%s\\%d",sVL,j);
					Create_Folder(val);
				}

				if(_myPrint.ibreaks==1 && _myPrint.inextpage==1) _itrang = _myPrint.ifirst;
				for (int k = 1; k <= isxin[0]; k++) f_lua_chon_bienban(j,SXI_VL_cn[k],1,0);
			}
		}

//////////// Lưu hồ sơ công việc
		if(slcv_CV>0)
		{
			_demIn=0;
			Create_Folder(sCV);
			for(int j=1;j<=slcv_CV;j++)
			{
				//_ptram = (int)(100*j/slcv_CV);
				//val.Format(L"Đang tiến hành lưu các biên bản công việc...(%d%s)",_ptram,L"%");
				//xl->PutStatusBar((_bstr_t)val);

				if(iChk==1)
				{
					val=L"";
					val.Format(L"%s\\%d",sCV,j);
					Create_Folder(val);
				}

				if(_myPrint.ibreaks==1 && _myPrint.inextpage==1) _itrang = _myPrint.ifirst;
				for (int k = 1; k <= isxin[1]; k++) f_lua_chon_bienban(j,SXI_CV_cn[k],2,0);
			}
		}

//////////// Lưu hồ sơ giai đoạn
		if(slcv_GD>0)
		{
			_demIn=0;
			Create_Folder(sGD);
			for(int j=1;j<=slcv_GD;j++)
			{
				//_ptram = (int)(100*j/slcv_GD);
				//val.Format(L"Đang tiến hành lưu các biên bản giai đoạn... (%d%s)",_ptram,L"%");
				//xl->PutStatusBar((_bstr_t)val);

				if(iChk==1)
				{
					val=L"";
					val.Format(L"%s\\%d",sGD,j);
					Create_Folder(val);
				}

				if(_myPrint.ibreaks==1 && _myPrint.inextpage==1) _itrang = _myPrint.ifirst;
				for (int k = 1; k <= isxin[2]; k++) f_lua_chon_bienban(j,SXI_GD_cn[k],3,0);
			}
		}
	}
	catch(...){}
}

// Hàm lựa chọn biên bản (hàm mới - 13.03.2018)
// Nếu giá trị icb1=-1 hay icb2=-1 mặc định không sử dụng đến 1 trong 2 biến đó
void Dlg_luu_Hso::f_lua_chon_bienban(int num_cviec, CString sCodename, int icb1, int icb2)
{
	try
	{
		int ithem=1;
		CString sktr = sCodename;

		if(ixkoin>0)
		{
			for (int i = 1; i <= ixkoin; i++)
			{
				// Đây là các sheet nhân bản không được tích chọn khi in
				if(SXI_KO_cn[i]==sktr) return;
			}
		}

		if(icb1==1)
		{
			// Trường hợp đặc biệt, vật liệu chỉ in lấy mẫu khi có mẫu và tích chọn in lấy mẫu
			if(sktr.Left(12)==L"shHSNTVLLAYM" && iLuuChk[8]==1 && inLMVL[num_cviec]!=1) return;

			// Trường hợp đặc biệt check1 : không in nội bộ vật liệu
			if(sktr.Left(12)==L"shHSNTVLNBVL" && iLuuChk[1]!=0) return;

			// Trường hợp đặc biệt check5 : không in yêu cầu vật liệu
			if(sktr.Left(11)==L"shHSNTVLYVL" && iLuuChk[5]!=0) return;
		}
		else if(icb1==2)
		{
			// Trường hợp đặc biệt, lấy mẫu công việc chỉ in khi có mẫu (không in: inLMCV[i]=0)
			if(sktr.Left(12)==L"shHSNTCVLMTN" && inLMCV[num_cviec]!=1) return;

			// Trường hợp đặc biệt check2 : không in nội bộ công việc
			if(sktr.Left(12)==L"shHSNTCVNBCV" 
				&& sktr.Left(13)!=L"shHSNTCVNBCV1" 
				&& sktr.Left(13)!=L"shHSNTCVNBCV2" 
				&& iLuuChk[2]!=0) return;

			// Trường hợp đặc biệt check6 : không in yêu cầu công việc
			if(sktr.Left(12)==L"shHSNTCVYNCV" && iLuuChk[6]!=0) return;

			// Trường hợp đặc biệt, Kiểm tra bê tông chỉ in khi mã công việc = AF.x  <-- x = {1,2,3,4} hoặc AG (bổ sung thêm AG 08.03.2018)
			if(sktr.Left(13)==L"shHSNTCVNBCV1" && (inKTBT[num_cviec]!=1 || iLuuChk[4]!=0)) return;

			// Trường hợp đặc biệt, Theo dõi bê tông chỉ in khi có mẫu (không in: inLMCV[i]=0; in: inLMCV[i]=1;)
			if(sktr.Left(13)==L"shHSNTCVNBCV2" && (inLMCV[num_cviec]!=1 || iLuuChk[4]!=0)) return;
		}		

		if(sktr.Left(11)==L"shHSNTVLNVL" 
			|| sktr.Left(12)==L"shHSNTCVNTCV" 
			|| sktr.Left(13)==L"shHSNTGDNTGD1") ithem=5;

		// Kiểm tra đánh số thứ tự lại từ đầu
		if(_myPrint.ibreaks==1 && _myPrint.inextpage==2) _itrang = _myPrint.ifirst;

		//////////// Lưu biên bản ////////////////
		Luu_chitiet_bienban((_bstr_t)sCodename,ithem,num_cviec);
	}
	catch(...){}
}


// Hàm in bổ sung các sheet thêm mới
void Dlg_luu_Hso::Luu_toan_hs_bsung()
{
	/*int _ptram=0;
	CString mss=L"",val=L"";
	for (int i = 1; i <= slbsThuc; i++)
	{
		_ptram = (int)(100*i/slbsThuc);
		mss.Format(L"Đang tiến hành lưu các biên bản bổ sung...(%d%s)",_ptram,L"%");
		xl->PutStatusBar((_bstr_t)mss);

		for (int j = 1; j <= ctInbs[i]; j++)
		{
			if(_myPrint.ibreaks==1 && _myPrint.inextpage==1) _itrang = _myPrint.ifirst;
			Luu_chitiet_bienban(csInbs[i],1,i);
		}
	}*/
}


void Dlg_luu_Hso::Luu_chitiet_bienban(_bstr_t sPrint,int ithem,int num_cviec)
{
	// istyle = kiểu in/xem trước
	// sPrint = sheet hiện hành
	// ithem = số dòng thêm vào đuôi (nếu cần thiết)
	// num_cviec = số thứ tự công việc cần in

	try
	{
		CString szval=L"";

		shPrint = name_sheet(sPrint);
		psPrint=xl->Sheets->GetItem(shPrint);

		// Kiểm tra ẩn hiện sheet
		int iVisible = psPrint->GetVisible();
		if(iVisible!=-1) return;

		psPrint->Activate();
		prPrint = psPrint->Cells;

		CString sDuoi = _duoiXPS, sPrt = _prXPS;
		if(iType==1)
		{
			sDuoi = _duoiPDF;	
			sPrt = _prPDF;
		}

		f_CodeNameActivate(psPrint,num_cviec);		

		//Loading dll
		HINSTANCE loadDL = LoadLibrary(L"OpenCLGXD.dll");
			typedef void (__stdcall *p)(int opt);
			p call_Worksheet_OnChange;
			call_Worksheet_OnChange = (p)GetProcAddress(loadDL, "Worksheet_OnChange");
		//Loading dll------------/

		// Bổ sung phần xét vùng in (11.03.2016)
		_bstr_t _bsPr = psPrint->PageSetup->GetPrintArea();
		CString _csPr = (LPCTSTR)_bsPr;
		_csPr.Trim();
		int ibdau=3;
		int ikthuc=29;
		if(_csPr!=L"")
		{
			_csPr.Replace(L"$",L"");
			int _pos = _csPr.Find(L":");
			if(_pos>0)
			{
				CString _cpr1 = _csPr.Left(_pos);
				CString _cpr2 = _csPr.Right(_csPr.GetLength()-_pos-1);

				// Cột bắt đầu
				PRS = psPrint->GetRange((_bstr_t)_cpr1);
				ibdau = PRS->GetColumn();

				// Cột kết thúc
				PRS = psPrint->GetRange((_bstr_t)_cpr2);
				ikthuc = PRS->GetColumn();
			}
		}

		// Nhập giá trị công việc vào ô Cell "O5"
		prPrint->PutItem(_RSpin,_CSpin,num_cviec);

		//Run button spin----
		if(call_Worksheet_OnChange == NULL) MessageBox(L"Not found Worksheet_OnChange in OpenCLGXD.dll ", L"Error", 0);
		else call_Worksheet_OnChange(0);
		FreeLibrary(loadDL);
		//Run button spin----/

		xl->PutScreenUpdating(VARIANT_FALSE);

		// Xóa bỏ các thiết lập vùng in
		if(iLuuChk[3]==1) psPrint->ResetAllPageBreaks();
		//xl->ActiveWindow->View = xlPageBreakPreview;
		
		// Kiểu in đứng nếu là biên bản thường
		psPrint->PageSetup->Orientation = xlPortrait;

		// Thiết lập chế độ in
		_psp = psPrint->PageSetup;
		f_Setup_Print(psPrint,1);

		int xde1 = getRow(psPrint,_bs1);
		int xde2 = getRow(psPrint,_bs2)+ithem;
		
		// Độ dài khổ in thực của 1 trang giấy A4( top,bottom = 13mmm) = 297 - 2x13 =  271mm  xấp xỉ  775 point
		int _fsize = (int)((297-(_myPrint.itop+_myPrint.ibottom))/0.35+1);

		int _ipage[100];_ipage[1]=1;
		int dem = 0,_tong1=0,_tich=1;
		
		for (int i = 1; i < xde1; i++)
		{
			// dem = Độ cao của 1 dòng --> _tong1 = Tổng độ cao của bản in từ 1 đến dòng xde1
			dem = psPrint->GetRange(prPrint->GetItem(i,1),prPrint->GetItem(i,1))->RowHeight;
			_tong1+=dem;

			if(_tong1>(_fsize*_tich))
			{
				_tich++;			// Số trang in có thể có
				_ipage[_tich]=i;	// Vị trí cuối của mỗi trang in _ipage[1]=1;
			}
		}

		// _tong1 đơn vị là 'point'
		//int page_in = (int)(_tong1*0.35/297)+5;

		// Kiểm tra điều kiện check: tự động căn chỉnh khi in (iLuuChk[3]=1)
		// istyle=1 kiểu pview / = 0 kiểu print
		int _islpage,_ingang,_idoc;
		CString sPrg = L"", sAutoName = L"";
		
		// In hồ sơ
		if(iLuuChk[3]==1)
		{
			// Tự động căn chỉnh dòng
			// Kiểm tra nếu _tich>1 (_tich là tổng số trang có thể có) --> In từ trang 1 đến trang gần cuối cùng (_tich-1)

			if(_tich>1) 
			for (int i = 1; i <_tich; i++)
			{
				// Đánh số trang
				if(_myPrint.inumpage==1)
				{
					psPrint->PageSetup->FirstPageNumber = _itrang;
					_ingang = psPrint->HPageBreaks->GetCount()+1;
					_idoc = psPrint->VPageBreaks->GetCount()+1;
					_islpage = _ingang*_idoc-1;
					_itrang+=_islpage;

					if(_myPrint.isecpage==0) psPrint->PageSetup->RightFooter = (_bstr_t)L"&P";
					else psPrint->PageSetup->CenterFooter = (_bstr_t)L"&P";
				}

				sAutoName = L"";
				sAutoName.Format(L"%s-%d%s",sPathF,i,sDuoi);

				// Tiến hành in đến từ đầu đến trang gần cuối cùng
				sPrg.Format(L"%s:%s",GAR(_ipage[i],ibdau,prPrint,0),GAR(_ipage[i+1]-1,ikthuc,prPrint,0));
				if(iType==1) f_SavePDF(sAutoName,sPrg);
				else
				psPrint->GetRange(
					prPrint->GetItem(_ipage[i],ibdau),
					prPrint->GetItem(_ipage[i+1]-1,ikthuc))
					->PrintOut(1,vtMissing,1,false,(_bstr_t)sPrt,true,false,(_bstr_t)sAutoName);
			}

			// Kiểm tra trang cuối cùng có bị chia đôi dòng chữ ký không?
			int _tong2=0;
			for (int i = _ipage[_tich]; i < xde2; i++)
			{
				dem = psPrint->GetRange(prPrint->GetItem(i,1),prPrint->GetItem(i,1))->RowHeight;
				_tong2+=dem;
			}

			// Nếu chia đôi dòng chữ ký --> Tách làm 2 trang riêng, ngược lại chỉ in 1 trang
			if(_tong2>_fsize)
			{
				// Đánh số trang
				if(_myPrint.inumpage==1)
				{
					psPrint->PageSetup->FirstPageNumber = _itrang;
					_itrang++;

					if(_myPrint.isecpage==0) psPrint->PageSetup->RightFooter = (_bstr_t)L"&P";
					else psPrint->PageSetup->CenterFooter = (_bstr_t)L"&P";
				}

				// In trang bị tách phần trên
				sAutoName = L"";
				sAutoName.Format(L"%s-%d%s",sPathF,_tich,sDuoi);
				sPrg.Format(L"%s:%s",GAR(_ipage[_tich],ibdau,prPrint,0),GAR(xde1-1,ikthuc,prPrint,0));
				if(iType==1) f_SavePDF(sAutoName,sPrg);
				else
				psPrint->GetRange(
					prPrint->GetItem(_ipage[_tich],ibdau),
					prPrint->GetItem(xde1-1,ikthuc))
					->PrintOut(1,vtMissing,1,false,(_bstr_t)sPrt,true,false,(_bstr_t)sAutoName);
					
				// Đánh số trang
				if(_myPrint.inumpage==1)
				{
					psPrint->PageSetup->FirstPageNumber = _itrang;
					_itrang++;

					if(_myPrint.isecpage==0) psPrint->PageSetup->RightFooter = (_bstr_t)L"&P";
					else psPrint->PageSetup->CenterFooter = (_bstr_t)L"&P";
				}

				// In trang bị tách phần dưới
				sAutoName = L"";
				sAutoName.Format(L"%s-%d%s",sPathF,_tich+1,sDuoi);
				sPrg.Format(L"%s:%s",GAR(xde1,ibdau,prPrint,0),GAR(xde2,ikthuc,prPrint,0));
				if(iType==1) f_SavePDF(sAutoName,sPrg);
				else
				psPrint->GetRange(
					prPrint->GetItem(xde1,ibdau),
					prPrint->GetItem(xde2,ikthuc))
					->PrintOut(1,vtMissing,1,false,(_bstr_t)sPrt,true,false,(_bstr_t)sAutoName);
			}
			else
			{
				// Đánh số trang
				if(_myPrint.inumpage==1)
				{
					psPrint->PageSetup->FirstPageNumber = _itrang;
					_itrang++;

					if(_myPrint.isecpage==0) psPrint->PageSetup->RightFooter = (_bstr_t)L"&P";
					else psPrint->PageSetup->CenterFooter = (_bstr_t)L"&P";
				}

				// Trường hợp này, dòng chữ ký không bị chia đôi
				sAutoName = L"";
				sAutoName.Format(L"%s-%d%s",sPathF,_tich,sDuoi);
				sPrg.Format(L"%s:%s",GAR(_ipage[_tich],ibdau,prPrint,0),GAR(xde2,ikthuc,prPrint,0));
				if(iType==1) f_SavePDF(sAutoName,sPrg);
				else
				psPrint->GetRange(
					prPrint->GetItem(_ipage[_tich],ibdau),
					prPrint->GetItem(xde2,ikthuc))
					->PrintOut(1,vtMissing,1,false,(_bstr_t)sPrt,true,false,(_bstr_t)sAutoName);
			}
		}
		else
		{
			// Không tích tùy chọn tự động căn chỉnh
			// Đánh số trang
			if(_myPrint.inumpage==1)
			{
				psPrint->PageSetup->FirstPageNumber = _itrang;
				_ingang = psPrint->HPageBreaks->GetCount()+1;
				_idoc = psPrint->VPageBreaks->GetCount()+1;
				_islpage = _ingang*_idoc;
				_itrang+=_islpage;

				if(_myPrint.isecpage==0) psPrint->PageSetup->RightFooter = (_bstr_t)L"&P";
				else psPrint->PageSetup->CenterFooter = (_bstr_t)L"&P";
			}

			sAutoName = sPathF + sDuoi;
			sPrg.Format(L"%s:%s",GAR(1,ibdau,prPrint,0),GAR(xde2,ikthuc,prPrint,0));
			if(iType==1) f_SavePDF(sAutoName,sPrg);
			else
			psPrint->GetRange(
				prPrint->GetItem(1,ibdau),
				prPrint->GetItem(xde2,ikthuc))
				->PrintOut(1,vtMissing,1,false,(_bstr_t)sPrt,true,false,(_bstr_t)sAutoName);
		}

		szval.Format(L"Lưu thành công biên bản số %d - %s",num_cviec,(LPCTSTR)shPrint);
		xl->PutStatusBar((_bstr_t)szval);
	}
	catch(_com_error & error){}
}


void Dlg_luu_Hso::f_CodeNameActivate(_WorksheetPtr ps1,int num_cv)
{
	try
	{
		_demIn++;
		CString snum=L"";
		if(_demIn<=9) snum.Format(L"00%d",_demIn);
		else if(_demIn>9 && _demIn<=99) snum.Format(L"0%d",_demIn); 
		else snum.Format(L"%d",_demIn); 

		int nTong=0,nLen=0;
		CString sGan1=L"", val = L"";
		_bstr_t _pSht = ps1->CodeName;
		CString sEdit = (LPCTSTR)_pSht;

		if(sEdit.Left(12)==L"shHSNTVLLAYM")
		{
			nTong=12;

			// Lấy mẫu vật liệu
			_RSpin = getRow(ps1,"MAUVL_BB");
			_CSpin = getColumn(ps1,"MAUVL_BB");

			_bs2=L"PR_LMVL";
			_bs1 =L"PR_LMVL2";

			val=sVL;
			if(iChk==1) val.Format(L"%s\\%d\\VL-%d-%s-LayMau",sVL,num_cv,num_cv,snum);
			else val.Format(L"%s\\VL-%d-%s-LayMau",sVL,num_cv,snum);
		}
		else if(sEdit.Left(12)==L"shHSNTVLNBVL")
		{
			nTong=12;

			// Nội bộ vật liệu
			_RSpin = getRow(ps1,"NTNBVL_BB");
			_CSpin = getColumn(ps1,"NTNBVL_BB");

			_bs2=L"PR_NTNBVL";
			_bs1 =L"PR_NTNBVL2";

			val=sVL;
			if(iChk==1) val.Format(L"%s\\%d\\VL-%d-%s-NoiBo",sVL,num_cv,num_cv,snum);
			else val.Format(L"%s\\VL-%d-%s-NoiBo",sVL,num_cv,snum);
		}
		else if(sEdit.Left(11)==L"shHSNTVLYVL")
		{
			nTong=11;

			// Yêu cầu vật liệu
			_RSpin = getRow(ps1,"YCVL_BB");
			_CSpin = getColumn(ps1,"YCVL_BB");

			_bs2=L"PR_YCVL";
			_bs1 =L"PR_YCVL2";

			val=sVL;
			if(iChk==1) val.Format(L"%s\\%d\\VL-%d-%s-YeuCau",sVL,num_cv,num_cv,snum);
			else val.Format(L"%s\\VL-%d-%s-YeuCau",sVL,num_cv,snum);
		}
		else if(sEdit.Left(11)==L"shHSNTVLNVL")
		{
			nTong=11;

			// Nghiệm thu vật liệu
			_RSpin = getRow(ps1,"NTVL_BB");
			_CSpin = getColumn(ps1,"NTVL_BB");

			_bs2=L"PR_NTVL";
			_bs1 =L"PR_NTVL2";

			val=sVL;
			if(iChk==1) val.Format(L"%s\\%d\\VL-%d-%s-NghiemThu",sVL,num_cv,num_cv,snum);
			else val.Format(L"%s\\VL-%d-%s-NghiemThu",sVL,num_cv,snum);
		}
		else if(sEdit.Left(12)==L"shHSNTCVLMTN")
		{
			nTong=12;

			// Lấy mẫu công việc
			_RSpin = getRow(ps1,"NLM_BB");
			_CSpin = getColumn(ps1,"NLM_BB");

			_bs2=L"PR_LMTN";
			_bs1 =L"PR_LMTN2";

			val=sCV;
			if(iChk==1) val.Format(L"%s\\%d\\CV-%d-%s-LayMau",sCV,num_cv,num_cv,snum);
			else val.Format(L"%s\\CV-%d-%s-LayMau",sCV,num_cv,snum);
		}
		else if(sEdit.Left(12)==L"shHSNTCVNBCV" 
			&& sEdit.Left(13)!=L"shHSNTCVNBCV1" 
			&& sEdit.Left(13)!=L"shHSNTCVNBCV2")
		{
			// Lưu ý nguy hiểm: codename sheet nội bộ công việc này được đặt khá giống với 2 sheet ở dưới là
			// kiểm tra bê tông & theo dõi bê tông (codename = shHSNTCVNBCV)
			nTong=12;

			// Nội bộ công việc
			_RSpin = getRow(ps1,"NTNBCVGD_BB");
			_CSpin = getColumn(ps1,"NTNBCVGD_BB");

			_bs2=L"PR_NTNB";
			_bs1 =L"PR_NTNB2";

			val=sCV;
			if(iChk==1) val.Format(L"%s\\%d\\CV-%d-%s-NoiBo",sCV,num_cv,num_cv,snum);
			else val.Format(L"%s\\CV-%d-%s-NoiBo",sCV,num_cv,snum);
		}
		else if(sEdit.Left(13)==L"shHSNTCVNBCV1")
		{
			nTong=13;

			// Kiểm tra bê tông
			_RSpin = getRow(ps1,"NTNB_BB");
			_CSpin = getColumn(ps1,"NTNB_BB");

			_bs2=L"PR_BBKT";
			_bs1 =L"PR_BBKT2";

			val=sCV;
			if(iChk==1) val.Format(L"%s\\%d\\CV-%d-%s-KiemTraBT",sCV,num_cv,num_cv,snum);
			else val.Format(L"%s\\CV-%d-%s-BBKiemTra",sCV,num_cv,snum);
		}
		else if(sEdit.Left(13)==L"shHSNTCVNBCV2")
		{
			nTong=13;

			// Theo dõi bê tông
			_RSpin = getRow(ps1,"NTNBGD_BB");
			_CSpin = getColumn(ps1,"NTNBGD_BB");

			_bs2=L"PR_BBTD";
			_bs1 =L"PR_BBTD2";

			val=sCV;
			if(iChk==1) val.Format(L"%s\\%d\\CV-%d-%s-TheoDoiBT",sCV,num_cv,num_cv,snum);
			else val.Format(L"%s\\CV-%d-%s-TheoDoiBT",sCV,num_cv,snum);
		}
		else if(sEdit.Left(12)==L"shHSNTCVYNCV")
		{
			nTong=12;

			// Yêu cầu công việc
			_RSpin = getRow(ps1,"YCNTCV_BB");
			_CSpin = getColumn(ps1,"YCNTCV_BB");

			_bs2=L"PR_YCCV";
			_bs1 =L"PR_YCCV2";

			val=sCV;
			if(iChk==1) val.Format(L"%s\\%d\\CV-%d-%s-YeuCau",sCV,num_cv,num_cv,snum);
			else val.Format(L"%s\\CV-%d-%s-YeuCau",sCV,num_cv,snum);
		}
		else if(sEdit.Left(12)==L"shHSNTCVNTCV")
		{
			nTong=12;

			// Nghiệm thu công việc
			_RSpin = getRow(ps1,"NTCV_BB");
			_CSpin = getColumn(ps1,"NTCV_BB");

			_bs2=L"PR_NTCV";
			_bs1 =L"PR_NTCV2";

			val=sCV;
			if(iChk==1) val.Format(L"%s\\%d\\CV-%d-%s-NghiemThu",sCV,num_cv,num_cv,snum);
			else val.Format(L"%s\\CV-%d-%s-NghiemThu",sCV,num_cv,snum);
		}
		else if(sEdit.Left(14)==L"shHSNTGDNTBGD1")
		{
			nTong=14;

			// Nội bộ giai đoạn
			_RSpin = getRow(ps1,"NTNBGDG_BB");
			_CSpin = getColumn(ps1,"NTNBGDG_BB");
			
			_bs2=L"PR_NTNBGD";
			_bs1 =L"PR_NTNBGD2";

			val=sGD;
			if(iChk==1) val.Format(L"%s\\%d\\CV-%d-%s-NoiBo",sGD,num_cv,num_cv,snum);
			else val.Format(L"%s\\CV-%d-%s-NoiBo",sGD,num_cv,snum);
		}
		else if(sEdit.Left(13)==L"shHSNTGDYCNT1")
		{
			nTong=13;

			// Yêu cầu giai đoạn
			_RSpin = getRow(ps1,"YCGD_BB");
			_CSpin = getColumn(ps1,"YCGD_BB");

			_bs2=L"PR_YCGD";
			_bs1 =L"PR_YCGD2";

			val=sGD;
			if(iChk==1) val.Format(L"%s\\%d\\CV-%d-%s-YeuCau",sGD,num_cv,num_cv,snum);
			else val.Format(L"%s\\CV-%d-%s-YeuCau",sGD,num_cv,snum);
		}
		else if(sEdit.Left(13)==L"shHSNTGDNTGD1")
		{
			nTong=13;

			// Nghiệm thu giai đoạn
			_RSpin = getRow(ps1,"NTGD_BB");
			_CSpin = getColumn(ps1,"NTGD_BB");

			_bs2=L"PR_NTGD";
			_bs1 =L"PR_NTGD2";

			val=sGD;
			if(iChk==1) val.Format(L"%s\\%d\\CV-%d-%s-NghiemThu",sGD,num_cv,num_cv,snum);
			else val.Format(L"%s\\CV-%d-%s-NghiemThu",sGD,num_cv,snum);
		}

		// Kiểm tra trường hợp in bổ sung
		nLen = sEdit.GetLength();
		if(nLen>nTong)
		{
			sGan1 = val;
			val.Format(L"%s-BS%s",sGan1,sEdit.Right(nLen-nTong));
		}

		// Full path (chưa thể gán đuôi vì có thể sử dụng tính năng tự động căn chỉnh)
		sPathF = val;
	}
	catch(_com_error & error){_s(L"#QL6.86: Lỗi xác định codename sheet.");}
}


void Dlg_luu_Hso::OnBnClickedB2()
{
	OnCancel();
	return;
}


void Dlg_luu_Hso::OnBnClickedB3()
{
	try
	{
		CString m_szDuongDan=L"";
		LPCTSTR lpszTitle = _T( "Chọn đường dẫn dữ liệu" );
		UINT	uFlags	  = BIF_RETURNONLYFSDIRS | BIF_USENEWUI;
		CFolderDialog dlgRoot( lpszTitle, m_szDuongDan, this, uFlags );
		if( dlgRoot.DoModal() == IDOK )	
		{
			m_szDuongDan=dlgRoot.GetFolderPath();
			Txt1.SetWindowText(m_szDuongDan);
		}
	}
	catch(...){_s(L"Lỗi chọn đường dẫn.");}
}
