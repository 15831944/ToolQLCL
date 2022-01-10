#include "Dlg_Chen_dong.h"
#include "Dg_In_ho_so.h"
#include "Dlg_importpic.h"
#include "Dlg_print_nhatky.h"
#include "Dlg_trauu_tieuchuan.h"
#include "Dlg_Luu_DMCV.h"
#include "cls_main.h"
#include "Dlg_manager_wb.h"
#include "Dlg_Printer.h"
#include "Dlg_ThuvienTL.h"
#include "Dlg_sort_dmhs.h"
#include "Dlg_Qly_image_nky.h"
#include "SeachCG.h"
#include "Dlg_TH_Cptuvan.h"
#include "Dlg_Timer.h"
#include "Dlg_TimeDM.h"
#include "Dlg_Sapxepcviec.h"
#include "Dlg_XuatbangKlg.h"
#include "Dlg_Quanly_NDNK.h"
#include "Dlg_ktra_ngaynghi.h"
#include "Dlg_AutoSTT_HSNT.h"
#include "Dlg_QCLMBetong.h"		// LaymauNgang
#include "Dlg_Themtieuchuan.h"
#include "Dlg_NameSheet.h"
#include "Dlg_listNKyMDB.h"
#include "Dlg_NhomkyBB.h"
#include "Dlg_NhomBangmau.h"
#include "Dlg_ChonCSDLVLCV.h"
#include "frm_QuycachOld.h"

// Gọi hộp thoại in hồ sơ
void __stdcall open_in_HoSo()
{
	try
	{
		xl_Create();

		shTS = name_sheet((_bstr_t)"shTS");
		psTS = xl->Sheets->GetItem(shTS);
		prTS = psTS->Cells;

		jInle=0;
		AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
		Dg_In_ho_so _dlg;
		_dlg.DoModal();

		CoUninitialize();
	}
	catch(...){}
}


void __stdcall open_inleHoso()
{
	try
	{
		xl_Create();

		shTS = name_sheet((_bstr_t)"shTS");
		psTS = xl->Sheets->GetItem(shTS);
		prTS = psTS->Cells;

		// Xác định vùng công việc được lựa chọn
		_WorksheetPtr pSh1 = pWb->ActiveSheet;
		RangePtr pRg1 = pSh1->Cells;
		_bstr_t sh = pSh1->CodeName;

		if(sh==(_bstr_t)L"shHSNTVL" || sh==(_bstr_t)L"shHSNTCV" || sh==(_bstr_t)L"shHSNTGD")
		{
			int sl=0;
			int numcv[1000];
			CString szval=L"";
			RangePtr PRS = pSh1->Application->Selection;
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
						szval = pRg1->GetItem(_wtoi(_arrSLS[i]),1);
						if(_wtoi(szval)>0)
						{
							sl++;
							numcv[sl] = _wtoi(szval);
						}
					}
					else
					{
						// Có nhiều lựa chọn
						int jlen = _arrSLS[i].GetLength();
						int jbd = _wtoi(_arrSLS[i].Left(_pos));
						int jkt = _wtoi(_arrSLS[i].Right(jlen-_pos-1));
						
						for (int j = jbd; j <=jkt; j++)
						{
							szval = pRg1->GetItem(j,1);
							if(_wtoi(szval)>0)
							{
								PRS = pRg1->GetItem(j,1);
								int ic = PRS->GetRowHeight();
								if(ic>0)
								{
									sl++;
									numcv[sl] = _wtoi(szval);
								}								
							}
						}
					}
				}
			}

			sTukhoa=L"";
			if(sl>0)
			{
				sTukhoa.Format(L"%d",numcv[1]);
				if(sl>1) for (int i = 2; i <= sl; i++) sTukhoa.Format(L"%s,%d",sTukhoa,numcv[i]);
			}

			jInle=1;
			/*AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
			Dg_In_ho_so _dlg;
			_dlg.DoModal();*/

			// Leo edit 17.04.2019
			putVal(psTS,(_bstr_t)L"CF_INLEHOSO",(_bstr_t)sTukhoa);
			
			typedef bool (__stdcall *p)();
			HINSTANCE loadDLL = LoadLibrary(L"UtilityQLCL9.dll");
			if(loadDLL!=NULL)
			{
				p pc = (p)GetProcAddress(loadDLL,"OpenFrmInhoso");
				if(pc!=NULL) pc();
			}
			FreeLibrary(loadDLL);

			putVal(psTS,(_bstr_t)L"CF_INLEHOSO",(_bstr_t)L"");
		}
		else _s(L"Bạn cần lựa chọn công việc cần in tại bảng danh mục vật liệu/công việc.");			

		CoUninitialize();
	}
	catch(...){}
}

// Gọi hộp thoại kiểm tra ngày nghỉ lễ
void __stdcall open_ktra_nghile()
{
	try
	{
		xl_Create();

		shTS = name_sheet((_bstr_t)"shTS");
		psTS = xl->Sheets->GetItem(shTS);
		prTS = psTS->Cells;

		AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
		Dlg_ktra_ngaynghi _dlg;
		_dlg.DoModal();	

		CoUninitialize();
	}
	catch(...){}
}


// Gọi hộp thoại tổng hợp chi phí tư vấn
void __stdcall open_chiphituvan()
{
	try
	{
		xl_Create();

		_WorksheetPtr pSh1 = pWb->ActiveSheet;
		_bstr_t sh = pSh1->CodeName;
		int _icol = pSh1->Application->ActiveCell->Column;
		if(sh==(_bstr_t)L"shTHCPTV" && _icol==3)
		{
			int iEnd = FindComment(pSh1,2,"END");
			int _irow = pSh1->Application->ActiveCell->Row;

			if(_irow>=5 && _irow<iEnd)
			{
				AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
				Dlg_TH_Cptuvan _dlg;
				_dlg.DoModal();	
			}			
		}			

		CoUninitialize();
	}
	catch(...){}
}


// Gọi hộp thoại sắp xếp vật liệu / công việc
void __stdcall open_in_sapxep(long n)
{
	try
	{
		xl_Create();

		shTS = name_sheet((_bstr_t)"shTS");
		psTS = xl->Sheets->GetItem(shTS);
		prTS = psTS->Cells;

		// n=0 sheet vật liệu - n=1 sheet công việc
		_bstr_t bst = L"shHSNTVL";
		if(n==1) bst = L"shHSNTCV";

		_bstr_t sh = name_sheet(bst);
		_WorksheetPtr pSheet = xl->Sheets->GetItem(sh);

		// Kiểm tra ẩn hiện -> activate
		int iVisible = pSheet->GetVisible();
		if (iVisible!=-1) pSheet->PutVisible(XlSheetVisibility::xlSheetVisible);
		pSheet->Activate();

		AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
		Dlg_Sapxepcviec _dlg;
		_dlg.DoModal();	

		CoUninitialize();
	}
	catch(...){}
}


// Gọi hộp thoại xuất bảng khối lượng
void __stdcall open_xuat_klg()
{
	try
	{
		xl_Create();

		shTS = name_sheet((_bstr_t)"shTS");
		psTS = xl->Sheets->GetItem(shTS);
		prTS = psTS->Cells;

		_bstr_t sh = name_sheet((_bstr_t)L"shHSNTCV");
		_WorksheetPtr pSheet = xl->Sheets->GetItem(sh);

		// Kiểm tra ẩn hiện -> activate
		int iVisible = pSheet->GetVisible();
		if (iVisible!=-1) pSheet->PutVisible(XlSheetVisibility::xlSheetVisible);
		pSheet->Activate();

		AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
		Dlg_XuatbangKlg _dlg;
		_dlg.DoModal();	

		CoUninitialize();
	}
	catch(...){}
}


// Gọi hộp thoại quản lý nội dung nhật ký
void __stdcall call_open_qlndNK()
{
	xl_Create();

	AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
	Dlg_Quanly_NDNK _dlg;
	_dlg.DoModal();

	CoUninitialize();
}


// Gọi hộp thoại in nhật ký
void __stdcall open_in_NhatKy()
{
	xl_Create();

	shTS = name_sheet((_bstr_t)"shTS");
	psTS = xl->Sheets->GetItem(shTS);
	prTS = psTS->Cells;

	AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
	Dlg_print_nhatky _dlg;
	_dlg.DoModal();

	CoUninitialize();
}


//In bảng tính khác
void __stdcall  call_printer()
{
	xl_Create();

	shTS = name_sheet((_bstr_t)"shTS");
	psTS = xl->Sheets->GetItem(shTS);
	prTS = psTS->Cells;

	AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
	Dlg_Printer open_dlg;
	open_dlg.DoModal();

	CoUninitialize();
}

// Quản lý Wb
void __stdcall  call_manager_sheet()
{
	xl_Create();

	AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
	Dlg_Manager_wb open_dlg;
	open_dlg.DoModal();

	CoUninitialize();
}

// Gọi hộp thoại chèn dòng
void __stdcall open_chen_dong()
{
	xl_Create();

	iKqua=0;	// lấy tạm biến
	AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
	Dlg_Chen_dong _dlg;
	_dlg.DoModal();

	CoUninitialize();
}

// Gọi hộp thoại chèn ảnh
void __stdcall open_importpic(CString macv_)
{
	//AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
	//Dlg_importpic _dlg;
	//_dlg.DoModal();
}

// Gọi hộp thoại tra cứu tiêu chuẩn
void __stdcall open_tracuu_tieuchuan()
{
	Dlg_trauu_tieuchuan _dlg;
	_dlg.xl_timkiem_tieuchuan();
}

// Gọi hộp thoại lưu danh mục công việc
void __stdcall open_luu_dmcviec()
{
	Dlg_Luu_DMCV _dlg;
	_dlg.f_kiemtra_sheet();
}

// Gọi hàm gọi dữ liệu sang danh mục biên bản
void __stdcall call_move_dmbban()
{
	xl_Create();
	xl->PutScreenUpdating(VARIANT_FALSE);

	shTS = name_sheet("shTS");
	psTS=xl->Sheets->GetItem(shTS);
	prTS=psTS->GetCells();

	cls_main _fcall;
	_fcall.f_select_gtri();

	xl->StatusBar = (_bstr_t)L"Ready";
	CoUninitialize();
}

// Gọi hàm nhập nhật ký
// chạy khi nhấn spin hoặc quản lý > xuất số liệu nhật ký
void __stdcall call_nhap_hoi_ky()
{
	CoInitialize(NULL);
	xl.GetActiveObject(L"Excel.Application");
	xl->Visible = true;
	xl->PutScreenUpdating(VARIANT_FALSE);
	xl->StatusBar = (_bstr_t)L"Đang tiến hành xuất dữ liệu nhật ký. Vui lòng đợi trong giây lát...";
	pWb = xl->GetActiveWorkbook();
		
	cls_main call_hk;
	call_hk.f_xuat_nhat_ky();

	xl->StatusBar = (_bstr_t)L"Ready";
	CoUninitialize();
}

// Gọi hàm ghi truyền dữ liệu sang sheet Vật liệu nhập về
void __stdcall call_truyen_shVL_nhapve()
{
	cls_main _fcall;
	_fcall.f_truyen_dulieu_VLNV();
}

// Gọi hàm autofilter
void __stdcall call_autoFilter_VLNV()
{
	cls_main _fcall;
	_fcall.f_autoFilter_VLNV();
}

// Gọi hàm load dữ liệu sang các sheet
void __stdcall call_load_stt_bienban()
{
	cls_main _fcall;
	_fcall.f_load_stt_bienban();
}

// Gọi hàm click cập nhật số thứ tự công việc
void __stdcall call_capnhat_stt()
{
	cls_main _fcall;
	_fcall.f_capnhat_stt();
}

// Gọi hàm ghi truyền dữ liệu sang sheet Danh mục hồ sơ
void __stdcall call_truyenDL_sang_sh_dmHoSo()
{
	xl_Create();
	xl->PutScreenUpdating(VARIANT_TRUE);

	isktrHS=0;
	AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
	Dlg_sort_dmhs _dlg;
	_dlg.DoModal();
	
	if(isktrHS==1)
	{
		cls_main _fcall;
		_fcall.iCheckTenTT = _dlg.iCheckTenTT;
		_fcall.f_truyen_dulieu_DMHS();
	}
}

	

// Gọi hàm chèn ảnh
void __stdcall call_imppic()
{
	cls_main _fcall;
	_fcall.f_importspic();
}

// Gọi hàm chèn link ảnh lên sheet danh mục giai đoạn
void __stdcall call_implinkpicsheet()
{
	cls_main _fcall;
	_fcall.f_implinkpicsheet();
}

// Gọi hàm tự động thêm ngày tháng khi click Enter
void __stdcall call_auto_date_bienban()
{
	//cls_main call_dbb;
	//call_dbb.auto_date_bienban();
}

// không sử dụng?
void __stdcall call_import_dmvl()
{
	cls_main call_dbb;
	call_dbb.f_FindBrowse();
}

// Tải dữ liệu vật liệu từ file EXCEL
void __stdcall call_import_xlsx()
{
	cls_main call_dbb;
	call_dbb.f_Load_file_vatlieu();
}

// Tải dữ liệu công việc từ file EXCEL
void __stdcall call_import_cviec()
{
	cls_main call_dbb;
	call_dbb.f_Load_file_congviec();
}

// Đồng bộ danh mục vật liệu hoặc công việc
void __stdcall call_db_dmanhmuc()
{
	cls_main call_dbb;
	call_dbb.f_xacdinh_danhmuc();
}

// Hàm kích enter trong bảng tính
void __stdcall call_even_enter()
{
	cls_main call_enter;
	call_enter.f_even_enter();
}

// Hàm gọi các hộp thoại tra cứu dữ liệu
void __stdcall call_open_tracuu_dl()
{
	// Lấy luôn hàm call_enter ở trên
	cls_main call_enter;
	call_enter.f_even_enter();
}

// Hàm chuột phải hiển thị chi phí chuyên gia
void __stdcall call_rcl_cpcgia()
{
	xl_Create();

	_WorksheetPtr pSh1 = pWb->ActiveSheet;
	_bstr_t sh = pSh1->CodeName;
	int _icol = pSh1->Application->ActiveCell->Column;
	if(sh==(_bstr_t)L"shCPCGia" && _icol==4)
	{
		int _irow = pSh1->Application->ActiveCell->Row;
		int iEnd = FindCEnd(pSh1,2,"END",_irow+2);
		if(iEnd==0)
		{
			CoUninitialize();
			return;
		}

		// Xác định vị trí đang activate thuộc Man Month thứ mấy?
		iVMMn=0;
		int iBDAU=6;
		CString szval=L"";
		RangePtr pRg1 = pSh1->Cells;
		for (int i = 3; i < _irow; i++)
		{
			szval = pRg1->GetItem(i,2);
			if(szval==L"STT")
			{
				iBDAU=i+1;
				iVMMn++;
			}
		}

		if(_irow>=iBDAU && _irow<iEnd)
		{
			AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
			SeachCG _dlg;
			_dlg.DoModal();	
		}			
	}			

	CoUninitialize();
}

// Lấy codename (A Dũng)
_variant_t __stdcall Get_CodeName()
{
	_variant_t __stdcall sh;
	try
	{
		CoInitialize(NULL);
		_ApplicationPtr _pExl;
		HRESULT hr = _pExl.GetActiveObject(L"Excel.Application");
		if(FAILED(hr) && _pExl == NULL) return NULL;
		_pExl->PutScreenUpdating(VARIANT_FALSE);
		_WorkbookPtr pWB1 = _pExl->GetActiveWorkbook();
		_WorksheetPtr pSh1 = pWB1->ActiveSheet;
		sh = pSh1->CodeName;
		
		return (_variant_t)sh;
	}
	catch(...){}
	CoUninitialize();
	return (_variant_t)sh;
}


void __stdcall call_save_nky()
{
	try
	{
		CoInitialize(NULL);
		xl.GetActiveObject(L"Excel.Application");
		xl->Visible = true;
		xl->PutScreenUpdating(VARIANT_FALSE);
		pWb = xl->GetActiveWorkbook();

		_WorksheetPtr pSh1 = pWb->ActiveSheet;
		_bstr_t sh = pSh1->CodeName;
		if(sh==(_bstr_t)L"shMauNKY4")
		{
			cls_main call_luu;
			call_luu.f_luu_noidung();
		}
		else _s(L"Tính năng được sử dụng tại sheet nhật ký.");		

		xl->StatusBar = (_bstr_t)L"Ready";
		CoUninitialize();
	}
	catch(...){_s(L"[2] Lỗi lưu dữ liệu nhật ký!");}
}


void __stdcall call_addimage_nky()
{
	try
	{
		CoInitialize(NULL);
		xl.GetActiveObject(L"Excel.Application");
		xl->Visible = true;
		xl->PutScreenUpdating(VARIANT_FALSE);
		pWb = xl->GetActiveWorkbook();
	
		_WorksheetPtr pSh1 = pWb->ActiveSheet;
		_bstr_t sh = pSh1->CodeName;
		if(sh==(_bstr_t)L"shMauNKY4")
		{
			AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
			Dlg_Qly_image_nky _dlg;
			_dlg.DoModal();			
		}
		else _s(L"Tính năng được sử dụng tại sheet nhật ký.");		

		xl->StatusBar = (_bstr_t)L"Ready";
		CoUninitialize();
	}
	catch(...){_s(L"[1] Lỗi lưu dữ liệu nhật ký!");}
}


/*HWND hWndxl*/
void __stdcall call_get_date()
{
	try
	{
		//hWndExcel = hWndxl;
		xl_Create();

		shTS = name_sheet("shTS");
		psTS=xl->Sheets->GetItem(shTS);
		prTS=psTS->GetCells();

		_bstr_t _bsh1 = name_sheet((_bstr_t)"shMauNKY4");
		_WorksheetPtr pSh1 = xl->Sheets->GetItem(_bsh1);

		// Xác định định dạng dd/mm/yyyy (10) hay dd/mm/yy (8)
		 CString szval = getVCell(pSh1,L"HK_NBD");
		_iNgay = szval.GetLength();

		pSh1 = pWb->ActiveSheet;
		RangePtr pRg1 = pSh1->Cells;

		int _irow = pSh1->Application->ActiveCell->Row;
		int _icol = pSh1->Application->ActiveCell->Column;
		
		szval = GIOR(_irow,_icol,pRg1,L"");
		szval.Trim();
		if(szval!=L"") _bNThang=szval;

		_iTimeDlg=0;
		AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
		Dlg_Timer _dlg;
		_dlg.DoModal();	

		CoUninitialize();
	}
	catch(...){_s(L"Lỗi xác định ngày tháng");}
}


void __stdcall call_get_time()
{
	try
	{
		xl_Create();

		shTS = name_sheet((_bstr_t)"shTS");
		psTS = xl->Sheets->GetItem(shTS);
		prTS = psTS->Cells;

		AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
		Dlg_TimeDM _dlg;
		_dlg.DoModal();	

		CoUninitialize();
	}
	catch(...){_s(L"Lỗi xác định thời gian");}
}


void __stdcall call_hien_an_DL()
{
	try
	{
		CoInitialize(NULL);
		xl.GetActiveObject(L"Excel.Application");
		xl->Visible = true;
		xl->PutScreenUpdating(VARIANT_FALSE);
		pWb = xl->GetActiveWorkbook();

		_WorksheetPtr pSheet = pWb->ActiveSheet;
		CString sh1 = (LPCTSTR)(_bstr_t)pSheet->CodeName;

		if (sh1!=L"shHSNTCV" && sh1!=L"shHSNTVL" && sh1!=L"shDTXD")
		{
			_s(L"Tính năng sử dụng tại bảng danh mục công việc hoặc vật liệu hoặc biểu giá hợp đồng");
			return;
		}

		RangePtr PRS;
		RangePtr pRange = pSheet->Cells;
		int iID = 1, iMaCV = 2, irowbegin=9;

		if (sh1==L"shHSNTCV")
		{
			iID = getColumn(pSheet,L"DMBB_STT");
			iMaCV = getColumn(pSheet,L"DMBB_MACV");
		}
		else if (sh1==L"shHSNTVL")
		{
			iID = getColumn(pSheet,L"DMVL_STT");
			iMaCV = getColumn(pSheet,L"DMVL_MAVL");
		}
		else
		{
			irowbegin=7;
			iID = getColumn(pSheet,L"BIEUGIA_ID");
			iMaCV = getColumn(pSheet,L"BIEUGIA_MHDG");
		}

		CString szval=L"";
		int iEnd = FindComment(pSheet,iID,"END");
		for (int i = iEnd-1; i >=irowbegin; i--)
		{
			PRS = pRange->GetItem(i,1);
			if((int)PRS->GetRowHeight()==0)
			{
				PRS = pSheet->GetRange(pRange->GetItem(irowbegin,1),pRange->GetItem(iEnd,1));
				PRS->EntireRow->PutHidden(false);
				break;
			}
			else
			{
				PRS = pRange->GetItem(i,1);
				szval = GIOR(i,iMaCV,pRange,L"");
				if(szval==L"") PRS->EntireRow->PutHidden(true);
				else PRS->EntireRow->PutHidden(false);
			}
		}

		CoUninitialize();
	}
	catch(...){}
}


// n=1 check hiển thị các sheet chi phí tư vấn
void __stdcall call_sheet_bsdtoan(long n)
{
	try
	{
		CoInitialize(NULL);
		xl.GetActiveObject(L"Excel.Application");
		xl->Visible = true;
		xl->PutScreenUpdating(VARIANT_FALSE);
		pWb = xl->GetActiveWorkbook();

		int pos;
		_bstr_t cd,sh1,sh2;
		_WorksheetPtr psheet_,pSh1;
		_bstr_t bsx[6] = {"shDTMM","shCPCGia","shCPTLuong","shCPKTV","shTDCViec","shTHCPTV"};

		// Xác định sheet đang activate
		if(n==0)
		{
			// Ẩn các sheet tư vấn, hiển thị các sheet mặc định
			if(slSheet>0)
			{
				// Hiển thị các sheet mặc định
				for (int i=1;i<=slSheet;i++)
				{
					psheet_ = pWb->Worksheets->Item[i];
					pos = psheet_->GetVisible();
					if(pos!=iHideS[i])
					{
						if (iHideS[i]==-1) psheet_->PutVisible(XlSheetVisibility::xlSheetVisible);
						else if (iHideS[i]==0) psheet_->PutVisible(XlSheetVisibility::xlSheetHidden);
						else psheet_->PutVisible(XlSheetVisibility::xlSheetVeryHidden);
					}					
				}
			}
			else
			{
				// Hiển thị toàn bộ các sheet
				for (int i=1;i<=slSheet;i++)
				{
					psheet_ = pWb->Worksheets->Item[i];
					psheet_->PutVisible(XlSheetVisibility::xlSheetVisible);				
				}
			}

			// Ẩn các sheet tư vấn
			for (int i = 0; i < 6; i++)
			{
				sh1 = name_sheet(bsx[i]);
				psheet_ = xl->Sheets->GetItem(sh1);
				psheet_->PutVisible(XlSheetVisibility::xlSheetHidden);
			}

			sh1 = name_sheet(bsActivate);
			psheet_ = xl->Sheets->GetItem(sh1);
			psheet_->PutVisible(XlSheetVisibility::xlSheetVisible);	
			psheet_->Activate();
		}
		else
		{
			pSh1 = pWb->ActiveSheet;
			bsActivate = pSh1->CodeName;
			if(bsActivate==bsx[0] || bsActivate==bsx[1] || bsActivate==bsx[2] || bsActivate==bsx[3] 
				|| bsActivate==bsx[4] || bsActivate==bsx[5]) bsActivate = (_bstr_t)"shHSNTCV";

			// Hiển thị các sheet tư vấn
			for (int i = 0; i < 6; i++)
			{
				sh1 = name_sheet(bsx[i]);
				pSh1 = xl->Sheets->GetItem(sh1);
				pSh1->PutVisible(XlSheetVisibility::xlSheetVisible);
			}

			// Ẩn toàn bộ các sheet còn lại
			// Xác định trạng thái ẩn hiện của sheet			
			slSheet = xl->ActiveWorkbook->Sheets->Count;
			for (int i=1;i<=slSheet;i++)
			{
				psheet_ = pWb->Worksheets->Item[i];
				iHideS[i] = psheet_->GetVisible();
				if(iHideS[i]!=0)
				{
					cd = psheet_->CodeName;
					if(cd!=bsx[0] && cd!=bsx[1] && cd!=bsx[2] && cd!=bsx[3] && cd!=bsx[4] && cd!=bsx[5])
						psheet_->PutVisible(XlSheetVisibility::xlSheetHidden);
				}
			}

			// Sheets = Số sheet để cuộn qua. Sử dụng một số dương để di chuyển về phía trước (bên phải),
			// số âm để cuộn ngược lại, hoặc = 0 để không cuộn. Bạn phải chỉ định bảng tính nếu bạn không chỉ định vị trí
			// Position = Sử dụng xlFirst để cuộn đến trang đầu tiên, hoặc sử dụng xlLast để cuộn đến trang cuối cùng.
			pSh1->Activate();
			xl->ActiveWindow->ScrollWorkbookTabs(8);
		}		

		CoUninitialize();
	}
	catch(...)
	{
		CoInitialize(NULL);
		xl.GetActiveObject(L"Excel.Application");
		xl->Visible = true;
		xl->PutScreenUpdating(VARIANT_FALSE);
		pWb = xl->GetActiveWorkbook();
		_WorksheetPtr psheet_;	

		slSheet = xl->ActiveWorkbook->Sheets->Count;
		for (int i=1;i<=slSheet;i++)
		{
			psheet_ = pWb->Worksheets->Item[i];
			psheet_->PutVisible(XlSheetVisibility::xlSheetVisible);
			iHideS[i] = -1;
		}

		CoUninitialize();
	}
}


void __stdcall call_create_dtmm()
{
	try
	{
		CoInitialize(NULL);
		xl.GetActiveObject(L"Excel.Application");
		xl->Visible = true;
		xl->PutScreenUpdating(VARIANT_FALSE);
		pWb = xl->GetActiveWorkbook();

		_WorksheetPtr pSheet = pWb->ActiveSheet;
		_bstr_t sh1 = pSheet->CodeName;

		int pos=-1;
		if (sh1!=(_bstr_t)L"shTHCPTV")
		{
			_s(L"Sử dụng tính năng này tại bảng tổng hợp chi phí tư vấn.");

			sh1 = name_sheet("shTHCPTV");
			pSheet = xl->Sheets->GetItem(sh1);
			pos = pSheet->GetVisible();
			if (pos!=-1) pSheet->PutVisible(XlSheetVisibility::xlSheetVisible);
			pSheet->Activate();

			CoUninitialize();
			return;
		}

		// Xác định tên gọi phương thức lập dự toán Man Month (mặc định sTenMM = "Lập dự toán riêng")
		shTS = name_sheet("shTS");	
		psTS = xl->Sheets->GetItem(shTS);
		prTS = psTS->Cells;

		CString sTenMM = getVCell(psTS,L"CF_LDTR");

		// Kiểm tra có phải là lập dự toán Man Month không?
		int _irow = pSheet->Application->ActiveCell->Row;
		int iEnd = FindCEnd(pSheet,2,"END",_irow+1);
		if(_irow<=4 || iEnd==0)
		{
			CoUninitialize();
			return;
		}

		RangePtr pRange = pSheet->Cells;
		CString _szval = GIOR(_irow,2,pRange,L"");
		_szval.Trim();
		if(_szval==L"")
		{
			CoUninitialize();
			return;
		}

		_szval = GIOR(_irow,6,pRange,L"");
		_szval.Trim();
		if(_szval!=sTenMM)
		{
			if(_yn(L"Bạn muốn chuyển sang lập dự toán Man Month?",0)==7)
			{
				CoUninitialize();
				return;
			}

			// Gán lập dự toán riêng
			pRange->PutItem(_irow,6,(_bstr_t)sTenMM);
		}

		// Gán giá trị
		wsTHCPtv = pSheet;
		iMMRow = _irow;

		// iVtri sẽ là vị trí bắt đầu đổ dữ liệu
		int iVtri=1;

		// Kiểm tra đã tồn tại bảng dự toán Man Month chưa?		
		_szval = pRange->GetItem(_irow,12);
		_szval.Trim();
		if(_szval!=L"")
		{
			if(_yn(L"Vị trí này đã tồn tại bảng tính dự toán Man Month. \nBạn có muốn xóa và thay thế?",0)!=6)
			{
				CoUninitialize();
				return;
			}

			// Xóa các bảng tính... 
			_bstr_t bx = pRange->GetItem(_irow,2);
			if(bx!=(_bstr_t)L"")
			{
				cls_main _fcall;
				_fcall.f_Delete_table(bx);
			}
		}

		_WorksheetPtr wSht[6];
		_bstr_t bsx[6] = {"shDTMM","shCPCGia","shCPTLuong","shCPKTV","shTDCViec","shTHCPTV"};

		// Hiển thị sheet (nếu cần)
		for (int i = 0; i < 6; i++)
		{
			sh1 = name_sheet(bsx[i]);
			wSht[i] = xl->Sheets->GetItem(sh1);
			pos = wSht[i]->GetVisible();
			if (pos!=-1) wSht[i]->PutVisible(XlSheetVisibility::xlSheetVisible);
		}

		// 1. Tạo bảng tính mới - DT Man Month
		cls_main _fcall;
		_mmvt[5] = _fcall.f_Create_table(wSht[4],5,300);	// Tiến độ (để tạm cột giới hạn = 300)
		_mmvt[4] = _fcall.f_Create_table(wSht[3],4,5);		// CPK
		_mmvt[3] = _fcall.f_Create_table(wSht[2],4,11);		// Tiền lương
		_mmvt[2] = _fcall.f_Create_table(wSht[1],4,8);		// CP chuyên gia
		_mmvt[1] = _fcall.f_Create_table(wSht[0],4,7);		// Man Month

		// Liên kết link giữa các bảng tính
		_fcall.f_Link_manmonth(wSht[0],wSht[1],wSht[3]);

		// Gán giá trị khởi tạo Man Month thành công
		pRange->PutItem(_irow,12,1);

		// Activate vào vị trí
		RangePtr PRS = pRange->GetItem(_irow,9);
		PRS->Activate();

		_s(L"Lập dự toán Man Month thành công!");
		CoUninitialize();
	}
	catch (...){_s(L"Lỗi tạo bảng chi phí tư vấn");}
}


void __stdcall call_delete_dtmm()
{
	try
	{
		CoInitialize(NULL);
		xl.GetActiveObject(L"Excel.Application");
		xl->Visible = true;
		xl->PutScreenUpdating(VARIANT_FALSE);
		pWb = xl->GetActiveWorkbook();

		_WorksheetPtr pSheet = pWb->ActiveSheet;
		_bstr_t sh1 = pSheet->CodeName;

		int pos=-1;
		if (sh1!=(_bstr_t)L"shTHCPTV")
		{
			_s(L"Sử dụng tính năng này tại bảng tổng hợp chi phí tư vấn.");

			sh1 = name_sheet("shTHCPTV");
			pSheet = xl->Sheets->GetItem(sh1);
			pos = pSheet->GetVisible();
			if (pos!=-1) pSheet->PutVisible(XlSheetVisibility::xlSheetVisible);
			pSheet->Activate();

			CoUninitialize();
			return;
		}

		// Xóa các bảng tính...
		int _irow = pSheet->Application->ActiveCell->Row;
		int iEnd = FindCEnd(pSheet,2,"END",_irow+1);
		if(_irow<=4 || iEnd==0)
		{
			CoUninitialize();
			return;
		}

		RangePtr pRange = pSheet->Cells;
		_bstr_t bx = pRange->GetItem(_irow,2);
		if(bx!=(_bstr_t)L"")
		{
			cls_main _fcall;
			_fcall.f_Delete_table(bx);

			// Xóa dữ liệu
			RangePtr PRS = pRange->GetItem(_irow,1);
			PRS->EntireRow->Delete(xlShiftUp);

			_s(L"Đã xóa nội dung chi phí tư vấn.");
		}
		
		CoUninitialize();
	}
	catch (...){_s(L"Lỗi xóa bảng chi phí tư vấn");}
}


void __stdcall open_danh_ma_HSNT()
{
	xl_Create();

	_WorksheetPtr pSheet = pWb->ActiveSheet;
	_bstr_t sh= pSheet->CodeName;

	if (sh==(_bstr_t)L"shHSNTVL" || sh==(_bstr_t)L"shHSNTCV")
	{
		AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
		Dlg_AutoSTT_HSNT open_dlg;
		open_dlg.DoModal();
	}
	else _s(L"Sử dụng tính năng này tại danh mục vật liệu hoặc công việc");

	CoUninitialize();
}


void __stdcall call_update_STTDM(long n)
{
	try
	{
		CoInitialize(NULL);
		xl.GetActiveObject(L"Excel.Application");
		xl->Visible = true;
		xl->PutScreenUpdating(VARIANT_FALSE);
		pWb = xl->GetActiveWorkbook();

		_WorksheetPtr pSheet = pWb->ActiveSheet;
		_bstr_t sh= pSheet->CodeName;
		RangePtr pRange=pSheet->Cells;

		if (sh==(_bstr_t)L"shHSNTVL" || sh==(_bstr_t)L"shHSNTCV")
		{
			int	_iCot[5];
			if (sh==(_bstr_t)L"shHSNTVL")
			{
				_iCot[1] = getColumn(pSheet,"DMVL_STT");
				_iCot[2] = getColumn(pSheet,"DMVL_MAVL");
				_iCot[4] = getColumn(pSheet,"DMVL_MAHS");
			}
			else
			{
				_iCot[1] = getColumn(pSheet,"DMBB_STT");
				_iCot[2] = getColumn(pSheet,"DMBB_MACV");
				_iCot[4] = getColumn(pSheet,"DMBB_MAHS");
			}

			int iEnd = FindComment(pSheet,_iCot[1],"END");

			// Lưu lại trạng thái đánh STT
			shTS = name_sheet((_bstr_t)"shTS");
			psTS = xl->Sheets->GetItem(shTS);
			putVal(psTS,"TS_ATOSTT",n);			

			// n=-1 chỉ đánh lại stt
			// n=0 sắp xếp toàn bộ, n=1,2 sắp xếp ngắt theo hạng mục, giai đoạn
			f_Auto_stt_dmuc(pRange,n,1,8,iEnd,_iCot[1],_iCot[2],_iCot[4]);
		}

		CoUninitialize();
	}
	catch (...){_s(L"Lỗi đánh lại số thứ tự");}	
}

// RCL Vật liệu -> getQCLM_VB -> open_quycach_laymau
void __stdcall open_quycach_laymau()
{
	try
	{
		xl_Create();
		
		_WorksheetPtr pSheet= pWb->ActiveSheet;		
		_bstr_t sSheet = pSheet->CodeName;

		if (sSheet != (_bstr_t)L"shHSNTVL") return;
		
		int iRow = pSheet->Application->ActiveCell->Row;
		int iCol = pSheet->Application->ActiveCell->Column;

		RangePtr pRange = pSheet->Cells;
		RangePtr PRS = pRange->GetItem(iRow,iCol);
		if(PRS->GetComment()==NULL)
		{
			// Sau khi kiểm tra vị trí có tồn tại comment không?
			// Kiểm tra tiếp giá trị mặc định là 0 hay 1. là 1 sẽ hiển thị quy cách mới, 0 - hiển thị quy cách cũ
			shTS = name_sheet("shTS");
			psTS = xl->Sheets->GetItem(shTS);
			prTS = psTS->Cells;

			int num = _wtoi(getVCell(psTS,L"CF_QCLMCV"));
			if(num!=1) num=0;
			if(num==0)
			{
				AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
				frm_QuycachOld open_dlg;
				open_dlg.DoModal();
			}
			else OQCLMAU();
		}
		else OQCLMAU();

		CoUninitialize();
	}
	catch(...){}
}

// Leo add 03.04.2018
// RCL Công việc -> Enter_detail -> Enter_RightClick
void __stdcall Enter_RightClick(long iRow, long iCol)
{
	try
	{
		xl_Create();

		_WorksheetPtr pSheet = pWb->ActiveSheet;
		_bstr_t sSheet = pSheet->CodeName;

		if (sSheet!=(_bstr_t)L"shHSNTCV") return;

		shTS = name_sheet((_bstr_t)"shTS");
		psTS = pWb->Worksheets->GetItem(shTS);
		prTS = psTS->Cells;

		int num = _wtoi(getVCell(psTS,L"CF_QCLMCV"));
		if(num!=1) num=0;

		// Lấy tạm biến sTukhoa vs pKey[i]
		sTukhoa = GIOR(6,68,prTS,L"");
		sTukhoa.Trim();
		if (sTukhoa != L"") for (int i = 0; i < 5; i++)
		{
			pKey[i + 1] = GIOR(6 + i, 68, prTS, L"");
		}
		else
		{
			pKey[1] = L"Nhà sản xuất/cung cấp:";
			pKey[2] = L"Mác:";
			pKey[3] = L"Độ sụt:";
			pKey[4] = L"Số lượng mẫu:";
			pKey[5] = L"Kích thước mẫu:";
		}

		for (int i = 1; i < 6; i++) pKLMBT[i] = pKey[i];

		RangePtr pRange = pSheet->Cells;
		CString szval = GIOR(iRow,iCol,pRange,L"");
		szval.Trim();

		RangePtr PRS = pRange->GetItem(iRow,iCol);
		if(PRS->GetComment()!=NULL || (PRS->GetComment()==NULL && szval==L"" && num==1))
		{
			iKqua=0;
			OQCLMAU();

			xl->PutScreenUpdating(VARIANT_FALSE);

			pSheet = pWb->ActiveSheet;
			sSheet = pSheet->CodeName;
			if (sSheet != (_bstr_t)L"shHSNTCV")
			{
				pSheet = pWb->Worksheets->GetItem(name_sheet((_bstr_t)"shHSNTCV"));
				pRange = pSheet->Cells;
				pSheet->Activate();
			}

			if(iKqua==1)
			{
				pRange->PutItem(iRow,iCol,(_bstr_t)schuoiMTN);
				PRS = pRange->GetItem(iRow, iCol);
				PRS->Interior->PutColor(xlNone);
				PRS->Font->PutColor(RGB(0,0,0));
				PRS->Font->PutItalic(1);
				PRS->Font->PutBold(0);
				PRS->EntireRow->Rows->AutoFit();
				if((int)PRS->GetRowHeight()>80) PRS->PutWrapText(FALSE);

				if(PRS->GetComment()!=NULL) PRS->ClearComments();	// Xóa comment QLCM nếu có
				int _colNK = getColumn(pSheet,L"DMBB_MH");
				pRange->PutItem(iRow,_colNK,L"LM");
			}

			xl->PutScreenUpdating(VARIANT_TRUE);

			return;
		}

		CString skt = szval;
		skt.Replace(L";",L"");
		skt.Trim();
		if(skt!=L"") schuoiMTN = szval;

		/////////////////////////////////////////////////			

		int iEnd = FindComment(pSheet,1,"END");
		if(iRow >= iEnd-1)
		{
			PRS = pSheet->GetRange(pRange->GetItem(iEnd,1),pRange->GetItem(iEnd+2,1));
			PRS->EntireRow->Insert(Excel::xlShiftDown);
		}

		if(f_Mod_check()!=1) return;	

		// Lấy tạm biến iKqua --> gán kiểm tra nút ok
		iKqua=0;

		AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
		Dlg_QCLMBetong open_dlg;
		open_dlg.DoModal();

		xl->PutScreenUpdating(VARIANT_FALSE);

		pSheet = pWb->ActiveSheet;
		sSheet = pSheet->CodeName;
		if (sSheet != (_bstr_t)L"shHSNTCV")
		{
			pSheet = pWb->Worksheets->GetItem(name_sheet((_bstr_t)"shHSNTCV"));
			pRange = pSheet->Cells;
			pSheet->Activate();
		}

		if(iKqua==1)
		{
			pRange->PutItem(iRow,iCol,(_bstr_t)schuoiMTN);
			PRS = pRange->GetItem(iRow, iCol);
			PRS->Interior->PutColor(xlNone);
			PRS->Font->PutColor(RGB(0,0,0));
			PRS->Font->PutItalic(1);
			PRS->Font->PutBold(0);
			PRS->EntireRow->Rows->AutoFit();
			if((int)PRS->GetRowHeight()>80) PRS->PutWrapText(FALSE);

			if(PRS->GetComment()!=NULL) PRS->ClearComments();	// Xóa comment QLCM nếu có
			int _colNK = getColumn(pSheet,L"DMBB_MH");
			pRange->PutItem(iRow,_colNK,L"LM");
		}

		xl->PutScreenUpdating(VARIANT_TRUE);

		CoUninitialize();
	}
	catch (...){}	
}


void __stdcall call_noilai_tieuchuan()
{
	try
	{
		xl_Create();		

		_bstr_t shDMVL = name_sheet((_bstr_t)"shHSNTVL");
		_WorksheetPtr psDMVL = pWb->Worksheets->GetItem(shDMVL);
		RangePtr prDMVL = psDMVL->Cells;

		_bstr_t shDMCV = name_sheet((_bstr_t)"shHSNTCV");
		_WorksheetPtr psDMCV = pWb->Worksheets->GetItem(shDMCV);
		RangePtr prDMCV = psDMCV->Cells;

		_bstr_t shNTC = name_sheet((_bstr_t)"shNhomTC");
		_WorksheetPtr psNTC = pWb->Worksheets->GetItem(shNTC);
		RangePtr prNTC = psNTC->Cells;

		int cotTCVL = getColumn(psDMVL,L"DMVL_TC");
		int iEndVL = FindComment(psDMVL,1,"END");

		int cotTCCV = getColumn(psDMCV,L"DMBB_TC");
		int iEndCV = FindComment(psDMCV,1,"END");

		int iTennhom = getColumn(psNTC,L"NHTC_TEN");

		xl->PutStatusBar(L"Đang tiến hành nối nhóm tiêu chuẩn. Vui lòng đợi trong giây lát...");

		int iktr=0;
		CString szval=L"";
		for (int i = 8; i < iEndVL; i++)
		{
			szval = GIOR(i,cotTCVL,prDMVL,L"");
			szval.Trim();
			if(szval!=L"")
			{
				iktr = FindValue(psNTC,iTennhom,1,0,(_bstr_t)szval);
				if(iktr>0)
				{
					szval.Format(L"='%s'!%s",(LPCTSTR)shNTC,GAR(iktr,iTennhom,prNTC,0));
					prDMVL->PutItem(i,cotTCVL,(_bstr_t)szval);
				}
			}
		}

		for (int i = 8; i < iEndCV; i++)
		{
			szval = GIOR(i,cotTCCV,prDMCV,L"");
			szval.Trim();
			if(szval!=L"")
			{
				iktr = FindValue(psNTC,iTennhom,1,0,(_bstr_t)szval);
				if(iktr>0)
				{
					szval.Format(L"='%s'!%s",(LPCTSTR)shNTC,GAR(iktr,iTennhom,prNTC,0));
					prDMCV->PutItem(i,cotTCCV,(_bstr_t)szval);
				}
			}			
		}

		xl->PutStatusBar(L"Ready");
		CoUninitialize();
	}
	catch (...){}
}


void __stdcall call_open_them_nhanh_tc()
{
	try
	{
		xl_Create();

		AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
		Dlg_Themtieuchuan open_dlg;
		open_dlg.DoModal();

		CoUninitialize();
	}
	catch(...){}
}


void __stdcall call_danh_sach_name()
{
	try
	{
		xl_Create();

		AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
		Dlg_NameSheet open_dlg;
		open_dlg.DoModal();

		CoUninitialize();
	}
	catch(...){}
}


void __stdcall call_selectCSDLNKy()
{
	try
	{
		xl_Create();

		sTukhoa=L"";
		AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
		Dlg_listNKyMDB open_dlg;
		open_dlg.DoModal();

		if(sTukhoa!=L"")
		{
			// Định nghĩa các sheet
			_bstr_t shNDNK = name_sheet("shMauNKY4");
			_WorksheetPtr psNDNK=xl->Sheets->GetItem(shNDNK);
			RangePtr prNDNK = psNDNK->Cells;

			int ir = getRow(psNDNK,"HK_NGAY");
			int ic = getColumn(psNDNK,"HK_NGAY");
			prNDNK->PutItem(ir,ic+2,(_bstr_t)sTukhoa);
		}
		CoUninitialize();
	}
	catch(...){}
}


void __stdcall call_tinhlaiDGKL()
{
	try
	{
		CoInitialize(NULL);
		xl.GetActiveObject(L"Excel.Application");
		xl->Visible = true;
		xl->PutScreenUpdating(VARIANT_FALSE);
		pWb = xl->GetActiveWorkbook();

		_WorksheetPtr pSheet = pWb->ActiveSheet;
		_bstr_t sh = pSheet->CodeName;
		RangePtr pRange = pSheet->Cells;

		int iRun=0;
		int cotSTT,cotMH,cotTen,cotKLG,iBegin,iEnd,iSum=0,iCuoi=0;
		if(sh==(_bstr_t)"shDTXD")
		{
			cotSTT = getColumn(pSheet,L"BIEUGIA_ID");
			cotMH = getColumn(pSheet,L"BIEUGIA_MHDG");
			cotTen = getColumn(pSheet,L"BIEUGIA_TEN");
			cotKLG = getColumn(pSheet,L"BIEUGIA_KL");
			
			iBegin=7;
			iEnd = FindComment(pSheet,1,L"END");
			iRun=1;
		}
		else if(sh==(_bstr_t)"shHSNTCV")
		{
			cotSTT = getColumn(pSheet,L"DMBB_STT");
			cotMH = getColumn(pSheet,L"DMBB_MACV");
			cotTen = getColumn(pSheet,L"DMBB_ND");
			cotKLG = getColumn(pSheet,L"DMBB_KL");

			iBegin=8;
			iEnd = FindComment(pSheet,cotSTT,L"END");
			iRun=1;
		}
		else if(sh==(_bstr_t)"shKhoiluong") iRun=2;

		if(iRun==1)
		{
			RangePtr PRS;
			CString szval=L"",sformula=L"",smahieu=L"";
			for (int i = iBegin; i <= iEnd; i++)
			{
				smahieu = GIOR(i,cotMH,pRange,L"");	// Mã hiệu
				if(smahieu==L"" && i<iEnd)
				{
					szval = GIOR(i,cotTen,pRange,L"");	// Tên công việc
					if(szval!=L"")
					{
						PRS = pRange->GetItem(i,cotKLG);	// Công thức cột khối lượng
						sformula = PRS->GetFormula();

						szval.Format(L"=kl(%s)",GAR(i,cotTen,pRange,0));
						pRange->PutItem(i,cotKLG,(_bstr_t)szval);
					
						szval = GIOR(i,cotKLG,pRange,L"");
						if(szval==L"") pRange->PutItem(i,cotKLG,(_bstr_t)L"");
					}
				}
				else
				{
					szval = GIOR(i,cotSTT,pRange,L"");
					if(_wtoi(szval)>0 || smahieu==L"HM" || smahieu==L"GD" || i==iEnd)
					{
						iCuoi=i;
						if(iCuoi==iEnd)
						{
							for (int j = iEnd-1; j >= iBegin; j--)
							{
								szval = GIOR(j,cotTen,pRange,L"");
								if(szval!=L"")
								{
									szval = GIOR(j,cotSTT,pRange,L"");
									if(szval==L"")
									{
										iCuoi=j+1;
										break;
									}
								}
							}
						}

						if(iSum>0)
						{
							if(iSum+2<=iCuoi)
							{
								szval.Format(L"=SUM(%s:%s)",GAR(iSum+1,cotKLG,pRange,0),GAR(iCuoi-1,cotKLG,pRange,0));
								pRange->PutItem(iSum,cotKLG,(_bstr_t)szval);
							}						
						}

						iSum=iCuoi;
					}
				}
			}
		}
		else if(iRun==2)
		{
			cotSTT = getColumn(pSheet,L"DTXD_STT");
			cotMH = getColumn(pSheet,L"DTXD_MH");
			cotTen = getColumn(pSheet,L"DTXD_TENCV");
			cotKLG = getColumn(pSheet,L"DTXD_KLTB");

			int iNHD=0,iStart=1;			
			RangePtr PRS;
			CString szval=L"",sadd=L"",sformula=L"",smahieu=L"";
			_bstr_t bscmt;

			while (1)
			{
				iBegin = FindValue(pSheet,cotSTT,iStart,0,(_bstr_t)"STT");
				if(iBegin==0) break;
				
				iNHD=0;
				iEnd = 0;
				iStart = iBegin+3;
				PRS = pRange->GetItem(iBegin-4,cotSTT);
				if(PRS->GetComment()!=NULL)
				{
					bscmt = PRS->GetComment()->Text();
					sadd = (LPCTSTR)(bscmt);
					sadd.Trim();
					sadd.Replace(L"DAU_GD",L"");
					szval.Format(L"CUOI_GD%s",sadd);

					iNHD = FindCEnd(pSheet,cotTen,(_bstr_t)"NHD",iStart);
					iEnd = FindComment(pSheet,cotSTT,(_bstr_t)szval);
				}

				if(iNHD==0 || iEnd==0) break;
				
				for (int i = iStart; i <= iNHD; i++)
				{
					smahieu = GIOR(i,cotMH,pRange,L"");	// Mã hiệu
					if(smahieu==L"" && i<iNHD)
					{
						szval = GIOR(i,cotTen,pRange,L"");	// Tên công việc
						if(szval!=L"")
						{
							PRS = pRange->GetItem(i,cotKLG);	// Công thức cột khối lượng
							sformula = PRS->GetFormula();

							szval.Format(L"=kl(%s)",GAR(i,cotTen,pRange,0));
							pRange->PutItem(i,cotKLG,(_bstr_t)szval);
					
							szval = GIOR(i,cotKLG,pRange,L"");
							if(szval==L"") pRange->PutItem(i,cotKLG,(_bstr_t)L"");
						}
					}
					else
					{
						szval = GIOR(i,cotSTT,pRange,L"");
						if(_wtoi(szval)>0 || smahieu==L"HM" || smahieu==L"GD" || i==iNHD)
						{
							iCuoi=i;
							if(iCuoi==iNHD)
							{
								for (int j = iNHD-1; j >= iBegin; j--)
								{
									szval = GIOR(j,cotTen,pRange,L"");
									if(szval!=L"")
									{
										szval = GIOR(j,cotSTT,pRange,L"");
										if(szval==L"")
										{
											iCuoi=j+1;
											break;
										}
									}
								}
							}

							if(iSum>0)
							{
								if(iSum+2<=iCuoi)
								{
									szval.Format(L"=SUM(%s:%s)",GAR(iSum+1,cotKLG,pRange,0),GAR(iCuoi-1,cotKLG,pRange,0));
									pRange->PutItem(iSum,cotKLG,(_bstr_t)szval);
								}						
							}

							iSum=iCuoi;
						}
					}
				}

				for (int i = iNHD; i <= iEnd; i++)
				{
					smahieu = GIOR(i,cotMH,pRange,L"");	// Mã hiệu
					if(smahieu==L"" && i<iEnd)
					{
						szval = GIOR(i,cotTen,pRange,L"");	// Tên công việc
						if(szval!=L"")
						{
							PRS = pRange->GetItem(i,cotKLG);	// Công thức cột khối lượng
							sformula = PRS->GetFormula();

							szval.Format(L"=kl(%s)",GAR(i,cotTen,pRange,0));
							pRange->PutItem(i,cotKLG,(_bstr_t)szval);
					
							szval = GIOR(i,cotKLG,pRange,L"");
							if(szval==L"") pRange->PutItem(i,cotKLG,(_bstr_t)L"");
						}
					}
					else
					{
						szval = GIOR(i,cotSTT,pRange,L"");
						if(_wtoi(szval)>0 || smahieu==L"HM" || smahieu==L"GD" || i==iEnd)
						{
							iCuoi=i;
							if(iCuoi==iEnd)
							{
								for (int j = iEnd-1; j >= iBegin; j--)
								{
									szval = GIOR(j,cotTen,pRange,L"");
									if(szval!=L"")
									{
										szval = GIOR(j,cotSTT,pRange,L"");
										if(szval==L"")
										{
											iCuoi=j+1;
											break;
										}
									}
								}
							}

							if(iSum>0)
							{
								if(iSum+2<=iCuoi)
								{
									szval.Format(L"=SUM(%s:%s)",GAR(iSum+1,cotKLG,pRange,0),GAR(iCuoi-1,cotKLG,pRange,0));
									pRange->PutItem(iSum,cotKLG,(_bstr_t)szval);
								}						
							}

							iSum=iCuoi;
						}
					}
				}
			}
		}

		CoUninitialize();
	}
	catch(...){_s(L"Lỗi chạy lại diễn giải khối lượng. ");}
}


void __stdcall call_doc_tieuchuan()
{
	try
	{
		CoInitialize(NULL);
		xl.GetActiveObject(L"Excel.Application");
		xl->Visible = true;
		xl->PutScreenUpdating(VARIANT_FALSE);
		pWb = xl->GetActiveWorkbook();

		_WorksheetPtr pSheet = pWb->ActiveSheet;
		_bstr_t sh= pSheet->CodeName;
		RangePtr pRange=pSheet->Cells;
		int _irow = pSheet->Application->ActiveCell->Row;
		int _icolumn = pSheet->Application->ActiveCell->Column;

		CString szval = GIOR(_irow,_icolumn,pRange,L"");
		szval.Replace(L"'",L"");
		szval.Trim();
		if(szval!=L"")
		{
			if(getPathQLCL()==L"") return;
			CString sFullPath = pathMDB;
			if(sFullPath==L"") return;		
			if(connectDsn(sFullPath)==-1) return;

			CConnection ObjConn;
			ObjConn.OpenAccessDB(sFullPath, L"", L"");

			CString sMaTC=L"";
			SqlString.Format(L"SELECT * FROM tbTentieuchuan WHERE (MaTC LIKE '%s' OR TenTC  LIKE '%s');",szval,szval);
			CRecordset* Recset = ObjConn.Execute(SqlString);
			while( !Recset->IsEOF() )
			{
				Recset->GetFieldValue(L"MaTC",sMaTC);
				break;
			}
			Recset->Close();
			ObjConn.CloseDatabase();

			if(sMaTC!=L"") f_DocFileTC(sMaTC, 1);
		}

		CoUninitialize();
	}
	catch(...){}
}


void __stdcall call_RClick_KyBB()
{
	try
	{
		xl_Create();

		_WorksheetPtr pSheet = pWb->ActiveSheet;
		_bstr_t sh= pSheet->CodeName;
		RangePtr pRange=pSheet->Cells;
		
		CString szcotKBB=L"";
		if(sh==(_bstr_t)L"shHSNTVL") szcotKBB = L"DMVL_KBB";
		else if(sh==(_bstr_t)L"shHSNTCV") szcotKBB = L"DMBB_KBB";
		else if(sh==(_bstr_t)L"shHSNTGD") szcotKBB = L"DMGD_KBB";
		if(szcotKBB!=L"")
		{
			int _irow = pSheet->Application->ActiveCell->Row;
			CString szval= GIOR(_irow,1,pRange,L"");
			if(_wtoi(szval)>0)
			{
				sTukhoa=L"";
				AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
				Dlg_NhomkyBB open_dlg;
				open_dlg.DoModal();

				if(sTukhoa!=L"")
				{
					if(sTukhoa==L"0") sTukhoa=L"";
					int cotKBB = getColumn(pSheet,szcotKBB);
					pRange->PutItem(_irow,cotKBB,(_bstr_t)sTukhoa);
				}
			}				
		}
				
		CoUninitialize();
	}
	catch(...){}
}


void __stdcall call_RClick_BangmauNT()
{
	try
	{
		xl_Create();

		_WorksheetPtr pSheet = pWb->ActiveSheet;
		RangePtr pRange = pSheet->Cells;
		CString sh = (LPCTSTR)pSheet->CodeName;
		if(sh==L"shHSNTVL" || sh==L"shHSNTCV" 
			|| sh==L"shHSNTGD" || sh==L"shBangmauNT")
		{
			sTukhoa=L"";
			CString szval=L"";
			int iRowActive = 1, iColumnActive = 1;
			if (sh != L"shBangmauNT")
			{
				iRowActive = pSheet->Application->ActiveCell->GetRow();
				iColumnActive = pSheet->Application->ActiveCell->GetColumn();
				szval = GIOR(iRowActive, iColumnActive, pRange, L"");
			}			

			AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
			Dlg_NhomBangmau p;
			p.szTenmau = szval;
			p.iRowActive = iRowActive;
			p.iColumnActive = iColumnActive;
			p.pShDMuc = pSheet;

			p.DoModal();

			if (sh != L"shBangmauNT")
			{
				if (sTukhoa != L"")
				{
					if (sTukhoa == L"0") sTukhoa = L"";					
					RangePtr PRS = pRange->GetItem(iRowActive, iColumnActive);
					PRS->PutValue2((_bstr_t)sTukhoa);
					PRS->PutHorizontalAlignment(xlLeft);
				}
			}			
		}
				
		CoUninitialize();
	}
	catch(...){}
}


void __stdcall call_Nhanbanmau()
{
	try
	{
		xl_Create();

		_WorksheetPtr pSheet = pWb->ActiveSheet;
		RangePtr pRange = pSheet->Cells;
		CString sh = (LPCTSTR)pSheet->CodeName;
		if (sh == L"shBangmauNT")
		{
			CString szval = L"";
			int iRowActive = pSheet->Application->ActiveCell->GetRow();
			int nvtri = 0, nEnd = 0;

			for (int i = iRowActive; i >= 1; i--)
			{
				szval = GIOR(i, 1, pRange, L"");
				szval.Trim();
				if (szval != L"" && szval != L"END")
				{
					nvtri = i;
					break;
				}
			}

			int iRowLast = pSheet->Cells->SpecialCells(xlCellTypeLastCell)->GetRow() + 1;
			for (int i = iRowActive + 1; i < iRowLast; i++)
			{
				szval = GIOR(i, 1, pRange, L"");
				szval.Trim();
				if (szval == L"END")
				{
					nEnd = i;
					break;
				}
			}

			if (nvtri > 0 && nEnd > 0)
			{
				xl->PutScreenUpdating(VARIANT_FALSE);

				int slchen = nEnd - nvtri + 2;

				RangePtr PRS = pSheet->GetRange(
					pRange->GetItem(nEnd + 1, 1), pRange->GetItem(nEnd + slchen, 1));
				PRS->EntireRow->Insert(xlShiftDown);

				PRS = pSheet->GetRange(
					pRange->GetItem(nvtri, 1), pRange->GetItem(nEnd, 1));
				PRS->EntireRow->Copy(pRange->GetItem(nEnd + 2, 1));

				xl->PutCutCopyMode(XlCutCopyMode(false));

				xl->PutScreenUpdating(VARIANT_TRUE);
				_s(L"Nhân bản mẫu thành công.");
			}			
		}

		CoUninitialize();
	}
	catch (...) {}
}


void __stdcall call_ChonCSDLVLCV()
{
	try
	{
		xl_Create();

		shTS = name_sheet((_bstr_t)"shTS");
		psTS = xl->Sheets->GetItem(shTS);
		prTS = psTS->Cells;

		AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
		Dlg_ChonCSDLVLCV open_dlg;
		open_dlg.DoModal();
				
		CoUninitialize();
	}
	catch(...){}
}


void __stdcall call_TinhdiengiaiME()
{
	try
	{
		xl_Create();

		_WorksheetPtr pSheet = pWb->ActiveSheet;
		_bstr_t sh= pSheet->CodeName;		

		if(sh==(_bstr_t)L"shKhoiluong")
		{
			shTS = name_sheet((_bstr_t)"shTS");
			psTS = xl->Sheets->GetItem(shTS);
			prTS = psTS->Cells;
			int icheckME = _wtoi(getVCell(psTS,L"CF_TDGME"));	// icheckME=-1 CHECK | =0 UNCHECK
			int iChange=0;

			CString szval=L"";
			RangePtr PRS, pRange=pSheet->Cells;
			int rowDai = getRow(pSheet,L"DTXD_DAI");
			int cotDai = getColumn(pSheet,L"DTXD_DAI");
			int cotCao = getColumn(pSheet,L"DTXD_CAO");
			int cotHS = getColumn(pSheet,L"DTXD_HS");

			// Hiển thị lại kích thước dài rộng cao
			if(icheckME==0)
			{
				iChange=1;
				PRS = pSheet->GetRange(pRange->GetItem(1,cotDai),pRange->GetItem(1,cotCao));
				PRS->EntireColumn->PutHidden(false);

				if(cotCao+1<cotHS)
				{
					PRS = pSheet->GetRange(pRange->GetItem(1,cotCao+1),pRange->GetItem(1,cotHS-1));
					PRS->EntireColumn->PutHidden(true);
				}
			}
			else
			{
				iKqua=1;	// lấy tạm biến
				AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
				Dlg_Chen_dong _dlg;
				_dlg.DoModal();

				if(iKqua>0)
				{
					iChange=1;
					PRS = pSheet->GetRange(pRange->GetItem(1,cotDai),pRange->GetItem(1,cotCao));
					PRS->EntireColumn->PutHidden(true);

					if(cotCao+1<cotHS)
					{
						PRS = pSheet->GetRange(pRange->GetItem(1,cotCao+1),pRange->GetItem(1,cotHS-1));
						PRS->EntireColumn->PutHidden(false);
					}

					// Xác định số cột cần chèn thêm
					iSLTang = iKqua;					// tổng số tầng sử dụnng
					int hientai = cotHS-cotCao-1;		// số tầng đang có
					int sothem = hientai-iSLTang;		// số tầng cần chèn thêm (sothem có thể >0 hoặc <0)

					if(sothem>0)
					{
						// Xóa bớt
						PRS = pSheet->GetRange(pRange->GetItem(1,cotHS-sothem),pRange->GetItem(1,cotHS-1));
						PRS->EntireColumn->Delete(xlToLeft);
					}
					else if(sothem<0)
					{
						// Chèn thêm (chú ý sothem đang âm)
						PRS = pSheet->GetRange(pRange->GetItem(1,cotHS),pRange->GetItem(1,cotHS-sothem-1));
						PRS->EntireColumn->Insert(xlToRight);

						cotHS = getColumn(pSheet,L"DTXD_HS");
						PRS = pSheet->GetRange(pRange->GetItem(rowDai,cotCao+1),pRange->GetItem(rowDai,cotHS-1));
						PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->PutWeight(xlThin);
					}

					pRange->PutItem(rowDai-1,cotCao+1,(_bstr_t)L"KHỐI LƯỢNG");					
					for (int i = 1; i <= iSLTang; i++)
					{
						szval.Format(L"Tầng %d",i);
						pRange->PutItem(rowDai,cotCao+i,(_bstr_t)szval);
					}
				}
				else
				{
					putVal(psTS,(_bstr_t)L"CF_TDGME",L"0");
				}
			}

			if(iChange==1)
			{
				// Đánh lại STT [i]
				int dem=0,numbd=12;
				int rowbd=getColumn(pSheet,L"DTXD_HS");
				if(icheckME==-1)
				{
					numbd=9;
					rowbd=cotCao+1;
					cotHS = getColumn(pSheet,L"DTXD_HS");
				}

				int cotCLech = getColumn(pSheet,L"DTXD_CLKL");
				for (int i = rowbd; i <= cotCLech; i++)
				{
					szval.Format(L"[%d]",numbd+dem);
					pRange->PutItem(rowDai+1,rowbd+dem,(_bstr_t)szval);
					dem++;
				}

				// Đổ lại công thức
				int pRowEnd = pRange->SpecialCells(xlCellTypeLastCell)->GetRow();
				int cotKLMBP = getColumn(pSheet,L"DTXD_KLBP");
				for (int i = 1; i <= pRowEnd; i++)
				{
					PRS = pRange->GetItem(i,cotKLMBP);
					szval = PRS->GetFormula();
					if(szval.Left(3)==L"=kl" || szval.Left(3)==L"=PR")
					{
						if(icheckME==-1)
						{
							szval=L"";
							if(cotCao+1<cotHS)
							{
								szval = L"=";
								for (int j = cotCao+1; j < cotHS; j++)
								{
									szval+=L"kl(";
									szval+=GAR(i,j,pRange,0);
									szval+=L")";					
									if(j<cotHS-1) szval+=L"+";
								}
							}
						}
						else
						{
							szval.Format(L"=PRODUCT(%s:%s,%s)",
								GAR(i,cotDai,pRange,0),GAR(i,cotCao,pRange,0),GAR(i,cotHS,pRange,0));
						}
						pRange->PutItem(i,cotKLMBP,(_bstr_t)szval);
					}
				}
			}			
		}	

		CoUninitialize();
	}
	catch(...){}
}

void __stdcall call_Apdungbangmau()
{
	try
	{
		xl_Create();

		_WorksheetPtr pSheet = pWb->ActiveSheet;
		RangePtr pRange = pSheet->Cells;
		CString sh = (LPCTSTR)pSheet->CodeName;
		if (sh == L"shBangmauNT")
		{
			if (iRowBMau > 0 && iColBMau > 0)
			{
				// Xác định vị trí
				int iRowActive = pSheet->Application->ActiveCell->GetRow();
				
				int irow = 0;
				CString szval = L"";
				for (int i = iRowActive; i >= 1; i--)
				{
					szval = GIOR(i, 1, pRange, L"");
					if (szval != L"" && szval != L"END")
					{
						irow = i;
						break;
					}
				}

				if (irow > 0)
				{
					szval = GIOR(irow, 1, pRange, L"");
					RangePtr Prg = pShDMBMau->Cells;
					RangePtr PRS = Prg->GetItem(iRowBMau, iColBMau);
					PRS->PutValue2((_bstr_t)szval);
					PRS->PutHorizontalAlignment(xlLeft);

					if ((int)pShDMBMau->GetVisible() != -1) pShDMBMau->PutVisible(XlSheetVisibility::xlSheetVisible);
					pShDMBMau->Activate();
				}
				
				iRowBMau = 0;
				iColBMau = 0;
			}			
		}

		CoUninitialize();
	}
	catch (...) {}
}

void __stdcall call_ChangePathEx()
{

}

void __stdcall call_MOpenFile()
{
	try
	{
		CoInitialize(NULL);
		xl.GetActiveObject(L"Excel.Application");
		xl->Visible = true;
		xl->PutScreenUpdating(VARIANT_FALSE);
		pWb = xl->GetActiveWorkbook();

		_WorksheetPtr pShActive = pWb->ActiveSheet;
		_bstr_t sh = pShActive->CodeName;
		if (sh == (_bstr_t)L"shNhomTC")
		{
			int iColumnFile = getColumn(pShActive, L"TCH_OPEN");
			if (iColumnFile > 0)
			{
				int iRowActive = pShActive->Application->ActiveCell->GetRow();

				RangePtr pRange = pShActive->Cells;
				CString szval = GIOR(iRowActive, iColumnFile, pRange, L"");
				if (szval != L"")
				{
					vector<CString> vecHyper;
					RangePtr PRS = pRange->GetItem(iRowActive, iColumnFile);
					int ncoutHyp = _xlGetAllHyperLink(PRS, vecHyper);
					if (ncoutHyp > 0)
					{
						if (_FileExistsUnicode(vecHyper[0]) == 1)
						{
							ShellExecute(NULL, L"open", vecHyper[0], NULL, NULL, SW_SHOWMAXIMIZED);
						}
					}
					vecHyper.clear();
				}
			}
		}

		CoUninitialize();
	}
	catch (...) {}
}

void __stdcall call_MOpenFolder()
{
	try
	{
		CoInitialize(NULL);
		xl.GetActiveObject(L"Excel.Application");
		xl->Visible = true;
		xl->PutScreenUpdating(VARIANT_FALSE);
		pWb = xl->GetActiveWorkbook();

		_WorksheetPtr pShActive = pWb->ActiveSheet;
		_bstr_t sh = pShActive->CodeName;
		if (sh == (_bstr_t)L"shNhomTC")
		{
			int iColumnFile = getColumn(pShActive, L"TCH_OPEN");
			if (iColumnFile > 0)
			{
				int iRowActive = pShActive->Application->ActiveCell->GetRow();

				RangePtr pRange = pShActive->Cells;
				CString szval = GIOR(iRowActive, iColumnFile, pRange, L"");
				if (szval != L"")
				{
					vector<CString> vecHyper;
					RangePtr PRS = pRange->GetItem(iRowActive, iColumnFile);
					int ncoutHyp = _xlGetAllHyperLink(PRS, vecHyper);
					if (ncoutHyp > 0)
					{
						if (_FileExistsUnicode(vecHyper[0]) == 1)
						{
							SHELLEXECUTEINFO shExecInfo = { 0 };
							shExecInfo.cbSize = sizeof(shExecInfo);
							shExecInfo.lpFile = _T("Explorer.exe");
							shExecInfo.lpParameters = L"/Select, " + vecHyper[0];
							shExecInfo.nShow = SW_SHOWNORMAL;
							shExecInfo.lpVerb = _T("Open");
							shExecInfo.fMask = SEE_MASK_INVOKEIDLIST | SEE_MASK_FLAG_DDEWAIT | SEE_MASK_FLAG_NO_UI;
							VERIFY(ShellExecuteEx(&shExecInfo));
						}
					}
					vecHyper.clear();
				}				
			}
		}

		CoUninitialize();
	}
	catch (...) {}
}

void __stdcall call_null_sub()
{
	
}
