#include "cls_main.h"
#include "Dlg_importpic.h"
#include "Dlg_Open_DMCV.h"
#include "Dlg_Open_DSVL.h"
#include "Dlg_move_dulieu.h"
#include "Dlg_trauu_tieuchuan.h"
#include "funct.h"

inline void SetFindEND(RangePtr pRng){
	pRng->Find(L"",vtMissing,XlFindLookIn::xlFormulas,XlLookAt::xlPart,
		XlSearchOrder::xlByColumns,XlSearchDirection::xlNext,false,false,false);
}

void cls_main::f_xacdinh_cot_DMVL()
{
	_colvl[1] = getColumn(psDMVL,"DMVL_STT");
	_colvl[2] = getColumn(psDMVL,"DMVL_MAVL");
	_colvl[3] = getColumn(psDMVL,"DMVL_MHDG");
	_colvl[4] = getColumn(psDMVL,"DMVL_MAHS");
	_colvl[5] = getColumn(psDMVL,"DMVL_ND");
	_colvl[6] = getColumn(psDMVL,"DMVL_TC");
	_colvl[7] = getColumn(psDMVL,"DMVL_NK_NG");
	_colvl[8] = getColumn(psDMVL,"DMVL_NK_SP");
	_colvl[9] = getColumn(psDMVL,"DMVL_LM_NG");
	_colvl[10] = getColumn(psDMVL,"DMVL_LM_GIO");
	_colvl[11] = getColumn(psDMVL,"DMVL_LM_KQ");
	_colvl[12] = getColumn(psDMVL,"DMVL_NTNB_NG");
	_colvl[13] = getColumn(psDMVL,"DMVL_NTNB_GIO");
	_colvl[14] = getColumn(psDMVL,"DMVL_YC");
	_colvl[15] = getColumn(psDMVL,"DMVL_AB_NG");
	_colvl[16] = getColumn(psDMVL,"DMVL_AB_GIO");
	_colvl[17] = getColumn(psDMVL,"DMVL_KBB");
	_colvl[18] = getColumn(psDMVL,"DMVL_TC2");
	//_colvl[19] = getColumn(psDMVL,"DMVL_XNK1");
	//_colvl[20] = getColumn(psDMVL,"DMVL_XNK2");

	_colvl[21] = getColumn(psDMVL,"DMVL_XXU");		// Cột xuất xứ (10.03.2018)
	_colvl[22] = getColumn(psDMVL,"DMVL_DOT");		// Cột đợt (11.09.2019)
}

void cls_main::f_xacdinh_cot_DMCV()
{
	_colcv[1] = getColumn(psDMCV,"DMBB_STT");
	_colcv[2] = getColumn(psDMCV,"DMBB_MACV");
	_colcv[3] = getColumn(psDMCV,"DMBB_MH");
	_colcv[4] = getColumn(psDMCV,"DMBB_MAHS");
	_colcv[5] = getColumn(psDMCV,"DMBB_ND");
	_colcv[6] = getColumn(psDMCV,"DMBB_VT");
	_colcv[7] = getColumn(psDMCV,"DMBB_KH");
	_colcv[8] = getColumn(psDMCV,"DMBB_TC");
	_colcv[9] = getColumn(psDMCV,"DMBB_TN_NGAY");
	_colcv[10] = getColumn(psDMCV,"DMBB_TN_GIO");
	_colcv[11] = getColumn(psDMCV,"DMBB_TN_KQ");
	_colcv[12] = getColumn(psDMCV,"DMBB_NB_NGAY");
	_colcv[13] = getColumn(psDMCV,"DMBB_NB_GIO");
	_colcv[14] = getColumn(psDMCV,"DMBB_PHIEUYC");
	_colcv[15] = getColumn(psDMCV,"DMBB_AB_NGAY");
	_colcv[16] = getColumn(psDMCV,"DMBB_AB_GIO");
	_colcv[17] = getColumn(psDMCV,"DMBB_KBB");
	_colcv[18] = getColumn(psDMCV,"DMBB_HK_NGAYBD");
	_colcv[19] = getColumn(psDMCV,"DMBB_HK_NGAYKT");
	//_colcv[20] = getColumn(psDMCV,"DMBB_NLUC");
	_colcv[21] = getColumn(psDMCV,"DMBB_PS");
	_colcv[22] = getColumn(psDMCV,"DMBB_IMG");
	_colcv[23] = getColumn(psDMCV,"DMBB_TC2");
	//_colcv[24] = getColumn(psDMCV,"DMBB_XNK1");
	//_colcv[25] = getColumn(psDMCV,"DMBB_XNK2");
	//_colcv[26] = getColumn(psDMCV,"DMBB_XNK3");
	_colcv[27] = getColumn(psDMCV,"DMBB_DOT");		// Đợt thanh toán	
	_colcv[28] = getColumn(psDMCV,"DMBB_DTLMAU");	// Đối tượng lấu mẫu
}

// Edit 08.08.2017
void cls_main::f_xacdinh_cot_DMGD()
{
	_colgd[1] = getColumn(psDMGD,"DMGD_STT");
	_colgd[2] = getColumn(psDMGD,"DMGD_MDM");
	_colgd[3] = getColumn(psDMGD,"DMGD_MHDG");
	_colgd[4] = getColumn(psDMGD,"DMGD_HS");
	_colgd[5] = getColumn(psDMGD,"DMGD_ND");
	_colgd[6] = getColumn(psDMGD,"DMGD_MATC");
	_colgd[7] = getColumn(psDMGD,"DMGD_TC");
	_colgd[8] = getColumn(psDMGD,"DMGD_NTNB_NG");
	_colgd[9] = getColumn(psDMGD,"DMGD_NTNB_GIO");
	_colgd[10] = getColumn(psDMGD,"DMGD_YC");
	_colgd[11] = getColumn(psDMGD,"DMGD_AB_NG");
	_colgd[12] = getColumn(psDMGD,"DMGD_AB_GIO");
	_colgd[13] = getColumn(psDMGD,"DMGD_KBB");
	_colgd[14] = getColumn(psDMGD,"DMGD_IMG");
	//_colgd[15] = getColumn(psDMGD,"DMGD_XNK1");
	_colgd[16] = getColumn(psDMGD,"DMGD_DOT");

}


void cls_main::f_xacdinh_cot_VLNV()
{
	_colnv[1] = getColumn(psVLNV,"VLNV_STT");
	_colnv[2] = getColumn(psVLNV,"VLNV_NNK");
	_colnv[3] = getColumn(psVLNV,"VLNV_MVL");
	_colnv[4] = getColumn(psVLNV,"VLNV_SBB");
	_colnv[5] = getColumn(psVLNV,"VLNV_LVL");
	_colnv[6] = getColumn(psVLNV,"VLNV_DG");
	_colnv[7] = getColumn(psVLNV,"VLNV_MAC");
	_colnv[8] = getColumn(psVLNV,"VLNV_DVI");
	_colnv[9] = getColumn(psVLNV,"VLNV_KLG");
	//_colnv[10] = getColumn(psVLNV,"VLNV_HSQD");
	//_colnv[11] = getColumn(psVLNV,"VLNV_DVQD");
	//_colnv[12] = getColumn(psVLNV,"VLNV_KLQD");
	_colnv[13] = getColumn(psVLNV,"VLNV_GC");
	_colnv[14] = getColumn(psVLNV,"VLNV_XXU");	// Cột xuất xứ (10.03.2018)
}


void cls_main::f_xacdinh_cot_DMHS()
{
	_colhs[1] = getColumn(psDMHS,"DMHS_STT");
	_colhs[2] = getColumn(psDMHS,"DMHS_MCV");
	_colhs[3] = getColumn(psDMHS,"DMHS_ND");
	_colhs[4] = getColumn(psDMHS,"DMHS_NNH");
	_colhs[5] = getColumn(psDMHS,"DMHS_HSNT");
	_colhs[6] = getColumn(psDMHS,"DMHS_NLM");
	_colhs[7] = getColumn(psDMHS,"DMHS_NNT");
	_colhs[8] = getColumn(psDMHS,"DMHS_SMH");
	_colhs[9] = getColumn(psDMHS,"DMHS_TTHS");
}


// hàm tách nội dung chi tiết công việc - vị trí chi tiết thứ n  (n=1,2,3,4,5)
CString cls_main::f_tachcv(CString noidung,int n)		
{
	wchar_t *wkey;
	wchar_t pKx[20][300];

	wkey  = new wchar_t[SIZE_LINE];
	wcscpy(wkey,noidung);

	// Gán giá trị cho các từ khóa được chia nhỏ
	for(int i=1;i<=6;i++) wcscpy(pKx[i],L"");

	// Tách chuỗi thành các từ khóa nhỏ pKx[i]
	wchar_t * tg = new wchar_t[SIZE_LINE];
	ikey=0;	// số lượng từ khóa
	tg = wcstok (wkey,L";");
	while (tg != NULL)
	{
		ikey++;
		wcscpy(pKx[ikey],tg);
		tg = wcstok(NULL,L";");
		if(ikey==5) break;
	}

	// Xóa khoảng trắng trước và sau chuỗi
	CString gtri;
	gtri.Format(L"%s", pKx[n]);
	gtri.Trim();

	return gtri;
}

// hàm tách khối lượng - n=1: khối lượng, n=2 đơn vị
CString cls_main::f_tachklg(CString kluong,int n)
{
	CString val=L"";
	int iLen = kluong.GetLength();

	if(iLen>0)
	{
		int k=0;
		CString ktra = kluong.Mid(k,1);
		while (ktra!=L" " && k<iLen)
		{
			k++;
			ktra = kluong.Mid(k,1);
		}

		if(n==2) return kluong.Right(iLen-k);
		else
			{
				val = kluong.Left(k);
				val.Replace(L",",L".");
				return val;
			}
	}

	return val;
}

// Vật liệu nhập về
void cls_main::f_truyen_dulieu_VLNV()
{
	try
	{
		CoInitialize(NULL);
		xl.GetActiveObject(L"Excel.Application");
		xl->Visible = true;
		xl->PutScreenUpdating(VARIANT_FALSE);		
		pWb = xl->GetActiveWorkbook();

		shDMVL = name_sheet("shHSNTVL");
		shVLNV = name_sheet("Sheet8");
	
		psDMVL = xl->Sheets->GetItem(shDMVL);
		psVLNV = xl->Sheets->GetItem(shVLNV);
		prDMVL = psDMVL->Cells;
		prVLNV = psVLNV->Cells;

		// Xác định cột
		f_xacdinh_cot_DMVL();
		f_xacdinh_cot_VLNV();

		// Xác định dòng bắt đầu		
		int inum = 0;
		CString f1, f4, f5, fta;
		_bstr_t f2, val1, val2, val3;		
		int xdeDMVL = FindComment(psDMVL, _colnv[1], "END");

		// Kiểm tra xem có tồn tại dữ liệu lấy mẫu hay không?
		int _arr = 0, _gan = 0, icheck = 0;
		for (int i = 8; i < xdeDMVL; i++)
		{
			// Kiểm tra có tồn tại vật liệu không
			f1 = prDMVL->GetItem(i, _colvl[3]);
			if (f1 == L"LM")
			{
				icheck = 1;
				break;
			}
		}

		if (icheck == 0)
		{
			_s(L"Không tồn tại dữ liệu lấy mẫu. Vui lòng kiểm tra lại.");
			return;
		}

		// Xóa dòng công việc tồn tại từ trước		
		
		int irowbd = getRow(psVLNV, L"VLNV_STT");
		int xdeVLNV = FindComment(psVLNV, _colnv[1], "END");
		if (xdeVLNV >= irowbd + 2)
		{
			psVLNV->GetRange(prVLNV->GetItem(irowbd + 2, 1), prVLNV->GetItem(xdeVLNV, 1))->EntireRow->Delete(xlShiftUp);
		}		

		// Load path
		if(getPathQCLM()==L"") return;
		if(connectDsn(pathQCLM)==-1) return;
		ObjConn.OpenAccessDB(pathQCLM, L"", L"");

		xl->StatusBar = (_bstr_t)L"Đang tiến hành phân tích dữ liệu. Vui lòng đợi trong giây lát...";

		int jd=0;
		BOOLEAN bf=TRUE;
		CString szval=L"";
		for (int i = 8; i < xdeDMVL; i++)
		{
			szval.Format(L"Đang phân tích dữ liệu vật liệu. "
				L"Vui lòng đợi trong giây lát... (%d%s)",(int)(100*i/xdeDMVL),L"%");
			xl->PutStatusBar((_bstr_t)szval);

			f1 = prDMVL->GetItem(i,_colvl[1]);
			f2 = prDMVL->GetItem(i,_colvl[2]);
			f4 = prDMVL->GetItem(i,_colvl[4]);
			if (f2 != (_bstr_t)L"GD" && f2 != (_bstr_t)L"HM")
			{
				if (_wtoi(f1) > 0 && f4 != L"")
				{
					// Kiểm tra nếu dòng (i) có mã vật liệu -> là vật liệu gốc
					val1 = prDMVL->GetItem(i, _colvl[7]);	// ngày
					val2 = prDMVL->GetItem(i, _colvl[4]);	// mã hsnt
					val3 = prDMVL->GetItem(i, _colvl[21]);	// xuất xứ

					// Kiểm tra xem mã vật liệu có trùng không?
					bf = TRUE;
					for (int j = 1; j <= _arr; j++)
					{
						if (_MyVL[j].sMa == f2)
						{
							bf = FALSE;
							_gan = j;
							break;
						}
					}

					// Không trùng --> Thêm vào mảng
					if (bf == TRUE)
					{
						_arr++;
						_MyVL[_arr].nDG = 0;
						_MyVL[_arr].sMa = prDMVL->GetItem(i, _colvl[2]);		// mã vl
						_MyVL[_arr].sTen = prDMVL->GetItem(i, _colvl[5]);		// tên
						_gan = _arr;
					}
				}
				else
				{
					// Dòng không có mã vật liệu
					// Phân tích diễn giải
					f5 = prDMVL->GetItem(i, _colvl[5]);	// tên
					if (f5 != L"")
					{
						_MyVL[_gan].nDG++;
						jd = _MyVL[_gan].nDG;

						// Xuất xứ
						msg = prDMVL->GetItem(i, _colvl[21]);
						msg.Trim();
						if (msg == L"") msg = (LPCTSTR)val3;
						_MyVL[_gan].sXuatxu[jd] = msg;

						// Kiểm tra diễn giải đó có ghi ngày không?
						msg = prDMVL->GetItem(i, _colvl[7]);		// ngày
						msg.Trim();
						if (msg == L"") msg = (LPCTSTR)val1;

						_MyVL[_gan].sNgay[jd] = (_bstr_t)msg;	// ngày
						_MyVL[_gan].sBB[jd] = val2;				// mã hsnt	

						// Leo 23.05.2018
						// Kiểm tra diễn giải có comment không?
						PRS = prDMVL->GetItem(i, _colvl[5]);
						if (PRS->GetComment() != NULL)
						{
							// Lấy mẫu ngang
							f_VLNV_Laymaungang(i, _colvl[5], _gan, jd);
						}
						else
						{
							_MyVL[_gan].sDG[jd] = (_bstr_t)f_tachcv(f5, 1);						// Diễn giải
							_MyVL[_gan].sMAC[jd] = (_bstr_t)f_tachcv(f5, 3);					// Mác
							_MyVL[_gan].sDVI[jd] = (_bstr_t)f_tachklg(f_tachcv(f5, 2), 2);		// đơn vị
							_MyVL[_gan].sKLG[jd] = (_bstr_t)f_tachklg(f_tachcv(f5, 2), 1);		// Khối lượng
							_MyVL[_gan].sGchu[jd] = (_bstr_t)L"";
						}
					}
				}
			}			
		}

		ObjConn.CloseDatabase();

		int iluu = 0;
		int vtchen = irowbd;
		for (int i = 1; i <= _arr; i++)
		{
			szval.Format(
				L"Đang xuất dữ liệu vật liệu nhập về... (%d%s)",
				(int)(100*i/_arr),L"%");
			xl->PutStatusBar((_bstr_t)szval);

			if(_MyVL[i].nDG>0)
			{
				// Tiêu đề
				PRS = psVLNV->GetRange(prVLNV->GetItem(4,_colnv[1]),prVLNV->GetItem(5,_colnv[13]));
				PRS->Copy(prVLNV->GetItem(vtchen,_colnv[1]));
				xl->PutCutCopyMode(XlCutCopyMode(false));

				// Định dạng khung
				vtchen+=2;
				PRS = psVLNV->GetRange(prVLNV->GetItem(vtchen,_colnv[1]),prVLNV->GetItem(vtchen+_MyVL[i].nDG,_colnv[13]));
				PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->Weight=xlThin;
				PRS->Borders->GetItem(XlBordersIndex::xlEdgeLeft)->Weight=xlThin;
				PRS->Borders->GetItem(XlBordersIndex::xlEdgeRight)->Weight=xlThin;
				PRS->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->Weight=xlThin;
				PRS->Borders->GetItem(XlBordersIndex::xlInsideVertical)->Weight=xlThin;

				PRS->Borders->GetItem(XlBordersIndex::xlInsideHorizontal)->PutWeight(xlThin);
				PRS->Borders->GetItem(XlBordersIndex::xlInsideHorizontal)->PutLineStyle(xlDot);

				// Nội dung
				iluu = vtchen;
				for (int j = 1; j <= _MyVL[i].nDG; j++)
				{
					prVLNV->PutItem(vtchen,_colnv[1],j);
					prVLNV->PutItem(vtchen,_colnv[3],_MyVL[i].sMa);
					prVLNV->PutItem(vtchen,_colnv[5],_MyVL[i].sTen);
					prVLNV->PutItem(vtchen,_colnv[14],_MyVL[i].sXuatxu[j]);	// Xuất xứ

					msg = (LPCTSTR)_MyVL[i].sNgay[j];
					msg.Trim();
					f_ktra_date(msg,prVLNV,vtchen,_colnv[2]);				// Ngày

					szval.Format(L"'%s",(LPCTSTR)_MyVL[i].sBB[j]);
					prVLNV->PutItem(vtchen,_colnv[4],(_bstr_t)szval);		// Mã HSNT

					prVLNV->PutItem(vtchen,_colnv[6],_MyVL[i].sDG[j]);		// Diễn giải
					prVLNV->PutItem(vtchen,_colnv[7],_MyVL[i].sMAC[j]);		// Mác
					prVLNV->PutItem(vtchen,_colnv[8],_MyVL[i].sDVI[j]);		// đơn vị
					prVLNV->PutItem(vtchen,_colnv[9],_MyVL[i].sKLG[j]);		// Khối lượng
					prVLNV->PutItem(vtchen,_colnv[13],_MyVL[i].sGchu[j]);	// Ghi chú

					vtchen++;
				}

				// Dòng cần để trống để lọc subtotal
				PRS = psVLNV->GetRange(prVLNV->GetItem(vtchen,_colnv[1]),prVLNV->GetItem(vtchen,_colnv[1]));
				PRS->EntireRow->RowHeight = 0;

				// Dòng tổng cộng
				prVLNV->PutItem(vtchen+1,_colnv[8],(_bstr_t)L"Tổng = ");
				msg.Format(L"=SUBTOTAL(9,%s:%s)",GAR(iluu,_colnv[9],prVLNV,0),GAR(vtchen,_colnv[9],prVLNV,0));
				prVLNV->PutItem(vtchen+1,_colnv[9],(_bstr_t)msg);

				PRS = psVLNV->GetRange(prVLNV->GetItem(vtchen+1,_colnv[1]),prVLNV->GetItem(vtchen+1,_colnv[13]));
				PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->Weight=xlThin;
				PRS->Borders->GetItem(XlBordersIndex::xlEdgeLeft)->Weight=xlThin;
				PRS->Borders->GetItem(XlBordersIndex::xlEdgeRight)->Weight=xlThin;
				PRS->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->Weight=xlThin;
				PRS->Font->PutBold(TRUE);
				PRS->Interior->PutColor(RGB(225,225,225));

				vtchen+=4; 
			}
		}

		// Comment END
		vtchen-=2;
		PRS=psVLNV->GetRange(prVLNV->GetItem(vtchen,_colnv[1]), prVLNV->GetItem(vtchen,_colnv[1]));
		PRS->ClearComments();
		PRS->AddComment((_bstr_t)L"END");

		// Kiểm tra ẩn hiện sheet
		int iVisible = psVLNV->GetVisible();
		if (iVisible!=-1) psVLNV->PutVisible(XlSheetVisibility::xlSheetVisible);
		psVLNV->Activate();

		//MessageBox(NULL,L"Xuất bảng vật liệu thành công.",L"Thông báo",MB_OK | MB_ICONINFORMATION);

		// Lựa chọn vùng in
		/*psVLNV->ResetAllPageBreaks();

		msg.Format(L"%s:%s",GAR(1,_colnv[1],prVLNV,0),GAR(vtchen-1,_colnv[13],prVLNV,0));
		psVLNV->PageSetup->PrintArea = (_bstr_t)msg;
		xl->ActiveWindow->PutView(xlPageBreakPreview);

		PageSetupPtr _psetPrt = psVLNV->PageSetup;
		_psetPrt->PutOrientation(xlLandscape);
		_psetPrt->PutZoom(100);

		int _idoc = psVLNV->VPageBreaks->GetCount()+1;
		if(_idoc>=2) psVLNV->VPageBreaks->GetItem(1)->DragOff(xlToRight,1);*/

		xl->PutStatusBar((_bstr_t)L"Ready");
		CoUninitialize();
	}
	catch(_com_error & error){_s(L"#QL6.54: Lỗi không thể ghi dữ liệu nhập về.");}
}


void cls_main::f_VLNV_Laymaungang(int irow, int icol, int gan, int num)
{
	try
	{
		_MyVL[gan].sDG[num] = L"";
		_MyVL[gan].sMAC[num] = L"";
		_MyVL[gan].sDVI[num] = L"";
		_MyVL[gan].sKLG[num] = L"";
		_MyVL[gan].sGchu[num] = L"";

		// Tách tên mẫu
		CString szval=L"",sChitiet=L"",sGchu=L"";
		CString sadd= GIOR(irow,icol,prDMVL,L"");

		int _pos = sadd.Find(L"|");
		if(_pos>0)
		{
			szval = sadd.Left(_pos);
			szval.Trim();
			sChitiet = sadd.Right(sadd.GetLength()-_pos-1);
			_pos = sChitiet.Find(_rowheight);
			if(_pos>0) sChitiet = sChitiet.Left(_pos);

			SqlString.Format(L"SELECT * FROM tbTenMau WHERE TenMau LIKE '%s';",szval);
			CRecordset* Recset = ObjConn.Execute(SqlString);
			while( !Recset->IsEOF() )
			{
				Recset->GetFieldValue(L"GChu",sGchu);
				_pos = sGchu.Find(L"@");
				if(_pos>0) sGchu = sGchu.Left(_pos);
				break;
			}
			Recset->Close();
		}

		if(sChitiet==L"" || sGchu==L"") return;		

		// Số lượng hàng cột
		int slcot=0;
		CString sGtCot[20];
		
		if(sGchu.Right(1)!=L"|") sGchu+=L"|";
		int vt=0;
		CString sel=L"";
		for (int i = 0; i < sGchu.GetLength(); i++)
		{
			sel = sGchu.Mid(i,1);
			if(sel==L"|" && i>vt)
			{
				// Số lượng phần tử
				slcot++;
				sGtCot[slcot] = sGchu.Mid(vt,i-vt);
				vt=i+1;
			}
		}

		int iKExl=0;
		CString sGtExl[1000];
		if(sChitiet.Right(1)!=L";") sChitiet+=L";";
		vt=0;
		sel=L"";
		for (int i = 0; i < sChitiet.GetLength(); i++)
		{
			sel = sChitiet.Mid(i,1);
			if(sel==L";" && i>vt)
			{
				// Số lượng phần tử
				iKExl++;
				sGtExl[iKExl] = sChitiet.Mid(vt,i-vt);
				vt=i+1;
			}
		}		

		int slhang = (int)(iKExl/slcot);
		for (int i = 0; i < slhang; i++)
		{
			_MyVL[gan].sDG[num+i] = L"";
			_MyVL[gan].sMAC[num+i] = L"";
			_MyVL[gan].sDVI[num+i] = L"";
			_MyVL[gan].sKLG[num+i] = L"";

			if(i>0)
			{
				_MyVL[gan].sXuatxu[num+i] = _MyVL[gan].sXuatxu[num];
				_MyVL[gan].sNgay[num+i] = _MyVL[gan].sNgay[num];
				_MyVL[gan].sBB[num+i] = _MyVL[gan].sBB[num];
			}
		}

		_MyVL[gan].nDG+=(slhang-1);	

		// Kiểm tra diễn giải (lấy cột 2)
		_pos=0;
		for (int j = 1; j <= iKExl; j++)
		{
			if (j%slcot == 2)
			{
				_MyVL[gan].sDG[num + _pos] = sGtExl[j];
				_pos++;
			}
		}

		// Kiểm tra mác thiết kế
		for (int i = 1; i <= slcot; i++)
		{
			if (sGtCot[i].Left(3) == L"Mác")
			{
				_pos = 0;
				for (int j = 1; j <= iKExl; j++)
					if (j%slcot == i || (i == slcot && j%slcot == 0))
					{
						_MyVL[gan].sMAC[num + _pos] = sGtExl[j];
						_pos++;
					}

				break;
			}
		}

		// Kiểm tra khối lượng
		int iTak=-1;
		for (int i = 1; i <= slcot; i++)
		{
			if (sGtCot[i].Left(10) == L"Khối lượng" || sGtCot[i].Left(2) == L"KL")
			{
				_pos = 0;
				for (int j = 1; j <= iKExl; j++)
					if (j%slcot == i || (i == slcot && j%slcot == 0))
					{
						iTak = sGtExl[j].Find(L" ");
						if (iTak > 0)
						{
							_MyVL[gan].sKLG[num + _pos] = sGtExl[j].Left(iTak);
							_MyVL[gan].sDVI[num + _pos] = sGtExl[j].Right(sGtExl[j].GetLength() - iTak - 1);
						}
						else _MyVL[gan].sKLG[num + _pos] = sGtExl[j];

						_pos++;
					}

				break;
			}
		}

		// Kiểm tra đơn vị tính		
		for (int i = 1; i <= slcot; i++)
		{
			if (sGtCot[i].Left(6) == L"Đơn vị" || sGtCot[i].Left(3) == L"ĐVT")
			{
				_pos = 0;
				for (int j = 1; j <= iKExl; j++)
					if (j%slcot == i || (i == slcot && j%slcot == 0))
					{
						_MyVL[gan].sDVI[num + _pos] = sGtExl[j];
						_pos++;
					}

				break;
			}
		}		

		// Kiểm tra ghi chú		
		for (int i = 1; i <= slcot; i++)
		{
			if (sGtCot[i].Left(7) == L"Ghi chú")
			{
				_pos = 0;
				for (int j = 1; j <= iKExl; j++)
					if (j%slcot == i || (i == slcot && j%slcot == 0))
					{
						_MyVL[gan].sGchu[num + _pos] = sGtExl[j];
						_pos++;
					}

				break;
			}
		}			
	}
	catch(...){}
}


void cls_main::f_autoFilter_VLNV()
{
	try
	{
		CoInitialize(NULL);
		xl.GetActiveObject(L"Excel.Application");
		xl->Visible = true;
		pWb = xl->GetActiveWorkbook();

		pSheet = pWb->ActiveSheet;
		pRange = pSheet->Cells;
		_code = pSheet->CodeName;

		if(_code == (_bstr_t)L"Sheet8")
		{
			// Xác định cột vật liệu nhập về
			psVLNV = pSheet;
			f_xacdinh_cot_VLNV();

			// Xác định dòng vs cột hiện hành
			int iRow = pSheet->Application->ActiveCell->Row;

			// Duyệt tìm 'STT'
			int lui=iRow;
			CString gtri = pRange->GetItem(lui,_colnv[1]);
			while (gtri!=L"STT" && lui>1)
			{
				lui--;
				gtri = pRange->GetItem(lui,_colnv[1]);
			}

			if(lui>0)
			{
				PRS = pSheet->GetRange(pRange->GetItem(lui+1,_colnv[1]),pRange->GetItem(lui+1,_colnv[13]));

				int ktra = pSheet->GetAutoFilterMode();
				if(ktra==-1) pSheet->PutAutoFilterMode(FALSE);
				PRS->AutoFilter(1,vtMissing,XlAutoFilterOperator::xlAnd,vtMissing,TRUE);
			}
		}

		CoUninitialize();
	}
	catch(_com_error & error){_s(L"#QL4.42: Lỗi lọc dữ liệu vật liệu nhập về.");}
}


void cls_main::f_load_stt_bienban()
{
	try
	{
		CoInitialize(NULL);
		xl.GetActiveObject(L"Excel.Application");
		xl->Visible = true;
		xl->PutScreenUpdating(VARIANT_FALSE);
		xl->StatusBar = (_bstr_t)L"Đang tiến hành phân tích dữ liệu. Vui lòng đợi trong giây lát...";

		pWb = xl->GetActiveWorkbook();
		pSheet = pWb->ActiveSheet;
		pRange = pSheet->Cells;
		_code = pSheet->CodeName;

		if(_code == (_bstr_t)L"shHSNTVL" || _code == (_bstr_t)L"shHSNTCV" || _code == (_bstr_t)L"shHSNTGD")
		{
			int nSht,col1,_rSp[8],_cSp[8];
			_bstr_t _sh[8];
			_WorksheetPtr _wsh[8];
			RangePtr _pr[8];

			if(_code == (_bstr_t)L"shHSNTVL")
			{
				nSht=4;
				col1 = getColumn(pSheet,"DMVL_STT");
				_sh[1] = name_sheet("shHSNTVLLAYM");		// lấy về
				_sh[2] = name_sheet("shHSNTVLNBVL");		// nội bộ
				_sh[3] = name_sheet("shHSNTVLYVL");			// yêu cầu
				_sh[4] = name_sheet("shHSNTVLNVL");			// nghiệm thu
			}
			else if(_code == (_bstr_t)L"shHSNTCV")
			{
				nSht=6;
				col1 = getColumn(pSheet,"DMBB_STT");
				_sh[1] = name_sheet("shHSNTCVLMTN");
				_sh[2] = name_sheet("shHSNTCVNBCV1");
				_sh[3] = name_sheet("shHSNTCVNBCV2");
				_sh[4] = name_sheet("shHSNTCVNBCV");
				_sh[5] = name_sheet("shHSNTCVYNCV");
				_sh[6] = name_sheet("shHSNTCVNTCV");
			}
			else if(_code == (_bstr_t)L"shHSNTGD")
			{
				nSht=3;
				col1 = getColumn(pSheet,"DMGD_STT");
				_sh[1] = name_sheet("shHSNTGDNTBGD1");
				_sh[2] = name_sheet("shHSNTGDYCNT1");
				_sh[3] = name_sheet("shHSNTGDNTGD1");
			}

			for (int i = 1; i <= nSht; i++)
			{
				_wsh[i] = xl->Sheets->GetItem(_sh[i]);
				_pr[i] = _wsh[i]->Cells;

				_code = _wsh[i]->CodeName;
				if(_code==(_bstr_t)L"shHSNTVLLAYM")
				{
					// Vật liệu
					_rSp[1] = getRow(pSheet,"MAUVL_BB");
					_cSp[1] = getColumn(pSheet,"MAUVL_BB");
				}
				else if(_code==(_bstr_t)L"shHSNTVLNBVL")
				{
					_rSp[2] = getRow(pSheet,"NTNBVL_BB");
					_cSp[2] = getColumn(pSheet,"NTNBVL_BB");
				}
				else if(_code==(_bstr_t)L"shHSNTVLYVL")
				{
					_rSp[3] = getRow(pSheet,"YCVL_BB");
					_cSp[3] = getColumn(pSheet,"YCVL_BB");
				}
				else if(_code==(_bstr_t)L"shHSNTVLNVL")
				{
					_rSp[4] = getRow(pSheet,"NTVL_BB");
					_cSp[4] = getColumn(pSheet,"NTVL_BB");
				}
				else if(_code==(_bstr_t)L"shHSNTCVLMTN")
				{
					// Công việc
					_rSp[1] = getRow(pSheet,"NLM_BB");
					_cSp[1] = getColumn(pSheet,"NLM_BB");
				}
				else if(_code==(_bstr_t)L"shHSNTCVNBCV1")
				{
					_rSp[2] = getRow(pSheet,"NTNB_BB");
					_cSp[2] = getColumn(pSheet,"NTNB_BB");
				}
				else if(_code==(_bstr_t)L"shHSNTCVNBCV2")
				{
					_rSp[3] = getRow(pSheet,"NTNBGD_BB");
					_cSp[3] = getColumn(pSheet,"NTNBGD_BB");
				}
				else if(_code==(_bstr_t)L"shHSNTCVNBCV")
				{
					_rSp[4] = getRow(pSheet,"NTNBCVGD_BB");
					_cSp[4] = getColumn(pSheet,"NTNBCVGD_BB");
				}
				else if(_code==(_bstr_t)L"shHSNTCVYNCV")
				{
					_rSp[5] = getRow(pSheet,"YCNTCV_BB");
					_cSp[5] = getColumn(pSheet,"YCNTCV_BB");
				}
				else if(_code==(_bstr_t)L"shHSNTCVNTCV")
				{
					_rSp[6] = getRow(pSheet,"NTCV_BB");
					_cSp[6] = getColumn(pSheet,"NTCV_BB");
				}
				else if(_code==(_bstr_t)L"shHSNTGDNTBGD1")
				{
					// Giai đoạn
					_rSp[1] = getRow(pSheet,"NTNBGDG_BB");
					_cSp[1] = getColumn(pSheet,"NTNBGDG_BB");
				}
				else if(_code==(_bstr_t)L"shHSNTGDYCNT1")
				{
					_rSp[2] = getRow(pSheet,"YCGD_BB");
					_cSp[2] = getColumn(pSheet,"YCGD_BB");
				}
				else if(_code==(_bstr_t)L"shHSNTGDNTGD1")
				{
					_rSp[3] = getRow(pSheet,"NTGD_BB");
					_cSp[3] = getColumn(pSheet,"NTGD_BB");
				}
			}

			// Xác định dòng vs cột hiện hành
			int iRow = pSheet->Application->ActiveCell->Row;
			int iCol = pSheet->Application->ActiveCell->Column;

			CString gtri = pRange->GetItem(iRow,iCol);

			// Xác định END
			int xde = FindComment(pSheet,col1,"END");
			if(iRow<xde && _wtoi(gtri)>0) for (int i = 1; i <= nSht; i++) _pr[i]->PutItem(_rSp[i],_cSp[i],(_bstr_t)gtri);
		}

		MessageBox(NULL,L"Thông tin đã được load thành công.",L"Thông báo",MB_OK | MB_ICONINFORMATION); 
		//xl->PutScreenUpdating(VARIANT_TRUE);
		xl->StatusBar = (_bstr_t)L"Ready";
		CoUninitialize();
	}
	catch(_com_error & error){_s(L"#QL4.41: Lỗi đánh số thứ tự.");}
}


void cls_main::f_capnhat_stt()
{
	try
	{
		CoInitialize(NULL);
		xl.GetActiveObject(L"Excel.Application");
		xl->Visible = true;
		xl->PutScreenUpdating(VARIANT_FALSE);
		pWb = xl->GetActiveWorkbook();

		f_stt((_bstr_t)L"shHSNTVL",8);	// Shet danh mục NT vật liệu
		f_stt((_bstr_t)L"shHSNTCV",8);	// Shet danh mục danh muc NT công việc
		f_stt((_bstr_t)L"shHSNTGD",8);	// Shet danh mục danh muc NT giai đoạn

		//xl->PutScreenUpdating(VARIANT_TRUE);
		CoUninitialize();
	}
	catch(_com_error & error){_s(L"#QL4.98: Lỗi cập nhật số thứ tự.");}
}


// hàm cập nhật số thứ tự công việc
void cls_main::f_stt(_bstr_t fSheet,int ibegin)
{
	try
	{
		_code = name_sheet(fSheet);
		pSheet = xl->Sheets->GetItem(_code);
		pRange = pSheet->Cells;
	
		_bstr_t mediateBSTR;
		wchar_t *stt_end;
		stt_end=L"";

		int col1,col2;
		if(fSheet== (_bstr_t)L"shHSNTVL")
		{
			col1 = getColumn(pSheet,"DMVL_STT");
			col2 = getColumn(pSheet,"DMVL_MAVL");
		}
		else if(fSheet== (_bstr_t)L"shHSNTCV")
		{
			col1 = getColumn(pSheet,"DMBB_STT");
			col2 = getColumn(pSheet,"DMBB_MACV");
		}
		else if(fSheet== (_bstr_t)L"shHSNTGD")
		{
			col1 = getColumn(pSheet,"DMGD_STT");
			col2 = getColumn(pSheet,"DMGD_HS");
		}

		int xde = FindComment(pSheet,col1,"END");
		int dem=0;
		for(int i=ibegin;i<xde;i++)
		{
			msg = pRange->GetItem(i,col1);
			if(_wtoi(msg)>0)
			{
				msg = pRange->GetItem(i,col2);
				if(msg!=L"" && msg.Find(L"*")==-1)
				{
					dem++;
					pRange->PutItem(i,col1,dem);
				}
				else pRange->PutItem(i,col1,(_bstr_t)L"");
			}
		}
	}
	catch(_com_error & error){_s(L"#QL4.99: Lỗi cập nhật số thứ tự.");}
}


CString cls_main::F3so(CString vl ,int num)
{
	CString szval=L"";
	if(num<9) szval.Format(L"%s00%d",vl,num);
	else if(num>=9 && num<99) szval.Format(L"%s0%d",vl,num);
	else szval.Format(L"%s%d",vl,num);

	return szval;
}

void cls_main::f_truyen_dulieu_DMHS()
{
	try
	{
		xl->PutScreenUpdating(VARIANT_FALSE);
		xl->StatusBar = (_bstr_t)L"Đang tiến hành phân tích dữ liệu. Vui lòng đợi trong giây lát...";

		shDMHS = name_sheet("Sheet6");
		psDMHS=xl->Sheets->GetItem(shDMHS);
		prDMHS = psDMHS->Cells;

		// Định nghĩa các sheet
		shDMVL = name_sheet("shHSNTVL");
		psDMVL=xl->Sheets->GetItem(shDMVL);
		prDMVL = psDMVL->Cells;

		shDMCV = name_sheet("shHSNTCV");
		psDMCV=xl->Sheets->GetItem(shDMCV);
		prDMCV = psDMCV->Cells;

		shDMGD = name_sheet("shHSNTGD");
		psDMGD=xl->Sheets->GetItem(shDMGD);
		prDMGD = psDMGD->Cells;

		// Xác định cột các sheet
		f_xacdinh_cot_DMHS();
		f_xacdinh_cot_DMVL();
		f_xacdinh_cot_DMCV();
		f_xacdinh_cot_DMGD();

		// Xác định các dòng danh mục
		int jvl = getRow(psDMHS,"DMHSVL");
		int jcv = getRow(psDMHS,"DMHSCV");
		int jgd = getRow(psDMHS,"DMHSGD");

		// Xác định vị trí END
		int xde = FindComment(psDMHS,_colhs[1],"END");

		// Xóa dữ liệu cũ
		if(xde-jgd>1) psDMHS->GetRange(prDMHS->GetItem(xde-1,1),prDMHS->GetItem(jgd+1,1))->EntireRow->Delete(xlShiftUp);
		if(jgd-jcv>1) psDMHS->GetRange(prDMHS->GetItem(jgd-1,1),prDMHS->GetItem(jcv+1,1))->EntireRow->Delete(xlShiftUp);
		if(jcv-jvl>1) psDMHS->GetRange(prDMHS->GetItem(jcv-1,1),prDMHS->GetItem(jvl+1,1))->EntireRow->Delete(xlShiftUp);

		// Chèn dòng từ dưới lên
		f_xacdinh_DMHS((_bstr_t)"shHSNTGD",
			_colgd[1],8,50,_colgd[5],50,_colgd[4],50,_colgd[11],_colgd[16],0,_colgd[10],jvl+2);
		
		f_xacdinh_DMHS((_bstr_t)"shHSNTCV",
			_colcv[1],8,_colcv[2],_colcv[5],_colcv[6],_colcv[4],_colcv[9],_colcv[15],_colcv[27],_colcv[18],_colcv[14],jvl+1);
		
		f_xacdinh_DMHS((_bstr_t)"shHSNTVL",
			_colvl[1],8,_colvl[2],_colvl[5],_colvl[7],_colvl[4],_colvl[9],_colvl[15],_colvl[22],0,_colvl[14],jvl);

		// Trường hợp sắp xếp theo đợt thanh toán
		if(iscbb==4 && ichkpnhom==1)
		{
			jvl = getRow(psDMHS,"DMHSVL");
			jcv = getRow(psDMHS,"DMHSCV");
			jgd = getRow(psDMHS,"DMHSGD");
			xde = FindComment(psDMHS,_colhs[1],"END");

			CString stren=L"",sduoi=L"";
			if(xde>jgd+1)
			for (int i = xde-2; i >= jgd; i--)
			{
				stren = GIOR(i,_colhs[9]+3,prDMHS,L"");
				sduoi = GIOR(i+1,_colhs[9]+3,prDMHS,L"");
				if(stren!=sduoi)
				{
					PRS = prDMHS->GetItem(i+1,1);
					PRS->EntireRow->Insert(xlShiftDown);

					if(sduoi==L"") stren = L"CHƯA NHẬP ĐỢT THANH TOÁN";
					else stren.Format(L"THANH TOÁN ĐỢT: %s",sduoi);
					prDMHS->PutItem(i+1,_colhs[3],(_bstr_t)stren);
					
					PRS = psDMHS->GetRange(prDMHS->GetItem(i+1,1),prDMHS->GetItem(i+1,_colhs[9] + 3));
					PRS->Font->PutBold(1);
					PRS->Interior->PutColor(LIGHTGRAY);
				}
			}

			if(jgd>jcv+1)
			for (int i = jgd-2; i >= jcv; i--)
			{
				stren = GIOR(i,_colhs[9]+3,prDMHS,L"");
				sduoi = GIOR(i+1,_colhs[9]+3,prDMHS,L"");
				if(stren!=sduoi)
				{
					PRS = prDMHS->GetItem(i+1,1);
					PRS->EntireRow->Insert(xlShiftDown);

					if(sduoi==L"") stren = L"CHƯA NHẬP ĐỢT THANH TOÁN";
					else stren.Format(L"THANH TOÁN ĐỢT: %s",sduoi);
					prDMHS->PutItem(i+1,_colhs[3],(_bstr_t)stren);
					
					PRS = psDMHS->GetRange(prDMHS->GetItem(i+1,1),prDMHS->GetItem(i+1,_colhs[9] + 3));
					PRS->Font->PutBold(1);
					PRS->Interior->PutColor(LIGHTGRAY);
				}
			}

			if(jcv>jvl+1)
			for (int i = jcv-2; i >= jvl; i--)
			{
				stren = GIOR(i,_colhs[9]+3,prDMHS,L"");
				sduoi = GIOR(i+1,_colhs[9]+3,prDMHS,L"");
				if(stren!=sduoi)
				{
					PRS = prDMHS->GetItem(i+1,1);
					PRS->EntireRow->Insert(xlShiftDown);

					if(sduoi==L"") stren = L"CHƯA NHẬP ĐỢT THANH TOÁN";
					else stren.Format(L"THANH TOÁN ĐỢT: %s",sduoi);
					prDMHS->PutItem(i+1,_colhs[3],(_bstr_t)stren);
					
					PRS = psDMHS->GetRange(prDMHS->GetItem(i+1,1),prDMHS->GetItem(i+1,_colhs[9] + 3));
					PRS->Font->PutBold(1);
					PRS->Interior->PutColor(LIGHTGRAY);
				}
			}
		}

		// Kiểm tra ẩn hiện sheet
		int iVisible = psDMHS->GetVisible();
		if (iVisible!=-1) psDMHS->PutVisible(XlSheetVisibility::xlSheetVisible);
		psDMHS->Activate();

		//MessageBox(NULL,L"Dữ liệu được cập nhật thành công.",L"Thông báo",MB_OK | MB_ICONINFORMATION);

		//xl->PutScreenUpdating(VARIANT_TRUE);
		xl->StatusBar = (_bstr_t)L"Ready";
		CoUninitialize();
	}
	catch(_com_error & error){_s(L"#QL4.45: Lỗi xuất dữ liệu sang danh mục hồ sơ.");}
}


// hàm xác định công việc, chèn dòng và truyền dữ liệu công việc vào biến gán
void cls_main::f_xacdinh_DMHS(_bstr_t fSheet,int col1,int ibegin,
	int vt1,int vt2,int vt3,int vt4,int vt5,int vt6,int vt7,int vt8,int vt9,int ichen)
{
	// iscbb: n = 0	chèn theo số thứ tự; iscbb = 1	chèn theo mã công việc
	// Bổ sung 16.02.2017: iscbb=2	chèn theo ngày lấy mẫu	iscbb=3	chèn theo ngày nghiệm thu
	// Bổ sung 11.09.2018: iscbb=4	sắp xếp theo đợt
	// istang=1 sắp xếp tăng dần =0 giảm dần
	// ichen = vị trí chèn tại dòng tại sheet Danh mục hồ sơ 

	try
	{
		int jscbb = iscbb;
		
		// Trường hợp sắp xếp theo mã công việc thì nếu ở sheet giai đoạn --> đổi jscbb từ 1 thành 0 (Do không có mã)
		if(fSheet==(_bstr_t)"shHSNTGD" && jscbb==1) jscbb=0;

		_code = name_sheet(fSheet);
		pSheet=xl->Sheets->GetItem(_code);
		pRange = pSheet->Cells;

		// Xác định END
		int xde = FindComment(pSheet,col1,"END");

		// Xác định khối lượng công việc
		int jSTT=0,inum=0;
		CString szval=L"";

		// ichkpnhom=0 --> Không phân nhóm hạng mục hoặc đợt thanh toán
		if(ichkpnhom==0)
		{			
			for(int i=ibegin;i<xde;i++)
			{
				msg = GIOR(i,col1,pRange,L"");
				if(_wtoi(msg)>0)
				{
					Hoso[inum].dqs[12] = msg;
					Hoso[inum].dqs[0] = F3so(L"",inum+1);			//  số thứ tự gốc					
					Hoso[inum].dqs[10] = Hoso[inum].dqs[0];
					Hoso[inum].dqs[10] +=L"D";

					Hoso[inum].dqs[1] = GIOR(i,vt1,pRange,L"");		//	mã công việc 

					// nội dung
					szval = L"";
					if(iCheckTenTT==1) szval = GIOR(i, vt2 + 1, pRange, L"");
					if(szval==L"") szval = GIOR(i, vt2, pRange, L"");
					Hoso[inum].dqs[2] = szval;
					Hoso[inum].dqs[4] = GIOR(i, vt4, pRange, L"");		//	mã hsnt		

					Hoso[inum].dqs[3] = L"";							
					Hoso[inum].dqs[5] = L"";
					Hoso[inum].dqs[6] = L"";

					Hoso[inum].dqs[13] = L"";
					Hoso[inum].dqs[14] = L"";

					Hoso[inum].dqs[16] = L"";
					Hoso[inum].dqs[17] = L"";
					Hoso[inum].dqs[18] = L"";

					Hoso[inum].dqs[7] = GIOR(i,vt3,pRange,L"");		//	ngày nhập
					Hoso[inum].dqs[8] = GIOR(i,vt5,pRange,L"");		//	ngày lấy mẫu
					Hoso[inum].dqs[9] = GIOR(i,vt6,pRange,L"");		//	ngày nghiệm thu VL/CV

					Hoso[inum].dqs[7].Trim();
					Hoso[inum].dqs[8].Trim();
					Hoso[inum].dqs[9].Trim();					

					if(Hoso[inum].dqs[7]!=L"") Hoso[inum].dqs[3].Format(L"='%s'!%s",(LPCTSTR)_code,GAR(i,vt3,pRange,0));	//	ngày nhập
					if(Hoso[inum].dqs[8]!=L"") Hoso[inum].dqs[5].Format(L"='%s'!%s",(LPCTSTR)_code,GAR(i,vt5,pRange,0));	//	ngày lấy mẫu
					if(Hoso[inum].dqs[9]!=L"") Hoso[inum].dqs[6].Format(L"='%s'!%s",(LPCTSTR)_code,GAR(i,vt6,pRange,0));	//	ngày nghiệm thu VL/CV

					if (vt8 > 0)
					{
						Hoso[inum].dqs[13] = GIOR(i, vt8, pRange, L"");			//	ngày bắt đầu
						Hoso[inum].dqs[14] = GIOR(i, vt8 + 1, pRange, L"");		//	ngày kết thúc
					}
					
					Hoso[inum].dqs[15] = GIOR(i, vt9, pRange, L"");			//	ngày yêu cầu VL/CV

					Hoso[inum].dqs[13].Trim();
					Hoso[inum].dqs[14].Trim();
					Hoso[inum].dqs[15].Trim();

					if (Hoso[inum].dqs[13] != L"") Hoso[inum].dqs[16].Format(L"='%s'!%s", (LPCTSTR)_code, GAR(i, vt8, pRange, 0));	//	ngày bắt đầu
					if (Hoso[inum].dqs[14] != L"") Hoso[inum].dqs[17].Format(L"='%s'!%s", (LPCTSTR)_code, GAR(i, vt8 + 1, pRange, 0));	//	ngày kết thúc
					if (Hoso[inum].dqs[15] != L"") Hoso[inum].dqs[18].Format(L"='%s'!%s", (LPCTSTR)_code, GAR(i, vt9, pRange, 0));	//	ngày yêu cầu VL/CV

					// Bổ sung 11.09.2018
					Hoso[inum].dqs[11] = GIOR(i,vt7,pRange,L"");		//	sắp xếp theo đợt

					// Bổ sung 16.02.2017
					if(jscbb==1) Hoso[inum].ival.Format(L"%s_%s",Hoso[inum].dqs[1],Hoso[inum].dqs[0]);			// theo mã công việc
					else if(jscbb==4) Hoso[inum].ival.Format(L"%s_%s",Hoso[inum].dqs[11],Hoso[inum].dqs[0]);	// theo đợt thanh toán
					else if(jscbb==2 || jscbb==3)
					{
						// theo ngày lấy mẫu (jscbb=2) hoặc theo ngày nghiệm thu (jscbb=3)
						if(Hoso[inum].dqs[jscbb+6]==L"") Hoso[inum].ival = L"10";
						else Hoso[inum].ival.Format(L"%s%s%s",
							Hoso[inum].dqs[jscbb+6].Right(Hoso[inum].dqs[jscbb+6].GetLength()-6),
							Hoso[inum].dqs[jscbb+6].Mid(3,2),
							Hoso[inum].dqs[jscbb+6].Left(2));
					}					
					else Hoso[inum].ival = Hoso[inum].dqs[0];

					inum++;
				}
			}
		}
		else
		{
			int jslHM=0,jslGD=0,jslDOT=0;
			for(int i=ibegin;i<xde;i++)
			{
				msg = GIOR(i,col1,pRange,L"");
				szval = GIOR(i,vt1,pRange,L"");				
				if((jscbb==4 && _wtoi(msg)>0) || (jscbb<4 && (_wtoi(msg)>0 || szval==L"HM" || szval==L"GD")))
				{
					if(jscbb==4)
					{
						szval = GIOR(i,vt7,pRange,L"");
						jslDOT = _wtoi(szval);
						if(jslDOT<=0) jslDOT=999;
						Hoso[inum].dqs[10] = F3so(L"T",jslDOT);
					}
					else if(jscbb<4)
					{
						if(szval==L"HM")
						{
							jslHM++;
							jslGD=0;
						}
						else if(szval==L"GD") jslGD++;

						Hoso[inum].dqs[10] = F3so(L"A",jslHM) + F3so(L"B",jslGD);
					}					

					if(_wtoi(msg)>0)
					{
						Hoso[inum].dqs[10]+=L"D";
						Hoso[inum].dqs[0] = F3so(L"",inum+1);	//  số thứ tự gốc
						Hoso[inum].dqs[12] = msg;
					}
					else
					{
						Hoso[inum].dqs[10]+=L"C";
						Hoso[inum].dqs[0] = L"";				//  số thứ tự gốc (HM hoặc GD)
						Hoso[inum].dqs[12] = L"";
					}					

					Hoso[inum].dqs[1] = GIOR(i,vt1,pRange,L"");		//	mã công việc 

					// nội dung
					szval = L"";
					if (iCheckTenTT == 1) szval = GIOR(i, vt2 + 1, pRange, L"");
					if (szval == L"") szval = GIOR(i, vt2, pRange, L"");
					Hoso[inum].dqs[2] = szval;
					Hoso[inum].dqs[4] = GIOR(i, vt4, pRange, L"");		//	mã hsnt	

					Hoso[inum].dqs[3] = L"";								
					Hoso[inum].dqs[5] = L"";
					Hoso[inum].dqs[6] = L"";

					Hoso[inum].dqs[13] = L"";
					Hoso[inum].dqs[14] = L"";

					Hoso[inum].dqs[16] = L"";
					Hoso[inum].dqs[17] = L"";
					Hoso[inum].dqs[18] = L"";

					Hoso[inum].dqs[7] = GIOR(i,vt3,pRange,L"");		//	ngày nhập
					Hoso[inum].dqs[8] = GIOR(i,vt5,pRange,L"");		//	ngày lấy mẫu
					Hoso[inum].dqs[9] = GIOR(i,vt6,pRange,L"");		//	ngày nghiệm thu VL/CV

					Hoso[inum].dqs[7].Trim();
					Hoso[inum].dqs[8].Trim();
					Hoso[inum].dqs[9].Trim();

					if(Hoso[inum].dqs[7]!=L"") Hoso[inum].dqs[3].Format(L"='%s'!%s",(LPCTSTR)_code,GAR(i,vt3,pRange,0));	//	ngày nhập
					if(Hoso[inum].dqs[8]!=L"") Hoso[inum].dqs[5].Format(L"='%s'!%s",(LPCTSTR)_code,GAR(i,vt5,pRange,0));	//	ngày lấy mẫu
					if(Hoso[inum].dqs[9]!=L"") Hoso[inum].dqs[6].Format(L"='%s'!%s",(LPCTSTR)_code,GAR(i,vt6,pRange,0));	//	ngày nghiệm thu VL/CV

					if (vt8 > 0)
					{
						Hoso[inum].dqs[13] = GIOR(i, vt8, pRange, L"");			//	ngày bắt đầu
						Hoso[inum].dqs[14] = GIOR(i, vt8 + 1, pRange, L"");		//	ngày kết thúc
					}

					Hoso[inum].dqs[15] = GIOR(i, vt9, pRange, L"");		//	ngày yêu cầu VL/CV

					Hoso[inum].dqs[13].Trim();
					Hoso[inum].dqs[14].Trim();
					Hoso[inum].dqs[15].Trim();

					if (Hoso[inum].dqs[13] != L"") Hoso[inum].dqs[16].Format(L"='%s'!%s", (LPCTSTR)_code, GAR(i, vt8, pRange, 0));	//	ngày bắt đầu
					if (Hoso[inum].dqs[14] != L"") Hoso[inum].dqs[17].Format(L"='%s'!%s", (LPCTSTR)_code, GAR(i, vt8 + 1, pRange, 0));	//	ngày kết thúc
					if (Hoso[inum].dqs[15] != L"") Hoso[inum].dqs[18].Format(L"='%s'!%s", (LPCTSTR)_code, GAR(i, vt9, pRange, 0));	//	ngày yêu cầu VL/CV

					// Bổ sung 11.09.2018 (Leo)
					Hoso[inum].dqs[11] = GIOR(i,vt7,pRange,L"");		// Sắp xếp theo đợt

					// Bổ sung 16.02.2017 (Leo)
					if(jscbb==1) Hoso[inum].ival.Format(L"%s.%s_%s",Hoso[inum].dqs[10],Hoso[inum].dqs[1],Hoso[inum].dqs[0]);		// theo mã công việc
					else if(jscbb==4) Hoso[inum].ival.Format(L"%s.E%s_%s",Hoso[inum].dqs[10],Hoso[inum].dqs[11],Hoso[inum].dqs[0]);	// theo đợt thanh toán
					else if(jscbb==2 || jscbb==3)
					{
						// Theo ngày lấy mẫu (jscbb=2) hoặc theo ngày nghiệm thu (jscbb=3)
						if(Hoso[inum].dqs[jscbb+6]==L"") Hoso[inum].ival = Hoso[inum].dqs[10] + L".10";
						else
						{
							Hoso[inum].ival.Format(L"%s.%s%s%s",Hoso[inum].dqs[10],
								Hoso[inum].dqs[jscbb+6].Right(Hoso[inum].dqs[jscbb+6].GetLength()-6),
								Hoso[inum].dqs[jscbb+6].Mid(3,2),
								Hoso[inum].dqs[jscbb+6].Left(2));
						}
					}
					else Hoso[inum].ival.Format(L"%s.%s",Hoso[inum].dqs[10],Hoso[inum].dqs[0]);

					inum++;
				}
			}
		}		

///////// Sắp xếp Qsort (sắp xếp theo biến ival)
		if(jscbb>0) qsort(Hoso,inum,sizeof(MyQSort),compare_AZ);

		// Chèn dòng và truyền dữ liệu công việc vào biến gán
		if(inum>0)
		{
			psDMHS->GetRange(prDMHS->GetItem(ichen+1,1),prDMHS->GetItem(ichen+inum,1))->EntireRow->Insert(xlShiftDown);			

			jSTT=0;
			int k=0,gt=0;
			if(istang==1)
			{
				// Sắp xếp tăng dần
				for(int i=ichen+1;i<=ichen+inum;i++)	
				{
					if(Hoso[k].dqs[10].Right(1)==L"D")
					{
						jSTT++;
						prDMHS->PutItem(i,_colhs[1],jSTT);							//  số thứ tự
						psDMHS->GetRange(prDMHS->GetItem(i,1),prDMHS->GetItem(i,_colhs[9] + 3))->Font->PutBold(0);
					}
					else
					{
						psDMHS->GetRange(prDMHS->GetItem(i,1),prDMHS->GetItem(i,_colhs[9] + 3))->Font->PutBold(1);
						
						if(Hoso[k].dqs[1]==L"HM")
						psDMHS->GetRange(prDMHS->GetItem(i,1),prDMHS->GetItem(i,_colhs[9] + 3))->Interior->PutColor(LIGHTGRAY);
					}

					prDMHS->PutItem(i,_colhs[2],(_bstr_t)Hoso[k].dqs[1]);			//	mã công việc 
					prDMHS->PutItem(i,_colhs[3],(_bstr_t)Hoso[k].dqs[2]);			//	nội dung
					prDMHS->PutItem(i,_colhs[4],(_bstr_t)Hoso[k].dqs[3]);			//	mã nhập
					prDMHS->PutItem(i,_colhs[5],(_bstr_t)Hoso[k].dqs[4]);			//	mã hsnt
					prDMHS->PutItem(i,_colhs[6],(_bstr_t)Hoso[k].dqs[5]);			//	ngày lấy mẫu

					// Bổ sung 18.02.2021
					prDMHS->PutItem(i, _colhs[6] + 1, (_bstr_t)Hoso[k].dqs[16]);	//	ngày bắt đầu
					prDMHS->PutItem(i, _colhs[6] + 2, (_bstr_t)Hoso[k].dqs[17]);	//	ngày kết thúc
					prDMHS->PutItem(i, _colhs[6] + 3, (_bstr_t)Hoso[k].dqs[18]);	//	ngày yêu cầu

					prDMHS->PutItem(i,_colhs[7],(_bstr_t)Hoso[k].dqs[6]);			//	ngày nghiệm thu
					prDMHS->PutItem(i,_colhs[8],(_bstr_t)Hoso[k].dqs[12]);			//	số thứ tự gốc				
					prDMHS->PutItem(i,_colhs[9]+3,(_bstr_t)Hoso[k].dqs[11]);		//	đợt thanh toán

					k++;
				}
			}
			else
			{
				// Sắp xếp giảm dần
				for(int i=ichen+1;i<=ichen+inum;i++)	
				{
					if(Hoso[k].dqs[10].Right(1)==L"D")
					{
						jSTT++;
						prDMHS->PutItem(i,_colhs[1],jSTT);							//  số thứ tự
						psDMHS->GetRange(prDMHS->GetItem(i,1),prDMHS->GetItem(i,_colhs[9] + 3))->Font->PutBold(0);
					}
					else
					{
						psDMHS->GetRange(prDMHS->GetItem(i,1),prDMHS->GetItem(i,_colhs[9] + 3))->Font->PutBold(1);
						
						if(Hoso[k].dqs[1]==L"HM")
						psDMHS->GetRange(prDMHS->GetItem(i,1),prDMHS->GetItem(i,_colhs[9] + 3))->Interior->PutColor(LIGHTGRAY);
					}

					gt=inum-k-1;
					prDMHS->PutItem(i,_colhs[2],(_bstr_t)Hoso[gt].dqs[1]);			//	mã công việc 
					prDMHS->PutItem(i,_colhs[3],(_bstr_t)Hoso[gt].dqs[2]);			//	nội dung
					prDMHS->PutItem(i,_colhs[4],(_bstr_t)Hoso[gt].dqs[3]);			//	mã nhập
					prDMHS->PutItem(i,_colhs[5],(_bstr_t)Hoso[gt].dqs[4]);			//	mã hsnt
					prDMHS->PutItem(i,_colhs[6],(_bstr_t)Hoso[gt].dqs[5]);			//	mã lấy mẫu
					prDMHS->PutItem(i,_colhs[7],(_bstr_t)Hoso[gt].dqs[6]);			//	ngày nghiệm thu
					prDMHS->PutItem(i,_colhs[8],(_bstr_t)Hoso[gt].dqs[12]);			//	số thứ tự gốc				
					prDMHS->PutItem(i,_colhs[9]+3,(_bstr_t)Hoso[k].dqs[11]);		//	đợt thanh toán

					k++;
				}			
			}

			// Autofit
			PRS = psDMHS->GetRange(prDMHS->GetItem(ichen+1,_colhs[3]),prDMHS->GetItem(ichen+inum,_colhs[3]));
			PRS->EntireRow->Rows->AutoFit();
		}
	}
	catch(_com_error & error){_s(L"#QL4.33: Lỗi không thể xác định danh mục hồ sơ.");}
 }


void cls_main::f_SetRowHight(int _nRow, int _ibd, int _ikt)
{
	PRS = prNDNK->GetItem(_nRow,_ibd);
	PRS->UnMerge();

	PRS = prNDNK->GetItem(_nRow,_ibd);
	int ntrc = PRS->GetColumnWidth();
	double tong=0,nval=0;

	for (int i = _ibd; i <= _ikt; i++)
	{
		PRS = prNDNK->GetItem(_nRow,i);
		nval = PRS->GetColumnWidth();
		tong+=nval;	// Độ rộng ô thứ [i]
		tong+=0.71;	// Độ rộng vạch kẻ ngang
	}

	PRS = prNDNK->GetItem(_nRow,_ibd);
	PRS->PutHorizontalAlignment(xlLeft);

	PRS->EntireRow->Columns->AutoFit();
	nval = PRS->GetColumnWidth();
	int gtri = (int)(nval/tong)+1;
	if(nval>250) gtri = 3;	// Trong Excel chỉ có thể autofit trong khoảng [0,255]
	gtri = 18*gtri;

	PRS->PutColumnWidth(ntrc);
	PRS = psNDNK->GetRange(prNDNK->GetItem(_nRow,_ibd),prNDNK->GetItem(_nRow,_ikt));
	PRS->PutMergeCells(TRUE);
	PRS->PutWrapText(TRUE);
	PRS->PutHorizontalAlignment(xlJustify);
	PRS->PutRowHeight(gtri);
}

// chạy khi nhập trực tiếp hoặc nhấn spin hoặc quản lý > xuất số liệu nhật ký
void cls_main::f_xuat_nhat_ky()
{
	try
	{
		shTS = name_sheet("shTS");
		psTS=xl->Sheets->GetItem(shTS);
		prTS=psTS->GetCells();

		// Định nghĩa lấy mẫu
		keyLMTN = getVCell(psTS,L"CF_LMTN");
		keyLMTN.Trim();
		if(keyLMTN==L"") keyLMTN=L"LM";

		// Định nghĩa các sheet
		shNDNK = name_sheet("shMauNKY4");
		psNDNK=xl->Sheets->GetItem(shNDNK);
		prNDNK = psNDNK->Cells;

		shDMVL = name_sheet("shHSNTVL");
		psDMVL=xl->Sheets->GetItem(shDMVL);
		prDMVL = psDMVL->Cells;

		shDMCV = name_sheet("shHSNTCV");
		psDMCV=xl->Sheets->GetItem(shDMCV);
		prDMCV = psDMCV->Cells;

		shDMGD = name_sheet("shHSNTGD");
		psDMGD=xl->Sheets->GetItem(shDMGD);
		prDMGD = psDMGD->Cells;

		// Xác định cột các sheet
		f_xacdinh_cot_DMVL();
		f_xacdinh_cot_DMCV();
		f_xacdinh_cot_DMGD();

		// Xác định các vị trí
		_ivtr[0] = getColumn(psNDNK,"HK_A");	// HK_A mặc định là "Lấy mẫu thí nghiệm"
		_ivtr[1] = getRow(psNDNK,"HK_A");

		_ivtr[2] = getRow(psNDNK,"HK_B");		// HK_B mặc định là "Công việc thi công"
		_ivtr[3] = getRow(psNDNK,"HK_C");		// HK_C mặc định là "Nghiệm thu công việc"
		_ivtr[4] = getRow(psNDNK,"HK_D");		// HK_D mặc định là "Nghiệm thu vật liệu..."
		_ivtr[5] = getRow(psNDNK,"HK_E");		// HK_E mặc định là "Nghiệm thu hạng mục..."

		_ivtr[6] = FindComment(psNDNK,1,"END");	// Chưa đặt name cho cột cần xác định cmt END -> mặc định =1

		// Xác định ngày nghỉ lễ

		// Xóa dữ liệu ban đầu
		for (int i = 6; i > 1; i--)
			if(_ivtr[i]>_ivtr[i-1]+1)
			{
				PRS = psNDNK->GetRange(prNDNK->GetItem(_ivtr[i]-1,1),prNDNK->GetItem(_ivtr[i-1]+1,1));
				PRS->EntireRow->Delete(xlShiftUp);
			}

		// Xác định lại vị trí sau khi xóa
		_ivtr[2] = getRow(psNDNK,"HK_B");
		_ivtr[3] = getRow(psNDNK,"HK_C");
		_ivtr[4] = getRow(psNDNK,"HK_D");
		_ivtr[5] = getRow(psNDNK,"HK_E");
		_ivtr[6] = FindComment(psNDNK,1,"END");

		// Lấy giá trị ngày tháng muốn xuất
		int ir = getRow(psNDNK,"HK_NGAY");
		int ic = getColumn(psNDNK,"HK_NGAY");
		_ngaythang = prNDNK->GetItem(ir,ic);
		if(_ngaythang!=L"") numNgay = _ngaythang.GetLength();
		else numNgay=10;

		// Duyệt ngược từ 5->1
		CString _sArr[5] = {
			L"Lấy mẫu thí nghiệm",L"Công việc thi công",
			L"Nghiệm thu công việc",L"Nghiệm thu vật liệu",L"Nghiệm thu hạng mục"};

		int gtri=-1;
		int _pos=-1;
		CString szval = L"";
		for (int i = 5; i >= 1; i--)
		{
			szval = GIOR(_ivtr[i],_ivtr[0],prNDNK,L"");
			for (int j = 0; j < 5; j++)
			{
				_pos = szval.Find(_sArr[j]);
				if(_pos>=0)
				{
					gtri = j;
					break;
				}
			}

			// Nếu không tìm thấy, giá trị lấy mặc định bằng stt trong temp
			if(_pos==-1) gtri=i-1;

			_numNK = i;
			switch (gtri)
			{
				case 0:
					// 1.Lấy mẫu thí nghiệm
					// Chú ý phần lấy mẫu công việc phải có ký hiệu LM và ngày lấy mẫu ở dòng LM tương ứng
					f_XNKY_1(ir);
					break;

				case 1:
					// 2.Công việc thi công
					f_XNKY_2(ir);
					break;

				case 2:
					// 3.Nghiệm thu công việc
					f_XNKY_3(ir);
					break;

				case 3:
					// 4.Nghiệm thu vật liệu đưa vào sử dụng
					f_XNKY_4(ir);
					break;

				case 4:
					// 5.Nghiệm thu hạng mục công trình, giai đoạn thi công xây dựng
					f_XNKY_5(ir);
					break;

				default:
				break;
			}			
		}		

///////// Bổ sung load dữ liệu nhật ký (23.12.2016 | 19.01.2017 | 09.06.2018)
		// Đổ dữ liệu in hoặc không in
		SqlString.Format(L"SELECT * FROM tbNdung WHERE ngayghi LIKE '%s';",_ngaythang);
		
		if(getPathNHKY()==L"") return;
		if(connectDsn(pathNK)==-1) return;
		ObjConn.OpenAccessDB(pathNK, L"", L"");
		CRecordset* Recset = ObjConn.Execute(SqlString);

		szval=L"";
		while( !Recset->IsEOF() )
		{
			Recset->GetFieldValue(L"boqua",szval);
			break;
		}
		Recset->Close();
		if(szval==L"1") szval=L"Bỏ qua không in";
		prNDNK->PutItem(ir+3,ic,(_bstr_t)szval);

/// ------------------------------------
		f_load_csdl_nky(_ngaythang,ir);
/// ------------------------------------

		ObjConn.CloseDatabase();

///////////////////////////////////////////////////		

		// Kiểm tra ẩn hiện sheet
		int iVisible = psNDNK->GetVisible();
		if (iVisible!=-1) psNDNK->PutVisible(XlSheetVisibility::xlSheetVisible);
		psNDNK->Activate();		
	}
	catch(_com_error & error){_s(L"#QL4.53: Lỗi xuất nội dung nhật ký.");}
}


void cls_main::f_load_csdl_nky(CString sngay,int iRend)
{
	try
	{
		// Hiển thị nội dung hay không?
		shTS = name_sheet("shTS");
		psTS=xl->Sheets->GetItem(shTS);
		prTS=psTS->GetCells();

		int jxnd = _wtoi(getVCell(psTS,L"CF_XNDNK"));

		// Hiển thị các điều kiện khác
		int jc1 = getRow(psNDNK,"NKBS_1");
		int jr1 = getColumn(psNDNK,"NKBS_1");
		CString szv = GIOR(1,jr1+10,prNDNK,L"");
		int iTp = _wtoi(szv);

		// jxnd=1 --> không quan tâm đến iTp
		// jxnd=0 --> Nếu iTp=1 -> Hiển thị điều kiện thời tiết,...
		if(jxnd==0  && iTp==0) return;		

/////// Load Checkbox ////////////////////////
		
		if(jxnd==1 || (jxnd==0 && iTp!=0))
		{
			CString val1=L"",val2=L"",val3=L"",val4=L"";
			SqlString.Format(L"SELECT * FROM tbNdung WHERE ngayghi LIKE '%s';",sngay);
			CRecordset* Recset = ObjConn.Execute(SqlString);

			while( !Recset->IsEOF() )
			{
				Recset->GetFieldValue(L"dkthoitiet",val1);
				Recset->GetFieldValue(L"nhietdo",val2);
				Recset->GetFieldValue(L"ctvsinh",val3);
				Recset->GetFieldValue(L"ctatldong",val4);
				break;
			}
			Recset->Close();

//////////// Điều kiện thời tiết			
			int vtri=0,dem=0;
			CString pK[5]={L"",L"",L"",L"",L""};
			if(val1!=L"")
			{
				if(val1.Right(1)!=L";") val1+=L";";
				for (int i = 0; i < val1.GetLength(); i++)
				{
					szv = val1.Mid(i,1);
					if(szv==L";" && i>=vtri)
					{
						dem++;						
						if(val1.Mid(i-1,1)==L";") pK[dem]=L"";
						else pK[dem] = val1.Mid(vtri,i-vtri);
						vtri=i+1;
						if(dem>=4) break;
					}
				}

				for(int i = 1; i <= 4; i++) prNDNK->PutItem(jc1+2,jr1+i,(_bstr_t)pK[i]);
			}

//////////// Điều kiện nhiệt độ
			vtri=0,dem=0;
			CString pX[5]={L"",L"",L"",L"",L""};
			if(val2!=L"")
			{
				if(val2.Right(1)!=L";") val2+=L";";
				for (int i = 0; i < val2.GetLength(); i++)
				{
					szv = val2.Mid(i,1);
					if(szv==L";" && i>=vtri)
					{
						dem++;						
						if(val2.Mid(i-1,1)==L";") pX[dem]=L"";
						else pX[dem] = val2.Mid(vtri,i-vtri);
						vtri=i+1;
						if(dem>=4) break;
					}
				}

				for(int i = 1; i <= 4; i++) prNDNK->PutItem(jc1+3,jr1+i,(_bstr_t)pX[i]);
			}

//////////// Công tác vệ sinh
			CString sVSinh[3]={L"Tốt",L"Bình thường",L"Kém"};			
			sVSinh[0] = GIOR(1,jr1+16,prNDNK,L"Tốt");
			sVSinh[1] = GIOR(2,jr1+16,prNDNK,L"Bình thường");
			sVSinh[2] = GIOR(3,jr1+16,prNDNK,L"Kém");

			if(val3==sVSinh[0]) f_chekbox1(L"Check Box 5",1,L"Check Box 7",0,L"Check Box 8",0);
			else if(val3==sVSinh[1]) f_chekbox1(L"Check Box 5",0,L"Check Box 7",1,L"Check Box 8",0);
			else if(val3==sVSinh[2]) f_chekbox1(L"Check Box 5",0,L"Check Box 7",0,L"Check Box 8",1);
			else f_chekbox1(L"Check Box 5",0,L"Check Box 7",0,L"Check Box 8",0);

//////////// An toàn lao động
			CString sATLD[3]={L"Tốt",L"Bình thường",L"Kém"};
			sATLD[0] = GIOR(1,jr1+17,prNDNK,L"Tốt");
			sATLD[1] = GIOR(2,jr1+17,prNDNK,L"Bình thường");
			sATLD[2] = GIOR(3,jr1+17,prNDNK,L"Kém");

			if(val4==sATLD[0]) f_chekbox1(L"Check Box 19",1,L"Check Box 20",0,L"Check Box 21",0);
			else if(val4==sATLD[1]) f_chekbox1(L"Check Box 19",0,L"Check Box 20",1,L"Check Box 21",0);
			else if(val4==sATLD[2]) f_chekbox1(L"Check Box 19",0,L"Check Box 20",0,L"Check Box 21",1);
			else f_chekbox1(L"Check Box 19",0,L"Check Box 20",0,L"Check Box 21",0);
		}

///////// LOAD CÁC NỘI DUNG CHÍNH --> Từ dưới lên trên //////////////////

		if(jxnd==1)
		{
			int jc32 = getRow(psNDNK,"NKBS_32");
			int jc33 = getRow(psNDNK,"NKBS_33");
			int jc34 = getRow(psNDNK,"NKBS_34");
			int jc35 = getRow(psNDNK,"NKBS_35");

			int jc4 = getRow(psNDNK,"NKBS_4");
			int jc6 = getRow(psNDNK,"NKBS_6");
			int jc7 = getRow(psNDNK,"NKBS_7");
			int jc8 = getRow(psNDNK,"NKBS_8");
			int jc9 = getRow(psNDNK,"NKBS_9");

			// Ý kiến của nhà thầu thi công
			f_load_ndung(sngay,jc8,jc9-1,jr1,1,L"tbYkNttc",0,1,iRend);

			// Nhận xét của chủ đầu tư
			f_load_ndung(sngay,jc7,jc8,jr1,1,L"tbNxCdt",0,1,iRend);

			// Nhận xét của cán bộ giám sát thi công
			f_load_ndung(sngay,jc6,jc7,jr1,1,L"tbNxGstc",0,1,iRend);

			// Tiến độ thực hiện
			f_load_ndung(sngay,jc35,jc4,jr1,1,L"tbTdth",0,1,iRend);

			// Chất lượng thi công
			f_load_ndung(sngay,jc34,jc35,jr1,1,L"tbCltc",0,1,iRend);

			// Tình trạng thực tế
			f_load_ndung(sngay,jc33,jc34,jr1,1,L"tbTTTTe",0,1,iRend);

			// Biện pháp thi công
			f_load_ndung(sngay,jc32,jc33,jr1,1,L"tbBptc",0,1,iRend);

////////// ~ . ~ . ~ . ~ . ~ . ~ . ~ . ~ . ~ . ~ . ~ . ~ . ~ . ~ . ~ . ~ . ~ . ~ . ~ . ~ . ~ . 
			// Bổ sung 24.06.2017 --> Load nhân công và máy thi công
			// Xác định các vị trí chứa nút
			int jc2 = getRow(psNDNK,"NKBS_2");
			int jc3 = getRow(psNDNK,"NKBS_3");
			int vt2=0;
			for (int i = jc2+2; i < jc3; i++)
			{
				szv = GIOR(i,jr1,prNDNK,L"");
				if(szv==L"STT")
				{
					vt2=i;
					break;
				}
			}

			// Vị trí chứa STT
			if(vt2==0) vt2 = jc3;

			int pos=-1;
			int jc22=0;
			if(vt2+1<jc3)
			for (int i = vt2+1; i < jc3; i++)
			{
				szv = GIOR(i,jr1,prNDNK,L"");
				pos = szv.Find(L"công");
				if(pos>0)
				{
					jc22=i;
					break;
				}
			}

			// Vị trí chứa thiết bị thi công
			if(jc22==0) jc22 = jc3;

			int vt3=0;
			if(jc22+1<jc3)
			for (int i = jc22+1; i < jc3; i++)
			{
				szv = GIOR(i,jr1,prNDNK,L"");
				if(szv==L"STT")
				{
					vt3=i;
					break;
				}
			}

			// Vị trí chứa STT thiết bị thi công
			if(vt3==0) vt3 = jc3;

			// B1- stt thiết bị thi công
			f_load_STT_NCM(sngay,vt3,jc3,jr1,1,L"tbSTTMtc");

			// B2- nội dung thiết bị thi công
			f_load_ndung(sngay,jc22,vt3,jr1,1,L"tbMtcong",1,0,iRend);

			// B3- stt nhân công
			f_load_STT_NCM(sngay,vt2,jc22,jr1,1,L"tbSTTNc");

			// B4- nội dung nhân công
			f_load_ndung(sngay,jc2+1,vt2,jr1,1,L"tbNcong",1,0,iRend);

////////// ~ . ~ . ~ . ~ . ~ . ~ . ~ . ~ . ~ . ~ . ~ . ~ . ~ . ~ . ~ . ~ . ~ . ~ . ~ . ~ . ~ . 
		}

/////////////////////////////////////////////////////////////////////////////////////////

	}
	catch(...){_s(L"Lỗi load dữ liệu nhật ký");}
}


void cls_main::f_load_STT_NCM(CString sNgayNK,int num1,int num2,int cot,int cong,CString frTable)
{
	try
	{
		// Kiểm tra dòng có bị ẩn không? --> Nếu ẩn thì return luôn
		PRS = prNDNK->GetItem(num1,1);
		int pos = PRS->GetRowHeight();
		if(pos==0) return;		

		if(num1+2<num2)
		{
			// Xóa toàn bộ dữ liệu cũ
			PRS = prNDNK->GetItem(num1+1,1);
			PRS->EntireRow->ClearContents();

			PRS = psNDNK->GetRange(prNDNK->GetItem(num1+2,1),prNDNK->GetItem(num2-1,1));
			PRS->EntireRow->Delete();
		}

		// Ghi dữ liệu mới
		CString szval=L"";
		CString sLoad[100],sfr1[100],sfr2[100],sfr3[100];
		SqlString.Format(L"SELECT * FROM %s WHERE ngayghi LIKE '%s' ORDER BY maso ASC;",frTable,sNgayNK);
		CRecordset* Recset = ObjConn.Execute(SqlString);
		int dem = 0;
		while( !Recset->IsEOF() )
		{
			dem++;

			// Tên nhân công / máy
			Recset->GetFieldValue(L"noidung",sLoad[dem]);
			Recset->GetFieldValue(L"ftr1",sfr1[dem]);
			Recset->GetFieldValue(L"ftr2",sfr2[dem]);
			Recset->GetFieldValue(L"ftr3",sfr3[dem]);

			Recset->MoveNext();
		}
		Recset->Close();

		// Xác định số dòng sẽ chèn thêm
		int sl=4;
		int ichk0=0;
		if(dem>0) sl=dem+1;
		else
		{
			PRS = prNDNK->GetItem(num1,cot);
			if(PRS->GetComment()!=NULL)
			{
				_bstr_t bscmt = PRS->GetComment()->Text();
				CString sadd = (LPCTSTR)(bscmt);
				sadd.Trim();
				sl = _wtoi(sadd);
				if(sl<=0)
				{
					ichk0=1;
					sl=1;
				}
				else if(sl>10) sl=10;
			}
		}

		PRS = psNDNK->GetRange(prNDNK->GetItem(num1+cong+1,1),prNDNK->GetItem(num1+cong+sl-1,1));
		PRS->EntireRow->Insert(xlShiftDown);

		// Định dạng khung
		PRS = psNDNK->GetRange(prNDNK->GetItem(num1+cong,cot),prNDNK->GetItem(num1+cong+sl-1,cot+4));

		PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->PutWeight(xlThin);
		PRS->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->PutWeight(xlThin);
		PRS->Borders->GetItem(XlBordersIndex::xlEdgeLeft)->PutWeight(xlThin);
		PRS->Borders->GetItem(XlBordersIndex::xlEdgeRight)->PutWeight(xlThin);
		PRS->Borders->GetItem(XlBordersIndex::xlInsideVertical)->PutWeight(xlThin);

		PRS->Borders->GetItem(XlBordersIndex::xlInsideHorizontal)->PutWeight(xlThin);
		PRS->Borders->GetItem(XlBordersIndex::xlInsideHorizontal)->PutLineStyle(xlDot);

		PRS->EntireRow->Rows->AutoFit();

		// Load dữ liệu
		sl = num1+cong-1;
		if(dem>0)
		for (int i = 1; i <=dem; i++)
		{
			szval.Format(L"%d",i);
			prNDNK->PutItem(sl+i,cot,(_bstr_t)szval);			
			prNDNK->PutItem(sl+i,cot+1,(_bstr_t)sLoad[i]);
			prNDNK->PutItem(sl+i,cot+2,(_bstr_t)sfr1[i]);
			prNDNK->PutItem(sl+i,cot+3,(_bstr_t)sfr2[i]);
			prNDNK->PutItem(sl+i,cot+4,(_bstr_t)sfr3[i]);
		}

		if(ichk0==1)
		{
			PRS = psNDNK->GetRange(prNDNK->GetItem(num1,1),prNDNK->GetItem(num1+cong,1));
			PRS->EntireRow->PutHidden(true);	
		}
	}
	catch(...){}
}


// Hàm load nội dung nhật ký
// sNgayNK = ngày lưu nhật ký
// num = dòng bắt đầu đổ dữ liệu; gt = số dòng được cộng thêm để bắt đầu tính vị trí num
// frTable = tên bảng truy vấn; type = kiểu truy vấn NCM
void cls_main::f_load_ndung(CString sNgayNK,int num1,int num2,int cot,int cong,CString frTable,int type,int type2,int iRend)
{
	try
	{
		// Kiểm tra dòng tiêu đề có bị ẩn không? --> Nếu ẩn thì return luôn
		PRS = prNDNK->GetItem(num1,1);
		int gtri = PRS->GetRowHeight();
		if(gtri==0) return;

		int chen=4;
		PRS = prNDNK->GetItem(num1,cot);
		if(PRS->GetComment()!=NULL)
		{
			_bstr_t bscmt = PRS->GetComment()->Text();
			CString sadd = (LPCTSTR)(bscmt);
			sadd.Trim();
			chen = _wtoi(sadd);
			if(chen<=-1) chen=-1;			
		}

		if(chen==-1)
		{
			PRS = psNDNK->GetRange(prNDNK->GetItem(num1,1),prNDNK->GetItem(num2-1,1));
			PRS->EntireRow->PutHidden(true);
			return;
		}
		else if(chen==0)
		{
			PRS = psNDNK->GetRange(prNDNK->GetItem(num1+1,1),prNDNK->GetItem(num2-1,1));
			PRS->EntireRow->PutHidden(true);
		}

		// Xóa dữ liệu cũ (nếu có)
		if(num2>num1+cong)
		{
			PRS = psNDNK->GetRange(prNDNK->GetItem(num1+cong,cot),prNDNK->GetItem(num1+cong,cot+4));
			PRS->ClearContents();

			if(num2>num1+cong+1)
			{
				PRS = psNDNK->GetRange(prNDNK->GetItem(num1+cong+1,1),prNDNK->GetItem(num2-1,1));
				PRS->EntireRow->Delete(xlShiftUp);	
			}
		}

		CString szval=L"";
		CString sNoidung=L"",sCheck=L"";
		SqlString.Format(L"SELECT * FROM %s WHERE ngayghi LIKE '%s';",frTable,sNgayNK);
		CRecordset* Recset = ObjConn.Execute(SqlString);
		while( !Recset->IsEOF() )
		{
			Recset->GetFieldValue(L"noidung",sNoidung);
			Recset->GetFieldValue(L"checknew",sCheck);
			break;
		}
		Recset->Close();

		// Xác định số lượng dòng cần chèn thêm
		int vtri=0,dem=0;
		CString sLoad[200];

		if(frTable==L"tbNcong" && sNoidung==L"")
			sNoidung = L"- Kỹ sư: ;- Cao đẳng, trung cấp: ;- Công nhân: ;- Nhân lực khác: ";

		sNoidung = ReplateChamphay(sNoidung);
		if(sNoidung!=L"")
		{
			if(sNoidung.Right(1)!=L";") sNoidung+=L";";
			for (int i = 0; i < sNoidung.GetLength(); i++)
			{
				szval = sNoidung.Mid(i,1);
				if(szval==L";" && i>=vtri)
				{
					dem++;
					if(sNoidung.Mid(i-1,1)==L";") sLoad[dem]=L"";
					else sLoad[dem] = sNoidung.Mid(vtri,i-vtri);
					vtri=i+1;
				}
			}
		}

		// Xác định số dòng sẽ chèn thêm
		if(dem>0) chen=dem;

		if(chen==1)
		{
			PRS = psNDNK->GetRange(prNDNK->GetItem(num1+cong,cot),prNDNK->GetItem(num1+cong,cot+4));
			PRS->EntireRow->PutHidden(false);
		}
		else if(chen>1)
		{
			PRS = psNDNK->GetRange(prNDNK->GetItem(num1+cong+1,1),prNDNK->GetItem(num1+cong+chen-1,1));
			PRS->EntireRow->Insert(xlShiftDown);
		}

///////// Load dữ liệu
		if(dem>0)
		for (int i = 1; i <=dem; i++)
		{
			if(sLoad[i].Left(1)==L"-") sLoad[i] = L"'" + sLoad[i];
			prNDNK->PutItem(num1+cong+i-1,cot,(_bstr_t)sLoad[i]);
			if(sLoad[i]!=L"") f_AutoMerge(psNDNK,num1+cong+i-1,cot,cot+4);
		}

		int gtcong=0;
		_variant_t sbox[3]={"","",""};
		if(frTable==L"tbBptc")
		{
			gtcong=12;
			sbox[0] = "Check Box 44";
			sbox[1] = "Check Box 45";
			sbox[2] = "Check Box 46";
		}
		else if(frTable==L"tbTTTTe")
		{
			gtcong=13;
			sbox[0] = "Check Box 47";
			sbox[1] = "Check Box 48";
			sbox[2] = "Check Box 49";
		}
		else if(frTable==L"tbCltc")
		{
			gtcong=14;
			sbox[0] = "Check Box 50";
			sbox[1] = "Check Box 51";
			sbox[2] = "Check Box 52";
		}
		else if(frTable==L"tbTdth")
		{
			gtcong=15;
			sbox[0] = "Check Box 53";
			sbox[1] = "Check Box 54";
			sbox[2] = "Check Box 55";
		}

		if(gtcong>0)
		{
			CString schk[3]={L"Tốt",L"Bình thường",L"Kém"};			
			schk[0] = GIOR(1,cot+gtcong,prNDNK,L"Tốt");
			schk[1] = GIOR(2,cot+gtcong,prNDNK,L"Bình thường");
			schk[2] = GIOR(3,cot+gtcong,prNDNK,L"Kém");

			if(sCheck==schk[0]) f_chekbox1(sbox[0],1,sbox[1],0,sbox[2],0);
			else if(sCheck==schk[1]) f_chekbox1(sbox[0],0,sbox[1],1,sbox[2],0);
			else if(sCheck==schk[2]) f_chekbox1(sbox[0],0,sbox[1],0,sbox[2],1);
			else f_chekbox1(sbox[0],0,sbox[1],0,sbox[2],0);
		}

		// Xóa định dạng
		if(chen>=1)
		{
			PRS = psNDNK->GetRange(prNDNK->GetItem(num1+cong,cot),prNDNK->GetItem(num1+cong+chen-1,iRend));
			f_dinh_dang_khung(PRS,chen-1,type2);
		}		
	}
	catch(...){}
}


CString cls_main::ReplateChamphay(CString sNoidung)
{
	try
	{
		int pos=0;
		while (pos>=0)
		{
			pos = sNoidung.Find(L";;");
			if(pos>=0) sNoidung.Replace(L";;",L"");
		}

		if(sNoidung==L";") sNoidung=L"";
		return sNoidung;
	}
	catch(...){return L"";}
	return L"";
}


void cls_main::f_AutoMerge(_WorksheetPtr pSh1,int iRw,int iCl1,int iCl2)
{
	try
	{
		// Lấy ngẫu nhiên cột để đổ dữ liệu
		int iClRand=67;

		// UnMerge (nếu có)
		RangePtr pRg1 = pSh1->Cells;
		RangePtr PRG = pRg1->GetItem(iRw,iCl1);
		PRG->UnMerge();

		// Lấy độ rộng ban đầu của cột (iClRand)
		PRG = pRg1->GetItem(1,iClRand);
		double ntrc = PRG->GetColumnWidth(); 

		// Copy & paste dữ liệu vào cột (iClRand) 
		PRG = pRg1->GetItem(iRw,iCl1);
		PRG->Copy(pRg1->GetItem(iRw,iClRand));
		xl->PutCutCopyMode(XlCutCopyMode(false));

		// Xác định tổng độ rộng các cột được lấy (tính cả đường line)
		double tong=0,nval=0;
		for (int i = iCl1; i <= iCl2; i++)
		{
			PRG = pRg1->GetItem(1,i);
			nval = PRG->GetColumnWidth();
			tong+=nval;	// Độ rộng ô thứ [i]
			tong+=0.71;	// Độ rộng vạch kẻ ngang (0.71 = con số tương đối của đường line)
		}

		// Merge & Gán giá trị vào cột gốc
		PRG = pSh1->GetRange(pRg1->GetItem(iRw,iCl1),pRg1->GetItem(iRw,iCl2));
		PRG->PutMergeCells(1);
		PRG->PutWrapText(1);
		PRG->PutHorizontalAlignment(xlJustify);

		// Set width & tự động autofit
		PRG = pRg1->GetItem(iRw,iClRand);
		PRG->PutColumnWidth(tong);
		PRG->PutWrapText(1);
		PRG->PutHorizontalAlignment(xlJustify);
		PRG->EntireRow->Rows->AutoFit();

		// Lấy giá trị độ cao vị trí (iClRand) --> Xóa dữ liệu cột iClRand --> Gán giá trị cột đã lấy được (height)
		double gtri = PRG->GetRowHeight();		
		if(PRG->GetComment()!=NULL) PRG->ClearComments();
		PRG->ClearContents();
		PRG->ClearFormats();

		PRG->PutColumnWidth(ntrc);
		PRG->PutRowHeight(gtri);
	}
	catch(...)
	{
		CString szval=L"";
		szval.Format(L"Lỗi auto merge tại dòng %d, từ cột số %d -> %d. "
			L"\nBạn có thể nhập từ khóa trên google: 'auto merge + gxd' để xử lý lỗi.",iRw,iCl1,iCl2);
		_s(szval);
	}
}


void cls_main::f_chekbox1(CString v1,int n1,CString v2,int n2,CString v3,int n3)
{
	psNDNK->GetShapes()->Item((_variant_t)v1)->ControlFormat->PutValue(n1);
	psNDNK->GetShapes()->Item((_variant_t)v2)->ControlFormat->PutValue(n2);
	psNDNK->GetShapes()->Item((_variant_t)v3)->ControlFormat->PutValue(n3);
}


// Nghiệm thu hạng mục (sửa 07.08.2017)
void cls_main::f_XNKY_5(int iRend)
{
	try
	{
		CString szval=L"";

		// Xác định END
		int xde = FindComment(psDMGD,_colgd[1],"END");
		CString f1,f4,f5,f11;
		int dem=0;

		for (int i = 8; i < xde; i++)
		{
			f1 = prDMGD->GetItem(i,_colgd[1]);
			f4 = prDMGD->GetItem(i,_colgd[4]);
		
			// Thỏa mãn các điều kiện để xuất nhật ký
			if(_wtoi(f1)>0 && f4!=L"")
			{
				f11 = prDMGD->GetItem(i,_colgd[11]);
				if(f11==_ngaythang)
				{
					dem++;
					f5 = prDMGD->GetItem(i,_colgd[5]);

					//szval = prDMGD->GetItem(i,_colgd[15]);
					//if(szval==L"") szval=L"Nghiệm thu";
					szval=L"Nghiệm thu";

					if(f5!=L"") _xnk[dem].Format(L"'-%s %s",szval,f5);
					else _xnk[dem].Format(L"'-%s ",szval);
				}
			}
		}

		// Kiểm tra có comment không?
		int sl=0;
		PRS = prNDNK->GetItem(_ivtr[_numNK],_ivtr[0]);		
		if(PRS->GetComment()==NULL) PRS = prNDNK->GetItem(_ivtr[_numNK],1);
		if(PRS->GetComment()!=NULL)
		{
			_bstr_t bscmt = PRS->GetComment()->Text();
			CString sadd = (LPCTSTR)(bscmt);
			sadd.Trim();
			sl = _wtoi(sadd);
			if(sl<0) sl=0;
			else if(sl>10) sl=10;
		}

		if(dem>0 || sl>0)
		{
			// Xuất dữ liệu
			PRS = psNDNK->GetRange(prNDNK->GetItem(_ivtr[_numNK+1],1),prNDNK->GetItem(_ivtr[_numNK+1]+dem+sl-1,1));
			PRS->EntireRow->Insert(xlShiftDown);
		
			if(dem>0)
			for (int i = 1; i <=dem; i++)
			{
				prNDNK->PutItem(_ivtr[_numNK]+i,_ivtr[0],(_bstr_t)_xnk[i]);
				f_AutoMerge(psNDNK,_ivtr[_numNK]+i,_ivtr[0],_ivtr[0]+3);
			}

			PRS = psNDNK->GetRange(prNDNK->GetItem(_ivtr[_numNK+1],1),prNDNK->GetItem(_ivtr[_numNK+1]+dem+sl-1,iRend));
			f_dinh_dang_khung(PRS,1,1);
		}
	}
	catch(_com_error & error){}
}


// Nghiệm thu vật liệu (sửa 07.08.2017)
void cls_main::f_XNKY_4(int iRend)
{
	try
	{
		CString szval=L"";

		// Xác định END
		int xde = FindComment(psDMVL,_colvl[1],"END");
		CString f0,f1,f3,f4,f5,f15;
		int dem=0;

		// Lấy mẫu sheet vật liệu
		for (int i = 8; i < xde; i++)
		{
			f1 = prDMVL->GetItem(i,_colvl[1]);
			f4 = prDMVL->GetItem(i,_colvl[4]);
		
			// Thỏa mãn các điều kiện để xuất nhật ký
			if(_wtoi(f1)>0 && f4!=L"")
			{
				f15 = prDMVL->GetItem(i,_colvl[15]);
				if(f15==_ngaythang)
				{
					dem++;
					f5 = prDMVL->GetItem(i,_colvl[5]);

					//szval = prDMVL->GetItem(i,_colvl[20]);
					//if(szval==L"") szval = L"Nghiệm thu vật liệu";
					szval = L"Nghiệm thu vật liệu";

					if(f5!=L"") _xnk[dem].Format(L"'-%s %s",szval,f5);
					else _xnk[dem].Format(L"'-%s ",szval);
				}
			}
		}

		// Kiểm tra có comment không?
		int sl=0;
		PRS = prNDNK->GetItem(_ivtr[_numNK],_ivtr[0]);		
		if(PRS->GetComment()==NULL) PRS = prNDNK->GetItem(_ivtr[_numNK],1);
		if(PRS->GetComment()!=NULL)
		{
			_bstr_t bscmt = PRS->GetComment()->Text();
			CString sadd = (LPCTSTR)(bscmt);
			sadd.Trim();
			sl = _wtoi(sadd);
			if(sl<0) sl=0;
			else if(sl>10) sl=10;
		}

		if(dem>0 || sl>0)
		{
			// Xuất dữ liệu
			PRS = psNDNK->GetRange(prNDNK->GetItem(_ivtr[_numNK+1],1),prNDNK->GetItem(_ivtr[_numNK+1]+dem+sl-1,1));
			PRS->EntireRow->Insert(xlShiftDown);
		
			if(dem>0)
			for (int i = 1; i <=dem; i++)
			{
				prNDNK->PutItem(_ivtr[_numNK]+i,_ivtr[0],(_bstr_t)_xnk[i]);
				f_AutoMerge(psNDNK,_ivtr[_numNK]+i,_ivtr[0],_ivtr[0]+3);
			}

			PRS = psNDNK->GetRange(prNDNK->GetItem(_ivtr[_numNK+1],1),prNDNK->GetItem(_ivtr[_numNK+1]+dem+sl-1,iRend));
			f_dinh_dang_khung(PRS,1,1);
		}
	}
	catch(_com_error & error){}
}


// Nghiệm thu công việc (sửa 07.08.2017)
void cls_main::f_XNKY_3(int iRend)
{
	try
	{
		CString szval=L"";

		// Xác định END
		int xde = FindComment(psDMCV,_colcv[1],"END");
		CString f0,f1,f3,f4,f5,f6,f15;
		int dem=0;

		for (int i = 8; i < xde; i++)
		{
			f1 = prDMCV->GetItem(i,_colcv[1]);
			f4 = prDMCV->GetItem(i,_colcv[4]);
		

			// Thỏa mãn các điều kiện để xuất nhật ký
			if(_wtoi(f1)>0 && f4!=L"")
			{
				f15 = prDMCV->GetItem(i,_colcv[15]);
				if(f15==_ngaythang)
				{
					dem++;
					f5 = prDMCV->GetItem(i,_colcv[5]);
					f6 = prDMCV->GetItem(i,_colcv[6]);

					//szval = prDMCV->GetItem(i,_colcv[26]);
					//if(szval==L"") szval = L"Nghiệm thu công việc";
					szval = L"Nghiệm thu công việc";

					if(f5!=L"" && f6!=L"") _xnk[dem].Format(L"'-%s %s / %s",szval,f5,f6);
					else if(f5!=L"" && f6==L"") _xnk[dem].Format(L"'-%s %s",szval,f5);
					else if(f5==L"" && f6!=L"") _xnk[dem].Format(L"'-%s / %s",szval,f6);
					else if(f5==L"" && f6==L"") _xnk[dem].Format(L"'-%s ",szval);
				}
			}
		}

		// Kiểm tra có comment không?
		int sl=0;
		PRS = prNDNK->GetItem(_ivtr[_numNK],_ivtr[0]);		
		if(PRS->GetComment()==NULL) PRS = prNDNK->GetItem(_ivtr[_numNK],1);
		if(PRS->GetComment()!=NULL)
		{
			_bstr_t bscmt = PRS->GetComment()->Text();
			CString sadd = (LPCTSTR)(bscmt);
			sadd.Trim();
			sl = _wtoi(sadd);
			if(sl<0) sl=0;
			else if(sl>10) sl=10;
		}

		if(dem>0 || sl>0)
		{
			// Xuất dữ liệu
			PRS = psNDNK->GetRange(prNDNK->GetItem(_ivtr[_numNK+1],1),prNDNK->GetItem(_ivtr[_numNK+1]+dem+sl-1,1));
			PRS->EntireRow->Insert(xlShiftDown);

			if(dem>0)
			for (int i = 1; i <=dem; i++)
			{
				prNDNK->PutItem(_ivtr[_numNK]+i,_ivtr[0],(_bstr_t)_xnk[i]);
				f_AutoMerge(psNDNK,_ivtr[_numNK]+i,_ivtr[0],_ivtr[0]+3);
			}

			PRS = psNDNK->GetRange(prNDNK->GetItem(_ivtr[_numNK+1],1),prNDNK->GetItem(_ivtr[_numNK+1]+dem+sl-1,iRend));
			f_dinh_dang_khung(PRS,1,1);
		}
	}
	catch(_com_error & error){}
}


// Nghiệm thu thi công (sửa 07.08.2017)
void cls_main::f_XNKY_2(int iRend)
{
	try
	{
		CString szval=L"";

		// Xác định END
		int xde = FindComment(psDMCV,_colcv[1],"END");
		CString f0,f1,f2,f3,f4,f5,ngay1,ngay2,gan1,gan2;
		CString _ktra=L"",stenHM=L"";		 
		int lui=0,dem=0,iktrHM=0;
		BOOLEAN bl=TRUE;

		for (int i = 8; i <= xde; i++)
		{
			f1 = prDMCV->GetItem(i,_colcv[1]);
			f2 = prDMCV->GetItem(i,_colcv[2]);
			f3 = prDMCV->GetItem(i,_colcv[3]);
			f3 = prDMCV->GetItem(i,_colcv[3]);
			f4 = prDMCV->GetItem(i,_colcv[4]);
			f5 = prDMCV->GetItem(i,_colcv[5]);
			
			//szval = prDMCV->GetItem(i,_colcv[25]);
			//if(szval==L"") szval = L"Thi công";
			szval = L"Thi công";

			if(f2==L"HM")
			{
				iktrHM=1;
				stenHM = f5;
				continue;
			}

			// Thỏa mãn các điều kiện để xuất nhật ký
			if((_wtoi(f1)>0 && f4!=L"") || i==xde)
			{
				// Kiểm tra xem có tồn tại TC ở công việc trước không?
				if(bl==FALSE)
				{
					lui=i-1;
					_ktra = prDMCV->GetItem(lui,_colcv[1]);
					while (_wtoi(_ktra)==0 && lui>8)
					{
						lui--;
						_ktra = prDMCV->GetItem(lui,_colcv[1]);
					}

					ngay1 = prDMCV->GetItem(lui,_colcv[18]);
					ngay2 = prDMCV->GetItem(lui,_colcv[19]);
					if( compare_date(numNgay,ngay1,_ngaythang)==1 && compare_date(numNgay,_ngaythang,ngay2)==1)
					{
						if(iktrHM==1)
						{
							dem++;
							_xnk[dem] = L"HM: " + stenHM;
							iktrHM=0;
						}

						dem++;
						if(gan2!=L"") _xnk[dem].Format(L"%s / %s",gan1,gan2);
						else _xnk[dem] = gan1;
					}
				}

				if(f5!=L"") gan1.Format(L"'-%s %s ",szval,f5);
				else gan1.Format(L"'-%s ",szval);
				gan2 = prDMCV->GetItem(i,_colcv[6]);
				bl=FALSE;
			}

			if(_wtoi(f1)==0 && f3.Left(2).MakeUpper()==L"TC" && f5!=L"")
			{
				bl=TRUE;
				ngay1 = prDMCV->GetItem(i,_colcv[18]);
				ngay2 = prDMCV->GetItem(i,_colcv[19]);
				if( compare_date(numNgay,ngay1,_ngaythang)==1 && compare_date(numNgay,_ngaythang,ngay2)==1)
				{
					dem++;					
					if(f5!=L"" && gan2!=L"") _xnk[dem].Format(L"'-%s %s / %s",szval,f5,gan2);
					else if(f5!=L"" && gan2==L"") _xnk[dem].Format(L"'-%s %s",szval,f5);
					else if(f5==L"" && gan2!=L"") _xnk[dem].Format(L"'-%s / %s",szval,gan2);
					else if(f5==L"" && gan2==L"") _xnk[dem].Format(L"'-%s ",szval);
				}
			}
		}

		// Kiểm tra có comment không?
		int sl=0;
		PRS = prNDNK->GetItem(_ivtr[_numNK],_ivtr[0]);		
		if(PRS->GetComment()==NULL) PRS = prNDNK->GetItem(_ivtr[_numNK],1);
		if(PRS->GetComment()!=NULL)
		{
			_bstr_t bscmt = PRS->GetComment()->Text();
			CString sadd = (LPCTSTR)(bscmt);
			sadd.Trim();
			sl = _wtoi(sadd);
			if(sl<0) sl=0;
			else if(sl>10) sl=10;
		}

		if(dem>0 || sl>0)
		{
			// Xuất dữ liệu
			PRS = psNDNK->GetRange(prNDNK->GetItem(_ivtr[_numNK+1],1),prNDNK->GetItem(_ivtr[_numNK+1]+dem+sl-1,1));
			PRS->EntireRow->Insert(xlShiftDown);
		
			if(dem>0)
			for (int i = 1; i <=dem; i++)
			{
				prNDNK->PutItem(_ivtr[_numNK]+i,_ivtr[0],(_bstr_t)_xnk[i]);
				f_AutoMerge(psNDNK,_ivtr[_numNK]+i,_ivtr[0],_ivtr[0]+3);
			}

			// Định dạng khung
			PRS = psNDNK->GetRange(prNDNK->GetItem(_ivtr[_numNK+1],1),prNDNK->GetItem(_ivtr[_numNK+1]+dem+sl-1,iRend));
			f_dinh_dang_khung(PRS,1,1);

			// Bôi đậm hạng mục
			for (int i = _ivtr[_numNK+1]; i < _ivtr[_numNK+1]+dem+sl; i++)
			{
				szval = GIOR(i,_ivtr[0],prNDNK,L"");
				if(szval.Left(2)==L"HM")
				{
					PRS = prNDNK->GetItem(i,_ivtr[0]);
					PRS->Font->PutBold(1);
				}
			}
		}
	}
	catch(_com_error & error){}
}


// Lấy mẫu thí nghiệm (sửa 07.08.2017)
void cls_main::f_XNKY_1(int iRend)
{
	try
	{
		CString szval=L"";

		// Xác định END
		int xde = FindComment(psDMVL,_colvl[1],"END");
		CString f0,f1,f3,f4,f5,f9,gan1,gan2;
		int dem=0;

		// Lấy mẫu sheet vật liệu
		for (int i = 8; i < xde; i++)
		{
			f1 = prDMVL->GetItem(i,_colvl[1]);
			f4 = prDMVL->GetItem(i,_colvl[4]);

			// Thỏa mãn các điều kiện để xuất nhật ký
			if(_wtoi(f1)>0 && f4!=L"")
			{
				f9 = prDMVL->GetItem(i,_colvl[9]);
				if(f9==_ngaythang)
				{
					dem++;
					f5 = prDMVL->GetItem(i,_colvl[5]);

					//szval = prDMVL->GetItem(i,_colvl[19]);
					//if(szval==L"") szval=L"Lấy mẫu";
					szval=L"Lấy mẫu";

					_xnk[dem].Format(L"'-%s %s.",szval,f5);
				}
			}
		}

		// 14.06.2017
		// Nếu không có ngày LM thì mặc định lấy ở trên

		// Lấy mẫu sheet công việc
		BOOLEAN bl = FALSE;
		xde = FindComment(psDMCV,_colcv[1],"END");
		for (int i = 8; i < xde; i++)
		{
			f1 = prDMCV->GetItem(i,_colcv[1]);
			f3 = prDMCV->GetItem(i,_colcv[3]);
			f4 = prDMCV->GetItem(i,_colcv[4]);
			f5 = prDMCV->GetItem(i,_colcv[5]);

			// Thỏa mãn các điều kiện để xuất nhật ký
			if(_wtoi(f1)>0 && f4!=L"")
			{
				//szval = prDMCV->GetItem(i,_colcv[24]);
				//if(szval==L"") szval=L"Lấy mẫu";
				szval=L"Lấy mẫu";

				if(f5!=L"") gan1.Format(L"'-%s %s ",szval,f5);
				else gan1.Format(L"'-%s ",szval);
				gan2 = prDMCV->GetItem(i,_colcv[6]);
			}

			if(_wtoi(f1)==0 && f3.Left(2).MakeUpper()==keyLMTN && f5!=L"")
			{
				f9 = prDMCV->GetItem(i,_colcv[9]);
				if(f9==L"") f9 = prDMCV->GetItem(i-1,_colcv[9]);	// bổ sung dòng này 14.09.2017
				
				if(f9==_ngaythang)
				{
					dem++;
					if(gan2!=L"") _xnk[dem].Format(L"%s / %s",gan1,gan2);
					else _xnk[dem] = gan1;
				}
			}
		}

		// Kiểm tra có comment không?
		int sl=0;
		PRS = prNDNK->GetItem(_ivtr[_numNK],_ivtr[0]);		
		if(PRS->GetComment()==NULL) PRS = prNDNK->GetItem(_ivtr[_numNK],1);
		if(PRS->GetComment()!=NULL)
		{
			_bstr_t bscmt = PRS->GetComment()->Text();
			CString sadd = (LPCTSTR)(bscmt);
			sadd.Trim();
			sl = _wtoi(sadd);
			if(sl<0) sl=0;
			else if(sl>10) sl=10;
		}

		if(dem>0 || sl>0)
		{
			// Xuất dữ liệu
			PRS = psNDNK->GetRange(prNDNK->GetItem(_ivtr[_numNK+1],1),prNDNK->GetItem(_ivtr[_numNK+1]+dem+sl-1,1));
			PRS->EntireRow->Insert(xlShiftDown);
		
			if(dem>0)
			for (int i = 1; i <=dem; i++)
			{
				prNDNK->PutItem(_ivtr[_numNK]+i,_ivtr[0],(_bstr_t)_xnk[i]);
				f_AutoMerge(psNDNK,_ivtr[_numNK]+i,_ivtr[0],_ivtr[0]+3);
			}

			// Định dạng khung
			PRS = psNDNK->GetRange(prNDNK->GetItem(_ivtr[_numNK+1],1),prNDNK->GetItem(_ivtr[_numNK+1]+dem+sl-1,iRend));
			f_dinh_dang_khung(PRS,1,1);
		}
	}
	catch(_com_error & error){}
}


void cls_main::f_dinh_dang_khung(RangePtr pRg,int itype,int itype2)
{
	pRg->Interior->PutColor(xlNone);
	pRg->Font->PutColor(BLACK);
	pRg->Font->PutBold(0);
	pRg->Font->PutItalic(0);

	if (itype > 0)
	{
		pRg->Borders->GetItem(XlBordersIndex::xlInsideHorizontal)->PutWeight(xlThin);
		pRg->Borders->GetItem(XlBordersIndex::xlInsideHorizontal)->PutLineStyle(xlDot);
	}

	if (itype2 > 0)
	{
		pRg->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->PutWeight(xlThin);
		pRg->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->PutLineStyle(xlDot);
	}
}


void cls_main::f_importspic()
{
	try
	{
		CoInitialize(NULL);
		xl.GetActiveObject(L"Excel.Application");
		xl->Visible = true;
		xl->PutScreenUpdating(VARIANT_FALSE);

		// Hiển thị hình ảnh
		pWb = xl->GetActiveWorkbook();
		pSheet = pWb->ActiveSheet;
		_code = pSheet->CodeName;
		pRange = pSheet->Cells;

		shTS = name_sheet("shTS");
		psTS = xl->Sheets->GetItem(shTS);
		prTS = psTS->GetCells();

		int iswRow = getRow(psTS,"CF_SWIMG");
		int iswCol = getColumn(psTS,"CF_SWIMG");

		BOOLEAN f = TRUE;
		int irow=0,icol=0;

		// Xác định dòng bắt đầu đổ ảnh
		for (int i = 1; i < 200; i++)
		{
			msg = GIOR(i,1,pRange,L"");
			if(msg==L"pic_begin")
			{
				irow=i+1;
				break;	
			}
		}

		_bstr_t _bienban;
		if(_code== (_bstr_t)L"shHSNTCVNBCV")
		{
			// Nội bộ công việc
			_RSpin = getRow(pSheet,"NTNBCVGD_BB");
			_CSpin = getColumn(pSheet,"NTNBCVGD_BB");

			// OK
			if(irow==0) irow = getRow(pSheet,"NTNB_VT")+1;
			//icol = getColumn(pSheet,"NTNB_DT");
			_bienban = (_bstr_t)L"shHSNTCV";

			iswRow+=2;
		}
		else if(_code== (_bstr_t)L"shHSNTCVYNCV")
		{
			// Yêu cầu công việc
			_RSpin = getRow(pSheet,"YCNTCV_BB");
			_CSpin = getColumn(pSheet,"YCNTCV_BB");

			// OK
			if(irow==0) irow = getRow(pSheet,"YCNTCV_TG")-1;
			//icol = getColumn(pSheet,"YCNTCV_DT");
			_bienban = (_bstr_t)L"shHSNTCV";

			iswRow+=3;
		}
		else if(_code== (_bstr_t)L"shHSNTCVNTCV")
		{
			// Nghiệm thu công việc
			_RSpin = getRow(pSheet,"NTCV_BB");
			_CSpin = getColumn(pSheet,"NTCV_BB");

			// OK
			if(irow==0) irow = getRow(pSheet,"NTCV_VT")+1;
			//icol = getColumn(pSheet,"NTCV_DT");
			_bienban = (_bstr_t)L"shHSNTCV";

			iswRow+=4;
		}
		else if(_code== (_bstr_t)L"shHSNTGDNTBGD1")
		{
			// Nội bộ giai đoạn
			_RSpin = getRow(pSheet,"NTNBGDG_BB");
			_CSpin = getColumn(pSheet,"NTNBGDG_BB");

			// OK
			if(irow==0) irow = getRow(pSheet,"NTNBGD_TPNT");
			//icol = getColumn(pSheet,"NTNBGD_DT");
			_bienban = (_bstr_t)L"shHSNTGD";

			iswRow+=7;
		}
		else if(_code== (_bstr_t)L"shHSNTGDYCNT1")
		{
			// Yêu cầu giai đoạn
			_RSpin = getRow(pSheet,"YCGD_BB");
			_CSpin = getColumn(pSheet,"YCGD_BB");

			// OK
			if(irow==0) irow = getRow(pSheet,"YCGD_TG")-1;
			//icol = getColumn(pSheet,"YCGD_DT");
			_bienban = (_bstr_t)L"shHSNTGD";

			iswRow+=8;
		}
		else if(_code== (_bstr_t)L"shHSNTGDNTGD1")
		{
			// Nghiệm thu giai đoạn
			_RSpin = getRow(pSheet,"NTGD_BB");
			_CSpin = getColumn(pSheet,"NTGD_BB");

			// OK
			if(irow==0) irow = getRow(pSheet,"NTGD_DTP");
			//icol = getColumn(pSheet,"NTGD_DT");
			_bienban = (_bstr_t)L"shHSNTGD";

			iswRow+=9;
		}
		else f = FALSE;

		if(f==TRUE)
		{
			// Xóa ảnh
			try
			{
				int _gcout = pSheet->Shapes->Count;

				_variant_t p0;
				CString _pic0;
				for (int i = 1; i <= _gcout; i++)
				{
					p0 = pSheet->GetShapes()->Item(i)->Name;
					_pic0 = p0.bstrVal;

					if(_pic0==L"GXDPIC") pSheet->GetShapes()->Item(i)->Delete();
				}
			}
			catch(_com_error & error){}

			pRange->PutItem(irow-1,1,(_bstr_t)L"pic_begin");			
			for (int i = irow; i < irow+100; i++)
			{
				msg = GIOR(i,1,pRange,L"");
				if(msg==L"pic_end")
				{
					if(i>irow)
						pSheet->GetRange(pRange->GetItem(irow,1),pRange->GetItem(i-1,1))->EntireRow->Delete(xlShiftUp);
					break;					
				}
			}
			pRange->PutItem(irow,1,(_bstr_t)L"pic_end");

			_ihidepic = _wtoi(getVCell(psTS,L"TS_PIC"));
			_inextpic = _wtoi(getVCell(psTS,L"TS_NEXTPIC"));
			int _ishw = _wtoi(GIOR(iswRow,iswCol,prTS,L""));

			if(_ihidepic==1 && _ishw==1)
			{
				int inumcv = pRange->GetItem(_RSpin,_CSpin);
				CString _pathpic = f_load_congviec(inumcv,_bienban);	// path image
				if(_pathpic!=L"")
				{
					// Số dòng cần chèn
					int _pnguyen = 0;
					int _sodu = 0;					

					// Độ rộng tổng cột (C-->AC) tính theo pixcel và point
					_iwidth=0,_ipoints=0;
					for (int i = 3; i < _CSpin-1; i++)
					{
						_iwidth+=_formatPixcel(pRange,i);
						
						PRS = pRange->GetItem(1,i);
						_ipoints += (int)PRS->GetWidth();
					}

					int x1 = _ipoints;	// Chiều dài ảnh sẽ load (luôn fix = độ rộng thực cột B) tính theo point
					int y1 = 0;	// Độ cao ảnh sẽ load tính theo point
			
					/*
					Kích thước ảnh được co tự động trong code
					Kích thước _iheight = f(_iwidth)
					Xét kích thước thực của ảnh --> Lấy ra 2 giá trị iWThuc & iHThuc
					_iwidth được xác định thông qua formatpixcel
					Xác định chiều cao ảnh sau khi điều chỉnh (point) = (pixcel)*(pixcel)/(pixcel)
					*/

					// Xét kích thước thực của ảnh --> Tạo tỷ lệ lấy ra được giá trị y1
					_fGetImage(_pathpic);
					if(iHThuc>0 && iWThuc>0)
					{
						y1 = (int)(3*_iwidth*iHThuc/(4*iWThuc));

						// Kiểm tra chiều cao có vượt quá giới hạn không (0-409)
						int k = 3*_imulti;
						_sodu = y1%k;
						_pnguyen = (int)(y1/k);
					}
					else y1=0;

					if(_pathpic.Right(4)==L".jpg" && y1>0)
					{
						try
						{
							// Chèn dòng
							int gt = irow+_inextpic;
							pSheet->GetRange(pRange->GetItem(irow,1),pRange->GetItem(gt+_pnguyen+1,1))->EntireRow->Insert(xlShiftDown);

							try
							{
								// Chèn ảnh vào biên bản & điều chỉnh kích thước theo (x1,y1) = (rộng,cao) đơn vị: pixcel
								PRS = pRange->GetItem(gt,3);
								ShapePtr ShapeP = pSheet->Shapes->AddPicture(
									(_bstr_t)_pathpic,Office::msoTrue,Office::msoTrue,
									PRS->Left,PRS->Top,(float)x1,(float)y1);
								ShapeP->PutName((_bstr_t)"GXDPIC");
								ShapeP->IncrementTop(6);
								ShapeP->IncrementLeft(6);
							}
							catch(_com_error & error){_s(L"Không tìm thấy tệp ảnh tương ứng.");}							

							// Tăng kích thước cho dòng chứa ảnh (fix cứng = 3*_imulti)
							if(_pnguyen>0)
							{
								// Trường hợp _pnguyen > 0 : y1 >= 3*_imulti
								// Chia tiếp làm 2 trường hợp nhỏ
								int k = 3*_imulti;

								// _sodu = 0 : y1 chia hết không cho (3*_imulti);
								if(_sodu>0)
								{
									// Phần fix cứng
									PRS = pSheet->GetRange(pRange->GetItem(gt+1,1),pRange->GetItem(gt+_pnguyen,1));
									PRS->PutRowHeight(k);

									// Phần resize theo số dư
									PRS = pRange->GetItem(gt+_pnguyen+1,1);
									PRS->PutRowHeight(_sodu);

									// Số dòng cộng thêm = _pnguyen+1
									gt+=(_pnguyen+1);
								}
								else
								{
									// Trường hợp chia hết, số dòng cộng thêm = _pnguyen
									PRS = pSheet->GetRange(pRange->GetItem(gt+1,1),pRange->GetItem(gt+_pnguyen,1));
									PRS->PutRowHeight(k);
									gt+=_pnguyen;
								}
							}
							else
							{
								// Trường hợp _pnguyen = 0 : y1 < 3*_imulti
								PRS = pRange->GetItem(gt+1,1);
								PRS->PutRowHeight(_sodu);

								// Cộng thêm 1 dòng duy nhất
								gt++;
							}
						}
						catch(...){}
					}
				}
			}
		}

		CoUninitialize();
	}
	catch(_com_error & error){_s(L"#QL4.102: Lỗi chèn ảnh.");}
}


void cls_main::f_implinkpicsheet()
{
	// Chèn link ảnh lên sheet (hiện hộp thoại tìm kiếm ảnh)
	try
	{
		CoInitialize(NULL);
		xl.GetActiveObject(L"Excel.Application");
		xl->Visible = true;
		
		pWb = xl->GetActiveWorkbook();
		pSheet = pWb->ActiveSheet;
		_code = pSheet->CodeName;
		pRange = pSheet->Cells;

		// Xác định sheet đang active + vị trí
		int iRow = pSheet->Application->ActiveCell->Row;
		int ic1,ic2,ic3;
		BOOLEAN f = TRUE;
		if(_code==(_bstr_t)L"shHSNTCV")
		{
			ic1 = getColumn(pSheet,"DMBB_STT");
			ic2 = getColumn(pSheet,"DMBB_MACV");
			ic3 = getColumn(pSheet,"DMBB_IMG");
		}
		else if(_code==(_bstr_t)L"shHSNTVL")
		{
			ic1 = getColumn(pSheet,"DMVL_STT");
			ic2 = getColumn(pSheet,"DMVL_MAVL");
			ic3 = getColumn(pSheet,"DMVL_IMG");
		}
		else if(_code==(_bstr_t)L"shHSNTGD")
		{
			ic1 = getColumn(pSheet,"DMGD_STT");
			ic2 = getColumn(pSheet,"DMGD_HS");
			ic3 = getColumn(pSheet,"DMGD_IMG");
		}
		else f = FALSE;

		if(f==TRUE)
		{
			// Kiểm tra điều kiện
			CString kt1,kt2;
			kt1 = pRange->GetItem(iRow,ic1);
			kt2 = pRange->GetItem(iRow,ic2);

			if(_wtoi(kt1)>0 && kt2!=L"" && kt2.Find(L"*")==-1)
			{
				_bienpic = L"";

				AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
				Dlg_importpic _dlg;
				_dlg.DoModal();

				pRange->PutItem(iRow,ic3,(_bstr_t)_bienpic);
			}
			else _s(L"Kích chọn dòng chứa công tác để chèn ảnh.");
		}
		else
		{
			_s(L"Tính năng sử dụng tại 3 sheet danh mục nghiệm thu. "
			L"Bạn kích chọn dòng chứa công tác và chạy lệnh chèn ảnh.");
		}

		CoUninitialize();
	}
	catch(_com_error & error){_s(L"#QL4.103: Lỗi chèn ảnh 2.");}
}


CString cls_main::f_load_congviec(int _numcv,_bstr_t _bb)
{
	_code = name_sheet(_bb);
	psBB = xl->Sheets->GetItem(_code);
	prBB = psBB->Cells;
	
	int _cot1,_cot2;
	CString _linkpic=L"";
	if(_bb==(_bstr_t)L"shHSNTCV")
	{
		_cot1 = getColumn(psBB,"DMBB_STT");
		_cot2 = getColumn(psBB,"DMBB_IMG");
	}
	else
	{
		_cot1 = getColumn(psBB,"DMGD_STT");
		_cot2 = getColumn(psBB,"DMGD_IMG");
	}
	
	int xde = FindComment(psBB,_cot1,"END");

	int tang = 8;
	CString ktra = prBB->GetItem(tang,_cot1);
	while (_wtoi(ktra)!=_numcv && tang<xde)
	{
		tang++;
		ktra = prBB->GetItem(tang,_cot1);
	}

	// Xác định mã ảnh
	if(_wtoi(ktra)==_numcv) _linkpic = prBB->GetItem(tang,_cot2);
	return _linkpic;
}


void cls_main::auto_date_bienban()
{
	try
	{
		CoInitialize(NULL);
		xl.GetActiveObject(L"Excel.Application");
		xl->Visible = true;
		xl->PutScreenUpdating(VARIANT_FALSE);

		pWb = xl->GetActiveWorkbook();
		pSheet = pWb->ActiveSheet;
		pRange = pSheet->Cells;
		_code = pSheet->CodeName;

		if(_code==(_bstr_t)L"shHSNTVL" || _code==(_bstr_t)L"shHSNTVL" || _code==(_bstr_t)L"shHSNTVL")
		{
			int nRow = pSheet->Application->ActiveCell->Row;
			int nCol = pSheet->Application->ActiveCell->Column;
			int vt,c0;

			if(_code==(_bstr_t)L"shHSNTCV")
			{
				vt = getColumn(pSheet,"DMBB_MAHS");
				c0 = getColumn(pSheet,"DMBB_STT");
			}
			else if(_code==(_bstr_t)L"shHSNTGD")
			{
				vt = getColumn(pSheet,"DMGD_HS");
				c0 = getColumn(pSheet,"DMGD_STT");
			}
			else
			{
				vt = getColumn(pSheet,"DMVL_MAHS");
				c0 = getColumn(pSheet,"DMVL_STT");
			}

			CString val = pRange->GetItem(nRow,nCol);
			if(nCol==vt && val!=L"")
			{
				// Chèn thêm dòng
				int xde = FindComment(pSheet,c0,"END");
				if(nRow+2>xde)
					pSheet->GetRange(pRange->GetItem(xde,1),pRange->GetItem(xde+2,1))->EntireRow->Insert(xlShiftDown);
			}
		}
	
		//xl->PutScreenUpdating(VARIANT_TRUE);
		CoUninitialize();
	}
	catch(_com_error & error){_s(L"#QL4.133: Lỗi chèn ngày tháng.");}
}


void cls_main::f_FindBrowse()
{
	try
	{
		TCHAR szFilters[]= _T("Danh mục vật liệu (*.csv)|*.csv|All Files (*.*)|*.*||");
		CFileDialog fileDlg(TRUE, _T("csv"), _T("*.csv"),
		OFN_FILEMUSTEXIST | OFN_HIDEREADONLY, szFilters);

		if(fileDlg.DoModal() == IDOK)
		{
			long gtri = f_Read_file_CSV(fileDlg.GetPathName());
			if(gtri>0) _s(L"Thêm mới dữ liệu thành công.");
		}
	}
	catch(_com_error & error){_s(L"Lỗi #QL4.221: Lỗi browse DMVL...");}
}


long cls_main::f_GetLineFile(CString szPath)
{
	FILE * pFile;
	wchar_t *c = (wchar_t *)malloc(2000*sizeof(wchar_t));
	long line = 0;
	pFile=_wfopen(szPath,L"r, ccs=UTF-16LE");
	if (pFile==NULL) perror ("Error opening file");
	else
	{
		while (fgetws(c, 2000, pFile))
		{
			line++;
		}
		fclose (pFile);
		printf ("File contains %d.\n",line);
	}
	free(c);
	return line;
}


long cls_main::f_Read_file_CSV(CString pathCSV)
{
	long cong=0;

	try
	{
	wchar_t **dArray;
	int numberOfLine = 0;
	int i;
	
	long _size;
	_size= f_GetLineFile(pathCSV);
	_size++;
	dArray = (wchar_t **)calloc(_size,sizeof(wchar_t *));
	wchar_t **p;
	p = (wchar_t **)calloc(_size,sizeof(wchar_t *));
	for (i = 0; i < _size; i++)
	{
		*(dArray+i) = (wchar_t *)malloc(SIZE_LINE*sizeof(wchar_t));
		*(p+i) = (wchar_t *)malloc(SIZE_LINE*sizeof(wchar_t));
	}

	FILE *pReadFile;
	if(_wfopen_s(&pReadFile, pathCSV, L"r, ccs=UTF-16LE") != 0)
	{
		MessageBox(NULL,
			L"Không thể mở được file dữ liệu. Vui lòng chọn lại file dữ liệu.",
			L"Thông báo",
			MB_OK | MB_ICONINFORMATION);

		for (i = 0; i < _size; i++)
		{
			free(*(dArray+i));
		}
		free(dArray);
		return 0;
	}	
	else
	{
		i=0;
		CString szVarString=L"";
		while(fgetws(*(p+i), ROW_LINE, pReadFile)) i++;
			
		numberOfLine= i;
		fclose(pReadFile);

		if(getPathQLCL()==L"") return 0;
		if(connectDsn(pathMDB)==-1) return 0;
		ObjConn.OpenAccessDB(pathMDB, L"", L"");
		CRecordset* Recset;

		int fdem=0;
		CString ft[5]={L"",L"",L"",L"",L""};

		// Tạo mã định mức mới	
		CString he[26] = {"a","b","c","d","e","f","g","h","i","j",
			"k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z"};

		int r1 = rand()%26;
		for(int i= 1;i<numberOfLine;i++)
		{
			wcscpy_s(*(dArray+i), wcslen(*(p+i))+1, *(p+i));
			wchar_t *textString = new wchar_t[SIZE_LINE];
			wcscpy_s(textString,wcslen(*(dArray+i))+1, *(dArray+i));
			szVarString= textString;
			int nDulieu=0;

			long length;
			CString szLeft;
			int nToken= szVarString.Find(L'\t');

			fdem=0;
			while(nToken>=0)
			{
				nDulieu++;
				length= szVarString.GetLength();
				szLeft= szVarString.Left(nToken);
				szLeft.TrimLeft();szLeft.TrimRight();
				szVarString=szVarString.Right(length-nToken-1);
				nToken= szVarString.Find(L'\t');

				fdem++;
				szLeft.Replace(L"'",L"");
				szLeft.Replace(L";",L",");
				szLeft.Replace(L"\"",L"");

				if(fdem==1) ft[1]=szLeft;
				else if(fdem==2) ft[2]=szLeft;

				// Thêm vào để lấy cột cuối cùng (by Leo 23.06.15)
				ft[3]=L"";
				if(nToken<0) ft[3]=szVarString;
			}

			ft[1].Trim();
			if(ft[1]!=L"")
			{
				// Kiểm tra dữ liệu có trong CSDL không?
				msg.Format(L"SELECT COUNT(*) AS iDem FROM tbMaVL_tenVL WHERE (MaVL = '%s');",ft[1]);
				Recset = ObjConn.Execute(msg);
				Recset->GetFieldValue(L"iDem",ft[0]);
				Recset->Close();

				while (_wtoi(ft[0])>0)
				{
					// Nếu trùng sẽ tạo ra 1 mã định mức mới
					r1 = rand()%26;
					ft[1] = ft[1] + he[r1];

					msg.Format(L"SELECT COUNT(*) AS iDem FROM tbMaVL_tenVL WHERE (MaVL = '%s');",ft[1]);
					Recset = ObjConn.Execute(msg);
					Recset->GetFieldValue(L"iDem",ft[0]);
					Recset->Close();
				}

				// Load dữ liệu CSV vào CSDL
				msg.Format(L"INSERT INTO tbMaVL_tenVL (MaVL,TenVL,Kichthuocmau) "
					L"VALUES ('%s','%s','%s');",ft[1],ft[2],ft[3]);
				ObjConn.ExecuteRB(msg);

				cong++;
			}
			free(textString);
		}
		fclose(pReadFile);
		ObjConn.CloseDatabase();
	}
	
	//free memory
	for (i = 0; i < _size; i++)
	{
		free(*(dArray+i));
		free(*(p+i));
	}
	free(dArray);
	free(p);
	
	CoUninitialize();
	}
	catch(_com_error & error){_s(L"#QL4.151: Lỗi thêm danh mục vật liệu vào CSDL.");}

	return cong;
}


// Load dữ liệu vật liệu
void cls_main::f_Load_file_vatlieu()
{
	try
	{
		CoInitialize(NULL);
		HRESULT hr = xl.GetActiveObject(L"Excel.Application");  // to connect to running Excep instance
		if(FAILED(hr) && xl == NULL) return;
		pWb = xl->GetActiveWorkbook();

		AFileFilter	ff(L"Excel Files|*.xlsx|All Files|*.*||");
		CFileDialog dlgFile(TRUE,NULL,NULL,OFN_HIDEREADONLY|OFN_NOCHANGEDIR,L"Excel Files|*.xls;*.xlsm;*.xlsx|All Files|*.*||");
		dlgFile.m_ofn.lpstrTitle = _T("Select the Excel File to Load");

		if (dlgFile.DoModal( ) != IDOK) return;

		// Load the Excel application in the background.
		_ApplicationPtr _pApplication;
		if ( FAILED( _pApplication.CreateInstance(_T("Excel.Application")))) return;

		_variant_t varOption((long) DISP_E_PARAMNOTFOUND,VT_ERROR);
		_WorkbookPtr _pWBook = _pApplication->Workbooks->Open(
			(_bstr_t)dlgFile.GetPathName(),
			varOption,varOption,varOption,varOption,varOption,varOption, 
			varOption,varOption,varOption,varOption,varOption,varOption );
	
		if (_pWBook == NULL) return;
		_WorksheetPtr _pSh0 = _pWBook->Sheets->Item[1];
		if (_pSh0 == NULL) return;

		// Load the column headers.
		RangePtr _pRg0 = _pSh0->GetRange(_bstr_t(_T("A1")),_bstr_t(_T("H65000")));
		if (_pRg0 == NULL) return;

		if(getPathCV()==L"") return;
		if(getPathQLCL()==L"") return;		

		if(connectDsn(pathCongviec)==-1) return;
		if(connectDsn(pathMDB)==-1) return;

		ObjConn.OpenAccessDB(pathMDB, L"", L"");
		ObjConnCV.OpenAccessDB(pathCongviec, L"", L"");
		CRecordset* Recset;

		int slg=0,cong=0;
		CString sDem=L"",sTVan=L"",val=L"";
		CString ft[10];

		for (int i=2;i<=65000;i++)
		{
			// Mã số
			ft[1] = _pRg0->GetItem(i,1);
			if(ft[1]!=L"")
			{
				// Kiểm tra xem mã vật liệu đã tồn tại chưa?
				sTVan.Format(L"SELECT COUNT(*) AS iDem FROM tbTuDienVatTu WHERE MSVT LIKE '%s';",ft[1]);
				Recset = ObjConnCV.Execute(sTVan);
				Recset->GetFieldValue(L"iDem",sDem);
				Recset->Close();

				if(_wtoi(sDem)==0)
				{
					// Tên vật liệu
					ft[2] = _pRg0->GetItem(i,2);
					ft[2].Replace(L"'",L"");
					ft[2].Trim();

					// Đơn vị tính
					ft[8] = _pRg0->GetItem(i,8);
					ft[8].Replace(L"'",L"");
					ft[8].Trim();

					// Load dữ liệu vật liệu từ CSV vào CSDL
					sTVan.Format(L"INSERT INTO tbTuDienVatTu (MSVT,TVT,DVT) "
						L"VALUES ('%s','%s','%s');",ft[1],ft[2],ft[8]);
					ObjConnCV.ExecuteRB(sTVan);

					// Kích thước mẫu
					ft[3] = _pRg0->GetItem(i,3);
					ft[3].Replace(L"'",L"");
					ft[3].Trim();

					sTVan.Format(L"INSERT INTO tbMaVL_tenVL (MaVL,Kichthuocmau) VALUES ('%s','%s');",ft[1],ft[3]);
					ObjConn.ExecuteRB(sTVan);

					slg++;

					// Lưu mã tiêu chuẩn
					int demtc=0,dempp=0;
					for (int j = i; j <= i+30; j++)
					{
						ft[1] = _pRg0->GetItem(j,1); ft[1].Trim();		// Mã vật liệu
						ft[4] = _pRg0->GetItem(j,4); ft[4].Trim();		// Mã tiêu chuẩn
						ft[6] = _pRg0->GetItem(j,6); ft[6].Trim();		// Phương pháp lấy mẫu
						ft[6].Replace(L"'",L"");

						if((j>i && ft[0]!=L"") || (ft[4]==L"" && ft[6]==L"")) break;

						if(ft[4]!=L"")
						{
							// Kiểm tra tiêu chuẩn có trong CSDL chưa? -> Thêm vào bảng tbTentieuchuan
							sTVan.Format(L"SELECT COUNT(*) AS iDem FROM tbTentieuchuan WHERE MaTC LIKE '%s';",ft[4]);
							Recset = ObjConn.Execute(sTVan);
							Recset->GetFieldValue(L"iDem",sDem);
							Recset->Close();

							// Chưa tồn tại mã tiêu chuẩn --> Thêm mới
							if(_wtoi(sDem)==0)
							{
								// Tên tiêu chuẩn
								ft[5] = _pRg0->GetItem(j,5);
								ft[5].Replace(L"'",L"");
								ft[5].Trim();
								if(ft[5]==L"") ft[5]=ft[4];

								sTVan.Format(L"INSERT INTO tbTentieuchuan (MaTC,TenTC) VALUES ('%s','%s');",ft[4],ft[5]);
								ObjConn.ExecuteRB(sTVan);
							}

							// Thêm vào bảng tbMaVL_Tieuchuan
							demtc++;
							if(demtc<9) val.Format(L"0%d",demtc);
							else val.Format(L"%d",demtc);

							sTVan.Format(L"INSERT INTO tbMaVL_Tieuchuan (ID,MaVL,MaTC) "
								L"VALUES ('%s','%s','%s');",val,ft[1],ft[4]);
							ObjConn.ExecuteRB(sTVan);
						}

						if(ft[6]!=L"")
						{
							dempp++;
							if(dempp<9) val.Format(L"%s.0%d",ft[1],dempp);
							else val.Format(L"%s.%d",ft[1],dempp);

							ft[7] = _pRg0->GetItem(j,7); ft[7].Trim(); ft[7].Replace(L"'",L"");
							ft[8] = _pRg0->GetItem(j,8); ft[8].Trim(); ft[8].Replace(L"'",L"");

							// Lưu phương pháp lấy mẫu
							sTVan.Format(L"INSERT INTO tbMau (MaLayMau,Phuongphaplaymau,TansuatLaymau,Donvi) "
								L"VALUES ('%s','%s','%s','%s');",val,ft[6],ft[7],ft[8]);
							ObjConn.ExecuteRB(sTVan);

							if(dempp<9) msg.Format(L"0%d",dempp);
							else msg.Format(L"%d",dempp);

							sTVan.Format(L"INSERT INTO tbMaVL_Quydinhlaymau (ID,MaVL,MaLayMau) "
								L"VALUES ('%s','%s','%s');",msg,ft[1],val);
							ObjConn.ExecuteRB(sTVan);
						}
					}
				}

				cong=0;
			}
			else cong++;

			if(cong==30) break;
		}

		ObjConn.CloseDatabase();
		ObjConnCV.CloseDatabase();

		if(slg>0) _s(L"Thêm mới dữ liệu thành công.");
		else _s(L"Không tồn tại dữ liệu thêm mới.");

		// Don't save any inadvertant changes to the .xls file.
		_pWBook->Close(VARIANT_FALSE);

		// Need to quit, otherwise Excel remains active and locks the .xls file.
		_pApplication->Quit();

		CoUninitialize();
	}
	catch(_com_error & error){_s(L"Lỗi load dữ liệu Excel.");}
}


// Load dữ liệu công việc
void cls_main::f_Load_file_congviec()
{
	try
	{
		CoInitialize(NULL);
		HRESULT hr = xl.GetActiveObject(L"Excel.Application");  // to connect to running Excep instance
		if(FAILED(hr) && xl == NULL) return;
		pWb = xl->GetActiveWorkbook();

		AFileFilter	ff(L"Excel Files|*.xlsx|All Files|*.*||");
		CFileDialog dlgFile(TRUE,NULL,NULL,OFN_HIDEREADONLY|OFN_NOCHANGEDIR,L"Excel Files|*.xls;*.xlsm;*.xlsx|All Files|*.*||");
		dlgFile.m_ofn.lpstrTitle = _T("Select the Excel File to Load");

		if (dlgFile.DoModal( ) != IDOK) return;
		
		// Load the Excel application in the background.
		_ApplicationPtr _pApplication;
		if ( FAILED( _pApplication.CreateInstance(_T("Excel.Application")))) return;
	
		_variant_t	varOption((long) DISP_E_PARAMNOTFOUND,VT_ERROR);
		_WorkbookPtr _pWBook = _pApplication->Workbooks->Open(
			(_bstr_t)dlgFile.GetPathName(),
			varOption,varOption,varOption,varOption,varOption,varOption, 
			varOption,varOption,varOption,varOption,varOption,varOption );
	
		if (_pWBook == NULL) return;
		_WorksheetPtr _pSh0 = _pWBook->Sheets->Item[1];
		if (_pSh0 == NULL) return;

		// Load the column headers.
		RangePtr _pRg0 = _pSh0->GetRange(_bstr_t(_T("A1")),_bstr_t(_T("E65000")));
		if (_pRg0 == NULL) return;

		if(getPathCV()==L"") return;
		if(getPathQLCL()==L"") return;		

		if(connectDsn(pathCongviec)==-1) return;
		if(connectDsn(pathMDB)==-1) return;

		ObjConn.OpenAccessDB(pathMDB, L"", L"");
		ObjConnCV.OpenAccessDB(pathCongviec, L"", L"");
		CRecordset* Recset;

		int slg=0,cong=0;
		CString sDem=L"",sRand=L"",sTVan=L"",val=L"";
		CString ft[10];

		for (int i=2;i<=65000;i++)
		{
			// Mã số
			ft[1] = _pRg0->GetItem(i,1);
			if(ft[1]!=L"")
			{
				// Kiểm tra xem mã công việc đã tồn tại chưa?
				sTVan.Format(L"SELECT COUNT(*) AS iDem FROM tbDonGia WHERE MHDG LIKE '%s';",ft[1]);
				Recset = ObjConnCV.Execute(sTVan);
				Recset->GetFieldValue(L"iDem",sDem);
				Recset->Close();

				if(_wtoi(sDem)==0)
				{
					// Nội dung công việc
					ft[2] = _pRg0->GetItem(i,2);
					ft[2].Replace(L"'",L"");
					ft[2].Trim();

					// Đơn vị tính
					ft[3] = _pRg0->GetItem(i,3);
					ft[3].Replace(L"'",L"");
					ft[3].Trim();

					// Load dữ liệu vật liệu từ CSV vào CSDL
					sTVan.Format(L"INSERT INTO tbDonGia (MHDG,TCV,DVT) VALUES ('%s','%s','%s');",ft[1],ft[2],ft[3]);
					ObjConnCV.ExecuteRB(sTVan);

					slg++;

					// Lưu mã tiêu chuẩn
					int demtc=0,dempp=0;
					for (int j = i; j <= i+30; j++)
					{
						ft[2] = _pRg0->GetItem(j,1); ft[2].Trim();		// Tên công việc
						ft[4] = _pRg0->GetItem(j,4); ft[4].Trim();		// Mã tiêu chuẩn
						ft[6] = _pRg0->GetItem(j,6); ft[6].Trim();		// Nội dung kiểm tra
						ft[2].Replace(L"'",L"");
						ft[6].Replace(L"'",L"");

						if((j>i && ft[2]!=L"") || (ft[4]==L"" && ft[6]==L"")) break;

						if(ft[4]!=L"")
						{
							// Kiểm tra tiêu chuẩn có trong CSDL chưa? -> Thêm vào bảng tbTentieuchuan
							sTVan.Format(L"SELECT COUNT(*) AS iDem FROM tbTentieuchuan WHERE MaTC LIKE '%s';",ft[4]);
							Recset = ObjConn.Execute(sTVan);
							Recset->GetFieldValue(L"iDem",sDem);
							Recset->Close();

							// Chưa tồn tại mã tiêu chuẩn --> Thêm mới
							if(_wtoi(sDem)==0)
							{
								// Tên tiêu chuẩn
								ft[5] = _pRg0->GetItem(j,5);
								ft[5].Replace(L"'",L"");
								ft[5].Trim();
								if(ft[5]==L"") ft[5]=ft[4];

								sTVan.Format(L"INSERT INTO tbTentieuchuan (MaTC,TenTC) VALUES ('%s','%s');",ft[4],ft[5]);
								ObjConn.ExecuteRB(sTVan);
							}

							// Thêm vào bảng tbMaCV_Tieuchuan
							demtc++;

							if(demtc<9) val.Format(L"0%d",demtc);
							else val.Format(L"%d",demtc);

							sTVan.Format(L"INSERT INTO tbMaCV_Tieuchuan (ID,MaCV,MaTC) VALUES ('%s','%s','%s');",val,ft[1],ft[4]);
							ObjConn.ExecuteRB(sTVan);
						}

						if(ft[6]!=L"")
						{
							dempp++;

							if(dempp<9) msg.Format(L"0%d",dempp);
							else msg.Format(L"%d",dempp);

							// Phương pháp kiểm tra
							ft[7] = _pRg0->GetItem(j,7);
							ft[7].Replace(L"'",L"");
							ft[7].Trim();

							sTVan.Format(L"INSERT INTO tbMaCV_NDnghiemthu (ID,MaCV,NoidungKiemtra,PhuongphapKiemtra)"
								L"VALUES ('%s','%s','%s','%s');",msg,ft[1],ft[6],ft[7]);
							ObjConn.ExecuteRB(sTVan);
						}
					}
				}

				cong=0;
			}
			else cong++;

			if(cong==30) break;
		}

		ObjConn.CloseDatabase();
		ObjConnCV.CloseDatabase();

		if(slg>0) _s(L"Thêm mới dữ liệu thành công.");
		else _s(L"Không tồn tại dữ liệu thêm mới.");

		// Don't save any inadvertant changes to the .xls file.
		_pWBook->Close(VARIANT_FALSE);

		// Need to quit, otherwise Excel remains active and locks the .xls file.
		_pApplication->Quit();

		CoUninitialize();
	}
	catch(_com_error & error){_s(L"Lỗi load dữ liệu Excel.");}
}


void cls_main::f_xacdinh_danhmuc()
{
	try
	{
		CoInitialize(NULL);
		xl.GetActiveObject(L"Excel.Application");
		xl->Visible = true;
		xl->PutScreenUpdating(VARIANT_FALSE);

		pWb = xl->GetActiveWorkbook();
		psDMCV = pWb->ActiveSheet;
		prDMCV = psDMCV->Cells;
		shDMCV = psDMCV->CodeName;

		if(shDMCV==(_bstr_t)L"shHSNTCV" || shDMCV==(_bstr_t)L"shHSNTVL")
		{
			xl->StatusBar = (_bstr_t)L"Đang tiến hành phân tích dữ liệu. Vui lòng đợi trong giây lát...";
			nDbo=0;	// Số lượng đồng bộ dữ liệu

			// Xác định cột & END
			if(shDMCV==(_bstr_t)L"shHSNTCV")
			{
				_ivtr[1] = getColumn(psDMCV,"DMBB_STT");
				_ivtr[2] = getColumn(psDMCV,"DMBB_MACV");
				_ivtr[3] = getColumn(psDMCV,"DMBB_ND");
				_ivtr[4] = getColumn(psDMCV,"DMBB_TC");
				_ivtr[5] = getColumn(psDMCV,"DMBB_TC2");
			}
			else
			{
				_ivtr[1] = getColumn(psDMCV,"DMVL_STT");
				_ivtr[2] = getColumn(psDMCV,"DMVL_MAVL");
				_ivtr[3] = getColumn(psDMCV,"DMVL_ND");
				_ivtr[4] = getColumn(psDMCV,"DMVL_TC");
				_ivtr[5] = getColumn(psDMCV,"DMVL_TC2");
			}

			// Xác định vị trí END
			_ivtr[0] = FindComment(psDMCV,_ivtr[1],"END");

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
						f_dong_bo_danhmuc(_wtoi(_arrSLS[i]));
					}
					else
					{
						// Có nhiều lựa chọn
						int jlen = _arrSLS[i].GetLength();
						int jbd = _wtoi(_arrSLS[i].Left(_pos));
						int jkt = _wtoi(_arrSLS[i].Right(jlen-_pos-1));
						
						for (int j = jbd; j <=jkt; j++) f_dong_bo_danhmuc(j);
					}
				}

				if(nDbo>0)
				{
					int k;
					CString _val1,_val2;
					for (int i = _ivtr[0]-1; i >=8; i--)
					{
						_val1 = prDMCV->GetItem(i,_ivtr[1]);
						_val2 = prDMCV->GetItem(i,_ivtr[2]);
						if(_wtoi(_val1)>0 && _val2!=L"" && _val2.Find(L"*")==-1)
						{
							// Kiểm tra mã được đồng bộ
							int _gtri=0;
							for (int j = 1; j <= nDbo; j++)
							{
								if(_val2==DBO[j].sMa)
								{
									_gtri=j;
									break;
								}
							}

							if(_gtri>0)
							{
								msg.Format(L"Đang tiến hành đồng bộ liệu '%s'. Vui lòng đợi trong giây lát...",DBO[_gtri].sMa);
								xl->StatusBar = (_bstr_t)msg;


								// Đồng bộ tên vật liệu hoặc công việc
								prDMCV->PutItem(i,_ivtr[3],(_bstr_t)DBO[_gtri].sTen);

								// Kiểm tra số lượng dòng cần chèn nếu có
								k= 0;
								for (int j = i+1; j <= _ivtr[0]; j++)
								{
									_val1 = prDMCV->GetItem(j,_ivtr[1]);
									_val2 = prDMCV->GetItem(j,_ivtr[2]);
									if((_wtoi(_val1)>0 && _val2!=L"" && _val2.Find(L"*")==-1)  || _val2.MakeUpper()==L"GD")
									{
										k=j;
										break;
									}
								}

								if(k==0) k=_ivtr[0];
								
								// Xóa tiêu chuẩn cũ
								psDMCV->GetRange(prDMCV->GetItem(i,_ivtr[4]),prDMCV->GetItem(k-1,_ivtr[4]))->ClearContents();
								psDMCV->GetRange(prDMCV->GetItem(i,_ivtr[5]),prDMCV->GetItem(k-1,_ivtr[5]))->ClearContents();

								if(DBO[_gtri].iNumTC-k+i>0)
								{
									// Chèn dòng
									PRS = psDMCV->GetRange(prDMCV->GetItem(k,1),prDMCV->GetItem(i+DBO[_gtri].iNumTC-1,1));
									PRS->EntireRow->Insert(xlShiftDown);
								}

								// Đồng bộ mã & tên tiêu chuẩn
								k=0;
								for (int j = i; j < i+DBO[_gtri].iNumTC; j++)
								{
									k++;
									prDMCV->PutItem(j,_ivtr[4],(_bstr_t)DBO[_gtri].tch[k]);
									prDMCV->PutItem(j,_ivtr[5],(_bstr_t)DBO[_gtri].mch[k]);
									
									_val1 = prDMCV->GetItem(j,_ivtr[4]+1);
									if(_val1==L"") prDMCV->PutItem(j,_ivtr[4]+1,(_bstr_t)L" ");

									_val1 = prDMCV->GetItem(j,_ivtr[5]+1);
									if(_val1==L"") prDMCV->PutItem(j,_ivtr[5]+1,(_bstr_t)L" ");
								}
							}
						}
					}

				MessageBox(NULL,L"Đồng bộ dữ liệu thành công.",L"Thông báo",MB_OK | MB_ICONINFORMATION);
				xl->StatusBar = (_bstr_t)L"Ready";
				}
			}
		}
		else _s(L"Sử dụng chức năng này tại sheet danh mục vật liệu hoặc danh mục công việc.");

		//xl->PutScreenUpdating(VARIANT_TRUE);
		CoUninitialize();
	}
	catch(_com_error & error){}
}


void cls_main::f_dong_bo_danhmuc(int _vtri)
{
	CString _val1,_val2;
	_val1 = prDMCV->GetItem(_vtri,_ivtr[1]);
	_val2 = prDMCV->GetItem(_vtri,_ivtr[2]);

	if(_wtoi(_val1)>0 && _val2!=L"" && _vtri<_ivtr[0])
	{
		// Kiểm tra mã có bị trùng lặp không?
		BOOLEAN f0=TRUE;
		if(nDbo>0) 
		for (int i = 1; i <= nDbo; i++){if(_val2==DBO[i].sMa) f0=FALSE;}

		if(f0==TRUE)
		{
			// Lưu mã & tên
			nDbo++;
			DBO[nDbo].sMa = _val2;
			DBO[nDbo].sTen = prDMCV->GetItem(_vtri,_ivtr[3]);

			// Xác định số lượng tiêu chuẩn & chi tiết tiêu chuẩn
			int k=0;
			for (int i = _vtri+1; i <= _ivtr[0]; i++)
			{
				_val1 = prDMCV->GetItem(i,_ivtr[1]);
				_val2 = prDMCV->GetItem(i,_ivtr[2]);
				if((_wtoi(_val1)>0 && _val2!=L"") || _val2.MakeUpper()==L"GD")
				{
					k=i;
					break;
				}
			}

			if(k==0) k=_ivtr[0];
			DBO[nDbo].iNumTC=0;
			CString sTTC[100],sMTC[100];
			for (int i = _vtri; i < k; i++)
			{
				_val1 = prDMCV->GetItem(i,_ivtr[4]);
				_val1.Trim();

				if(_val1!=L"")
				{
					DBO[nDbo].iNumTC++;
					DBO[nDbo].tch[DBO[nDbo].iNumTC] = _val1;
					DBO[nDbo].mch[DBO[nDbo].iNumTC] = prDMCV->GetItem(i,_ivtr[5]);
				}
			}
		}
	}
}


void cls_main::f_even_enter()
{
	try
	{
		CoInitialize(NULL);
		xl.GetActiveObject(L"Excel.Application");
		xl->Visible = true;
		xl->PutScreenUpdating(VARIANT_FALSE);
		xl->PutEnableCancelKey(XlEnableCancelKey(FALSE));
		xl->PutStatusBar((_bstr_t)L"Enter");

		pWb = xl->GetActiveWorkbook();
		pSheet = pWb->ActiveSheet;
		pRange = pSheet->Cells;
		_code = pSheet->CodeName;

		if(_code == (_bstr_t)L"shHSNTCV")
		{
			// SHEET DANH MỤC CÔNG VIỆC
			Dlg_Open_DMCV _fcall;
			_fcall.f_Tim_kiem_DMCV();
		}
		else if(_code == (_bstr_t)L"shHSNTVL")
		{
			// SHEET DANH MỤC VẬT LIỆU
			Dlg_Open_DSVL _fcall;
			_fcall.f_Tim_kiem_DMVL();
		}
		else if(_code == (_bstr_t)L"shHSNTGD")
		{
			// SHEET DANH MỤC GIAI ĐOẠN
			f_enter_giaidoan();
		}
		else if(_code == (_bstr_t)L"shKhoiluong")
		{
			// Sheet khối lượng
			Enter_click_KLG();
		}
		else if(_code == (_bstr_t)L"shNhomTC")
		{
			Dlg_trauu_tieuchuan _fcall;
			_fcall.xl_timkiem_tieuchuan();
		}
		else if(_code == (_bstr_t)L"shCPTLuong")
		{
			// Sheet tiền lương
			autofill_TL();
		}
		else if(_code == (_bstr_t)L"shTDCViec")
		{
			// Sheet tiến độ
			auto_tiendo();
		}
		//else
		//{
		//	// Tổng hợp các sheet gõ keyboard
		//	f_keyboard();
		//}

		//xl->StatusBar = (_bstr_t)L"Ready";

		SetFindEND(pSheet->Cells);
		CoUninitialize();
	}
	catch(...){}
}


// Leo 07.02.2018
void cls_main::Enter_click_KLG()
{
	try
	{
		xl->PutScreenUpdating(VARIANT_FALSE);
		pRange = pSheet->Cells;
		int iRow = pSheet->Application->ActiveCell->Row;
		int iCol = pSheet->Application->ActiveCell->Column;

		// Xác định vị trí giai đoạn
		int ibd=0;
		int iCotSTT = getColumn(pSheet,L"DTXD_STT");
		CString szval=L"";
		for (int i = iRow; i >= 1; i--)
		{
			szval = GIOR(i,iCotSTT,pRange,L"");
			if(szval==L"STT")
			{
				ibd = i-4;
				break;
			}
		}

		if(ibd==0) return;

		CString sadd=L"";
		PRS = pRange->GetItem(ibd,iCotSTT);
		if(PRS->GetComment()!=NULL)
		{
			_bstr_t bscmt = PRS->GetComment()->Text();
			sadd = (LPCTSTR)(bscmt);
			sadd.Replace(L"DAU",L"CUOI");
		}

		if(sadd==L"") return;

		int iTencv = getColumn(pSheet,L"DTXD_TENCV");
		int ips = FindCEnd(pSheet,iTencv,"NHD",ibd+1);
		int ikt = FindComment(pSheet,iCotSTT,(_bstr_t)sadd);

		// Kiểm tra có nằm trong vùng giai đoạn không?
		if(iRow<=ibd+7 || iRow==ips || iRow>=ikt) return;

		int iCotMH = getColumn(pSheet,L"DTXD_MH");
		if (iCol==iCotMH)
		{
			szval = GIOR(iRow+1,iTencv,pRange,L"");
			if(szval!=L"" || iRow+1 == ikt)
			{
				PRS = pRange->GetItem(iRow+1,1);
				PRS->EntireRow->Insert(xlShiftDown);
			}

			HINSTANCE loadDLL = LoadLibrary(L"QLCLTracuu.dll");
			typedef void (__stdcall *p)(int opt);
			p pcall = (p)GetProcAddress(loadDLL, "call_Timkiem_DTXD");
			if(pcall!=NULL) pcall(0);
			FreeLibrary(loadDLL);

			// Đánh lại STT
			int dem=0, jvtbd=ibd+7, jvtkt=ips;
			if(iRow>ips)
			{
				jvtbd = ips+1;
				jvtkt = ikt;
			}

			for (int i = jvtbd; i < jvtkt; i++)
			{
				szval = GIOR(i,iCotMH,pRange,L"");
				if(szval!=L"")
				{
					dem++;
					pRange->PutItem(i,1,dem);
				}				
			}

			return;
		}

		// Kiểm tra có phải công tác không?
		szval = GIOR(iRow,iCotMH,pRange,L"");
		if(szval!=L"") return;		

		// Xác định vị trí công tác tương ứng
		int ivtri=0;
		for (int i = iRow-1; i >= ibd+8; i--)
		{
			szval = GIOR(i,iCotMH,pRange,L"");
			if(szval!=L"")
			{
				ivtri=i;
				break;
			}

			if(i==ips) return;
		}

		if(ivtri==0) return;

		// Xác định vị trí công tác kế dưới (nếu có)
		int iKeduoi=ikt;
		for (int i = iRow+1; i < ikt; i++)
		{
			szval = GIOR(i,iCotMH,pRange,L"");
			if(szval!=L"")
			{
				iKeduoi=i;
				break;
			}

			if(i==ips)
			{
				iKeduoi=ips;
				break;
			}
		}

////////////////////// Xác định được tất cả các điều kiện --> Chạy lệnh

		shTS = name_sheet("shTS");
		psTS=xl->Sheets->GetItem(shTS);
		prTS=psTS->GetCells();
		
		int icheckME = _wtoi(getVCell(psTS, L"CF_TDGME"));	// icheckME=-1 CHECK | =0 UNCHECK

		// Chèn thêm dòng nếu có thể
		if(iRow+1==ips || iRow+1==ikt)
		{
			PRS = pRange->GetItem(iRow+1,1);
			PRS->EntireRow->Insert(xlShiftDown);

			iKeduoi++;
			if(iRow+1==ips) ips++;			
			ikt++;
		}

		// In nghiêng
		PRS = pRange->GetItem(iRow,1);
		PRS->EntireRow->Font->PutItalic(1);
		PRS->EntireRow->Font->PutBold(0);
		
		int iCotKL = getColumn(pSheet,L"DTXD_KL");
		int iCotKLTB = getColumn(pSheet,L"DTXD_KLTB");
		int iCotKLNT = getColumn(pSheet,L"DTXD_GHC");
		int iCotCLKL = getColumn(pSheet,L"DTXD_CLKL");
		int iCotBP = getColumn(pSheet,L"DTXD_BP");
		int iCotDai = getColumn(pSheet,L"DTXD_DAI");
		int iCotCao = getColumn(pSheet,L"DTXD_CAO");
		int iCotHS = getColumn(pSheet,L"DTXD_HS");
		int iCotKLBP = getColumn(pSheet,L"DTXD_KLBP");

		if(iCol==iTencv || (iCol>=iCotBP && iCol<iCotKLBP))
		{
			// KL toàn bộ
			if(iCol==iTencv)
			{
				szval.Format(L"=kl(%s)",GAR(iRow,iTencv,pRange,0));
				pRange->PutItem(iRow,iCotKLTB,(_bstr_t)szval);
			}
			else
			{
				if(icheckME==0)
				{
					szval.Format(L"=PRODUCT(%s:%s,%s)",
						GAR(iRow,iCotDai,pRange,0),GAR(iRow,iCotCao,pRange,0),GAR(iRow,iCotHS,pRange,0));
				}
				else
				{
					szval=L"";
					if(iCotCao+1<iCotHS)
					{
						szval = L"=";
						for (int i = iCotCao+1; i < iCotHS; i++)
						{
							szval+=L"kl(";
							szval+=GAR(iRow,i,pRange,0);
							szval+=L")";					
							if(i<iCotHS-1) szval+=L"+";
						}
					}					
				}
				pRange->PutItem(iRow,iCotKLBP,(_bstr_t)szval);
				
				szval.Format(L"=PRODUCT(%s,%s)",GAR(iRow,iCotBP,pRange,0),GAR(iRow,iCotKLBP,pRange,0));
				pRange->PutItem(iRow,iCotKLTB,(_bstr_t)szval);
			}

			// KL nghiệm thu
			szval.Format(L"=%s",GAR(iRow,iCotKLTB,pRange,0));
			pRange->PutItem(iRow,iCotKLNT,(_bstr_t)szval);

			// SUM KL toàn bộ
			// Xác định lại vị trí kế dưới
			for (int i = iKeduoi-1; i > ivtri; i--)
			{
				szval = GIOR(i,iCotKLTB,pRange,L"Error");
				if(szval!=L"")
				{
					iKeduoi=i;
					break;
				}
			}

			szval.Format(L"=SUM(%s:%s)",GAR(ivtri+1,iCotKLTB,pRange,0),GAR(iKeduoi,iCotKLTB,pRange,0));
			pRange->PutItem(ivtri,iCotKLTB,(_bstr_t)szval);
			PRS = pRange->GetItem(ivtri,iCotKLTB);
			PRS->Font->PutBold(1);

			// KL nghiệm thu (tổng)
			szval.Format(L"=%s",GAR(ivtri,iCotKLTB,pRange,0));
			pRange->PutItem(ivtri,iCotKLNT,(_bstr_t)szval);

			// Chênh lệch (tổng)
			szval.Format(L"=%s-%s",GAR(ivtri,iCotKLNT,pRange,0),GAR(ivtri,iCotKL,pRange,0));
			pRange->PutItem(ivtri,iCotCLKL,(_bstr_t)szval);
		}
		else if(iCol==iCotKLTB)
		{
			// Nhập trực tiếp khối lượng toàn bộ
			
			// KL nghiệm thu
			szval.Format(L"=%s",GAR(iRow,iCotKLTB,pRange,0));
			pRange->PutItem(iRow,iCotKLNT,(_bstr_t)szval);

			// Chênh lệch (tổng)
			szval.Format(L"=%s-%s",GAR(iRow,iCotKLNT,pRange,0),GAR(ivtri,iCotKL,pRange,0));
			pRange->PutItem(iRow,iCotCLKL,(_bstr_t)szval);
		}
	}
	catch(...){}
}


// LEO 22.04.2017
void cls_main::autofill_TL()
{
	try
	{
		int iCol = pSheet->Application->ActiveCell->Column;
		if(iCol!=4) return;

		int iRow = pSheet->Application->ActiveCell->Row;
		int iEnd = FindCEnd(pSheet,2,"END",iRow+1);
		if(iEnd==0) return;

		xl->PutScreenUpdating(VARIANT_FALSE);

		pRange = pSheet->Cells;
		CString szval= pRange->GetItem(iRow,iCol);
		szval.Trim();
		if(szval==L"") return;

		// Viết hoa chữ thường
		f_proper_string(pRange,iRow,iCol);

		PRS = pRange->GetItem(iRow,iCol);
		PRS->Validation->Delete();

		// Đổ dữ liệu
		pRange->PutItem(iRow,5,(_bstr_t)L"Chuyên gia");
		
		szval.Format(L"=%s*%s/%s",GAR(iRow,6,pRange,0),GAR(2,16,pRange,3),GAR(3,16,pRange,3));
		pRange->PutItem(iRow,7,(_bstr_t)szval);

		szval.Format(L"=%s*%s",GAR(iRow,7,pRange,0),GAR(11,16,pRange,3));
		pRange->PutItem(iRow,8,(_bstr_t)szval);

		szval.Format(L"=%s*%s/%s",GAR(iRow,9,pRange,0),GAR(2,16,pRange,3),GAR(3,16,pRange,3));
		pRange->PutItem(iRow,10,(_bstr_t)szval);

		szval.Format(L"=%s+%s+%s",GAR(iRow,7,pRange,0),GAR(iRow,8,pRange,0),GAR(iRow,10,pRange,0));
		pRange->PutItem(iRow,11,(_bstr_t)szval);

		// Đánh lại stt
		int iBDAU=6;
		for (int i = iRow; i >=3; i--)
		{
			szval = pRange->GetItem(i,2);
			if(szval==L"STT")
			{
				iBDAU=i+1;
				break;
			}
		}

		int stt=1;
		for (int i = iBDAU; i < iEnd;i++ )
		{
			szval = pRange->GetItem(i,4);
			if(szval!=L"")
			{
				szval.Format(L"CG%d",stt);
				pRange->PutItem(i,2,stt);
				pRange->PutItem(i,3,(_bstr_t)szval);
				stt++;
			}
		}

		if(iRow+1==iEnd)
		{
			// Chèn dòng
			PRS = pSheet->GetRange(pRange->GetItem(iEnd,1),pRange->GetItem(iEnd,11));
			PRS->Insert(xlShiftDown);
		}

		// Định dạng lại bảng
		PRS = pSheet->GetRange(pRange->GetItem(iBDAU,2),pRange->GetItem(iEnd-1,11));
		PRS->Borders->GetItem(XlBordersIndex::xlEdgeLeft)->Weight=xlThin;
		PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->Weight=xlThin;
		PRS->Borders->GetItem(XlBordersIndex::xlEdgeRight)->Weight=xlThin;
		PRS->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->Weight=xlThin;
		PRS->Borders->GetItem(XlBordersIndex::xlInsideVertical)->Weight=xlThin;

		PRS->Borders->GetItem(XlBordersIndex::xlInsideHorizontal)->PutWeight(xlThin);
		PRS->Borders->GetItem(XlBordersIndex::xlInsideHorizontal)->PutLineStyle(xlDot);

		PRS->Font->PutBold(0);
	}
	catch(...){}
}


// LEO 04.05.2017
void cls_main::auto_tiendo()
{
	try
	{
		int iCol = pSheet->Application->ActiveCell->Column;
		if(iCol!=3) return;

		int iRow = pSheet->Application->ActiveCell->Row;
		int iEnd = FindCEnd(pSheet,2,"END",iRow+1);
		if(iEnd==0) return;

		xl->PutScreenUpdating(VARIANT_FALSE);

		pRange = pSheet->Cells;
		CString szval= pRange->GetItem(iRow,iCol);
		szval.Trim();
		if(szval==L"") return;

		// Kiểm tra vị trí nhập có phải dòng đầu tiên không?
		int iVtri=0;
		CString sk1=L"",sk2=L"";
		sk1= pRange->GetItem(iRow,2);
		sk2= pRange->GetItem(iRow-2,2);
		if(_wtoi(sk1)==1 || sk2==L"STT") iVtri=1;

///////// Sao chép & đổ công thức ///////////

		// Bắt đầu
		if(iVtri==1) szval=L"1";
		else
		{
			szval.Format(L"=IF(%s=\"\",\"\",ABS(%s-%s+1))",
				GAR(iRow,5,pRange,0),GAR(iRow,5,pRange,0),GAR(iRow-1,5,pRange,3));
		}
		pRange->PutItem(iRow,6,(_bstr_t)szval);

		// Ngày kết thúc
		szval.Format(L"=IF(%s=\"\",\"\",%s+%s-1)",
			GAR(iRow,9,pRange,0),GAR(iRow,5,pRange,3),GAR(iRow,9,pRange,0));
		pRange->PutItem(iRow,8,(_bstr_t)szval);

		// Kết thúc
		szval.Format(L"=IF(%s=\"\",\"\",%s+%s-1)",
			GAR(iRow,6,pRange,0),GAR(iRow,6,pRange,0),GAR(iRow,7,pRange,0));
		pRange->PutItem(iRow,9,(_bstr_t)szval);

		// Bắt đầu thực tế
		szval.Format(L"=%s",GAR(iRow,6,pRange,0));
		pRange->PutItem(iRow,10,(_bstr_t)szval);

		// Thời gian
		szval.Format(L"=%s",GAR(iRow,7,pRange,0));
		pRange->PutItem(iRow,11,(_bstr_t)szval);

		// Tỷ lệ hoàn thành
		szval.Format(L"=100%s",L"%");
		pRange->PutItem(iRow,12,(_bstr_t)szval);

////////////////////////////////////////////

		// Đánh lại stt
		int iBDAU=8;
		for (int i = iRow; i >=3; i--)
		{
			szval = pRange->GetItem(i,2);
			if(szval==L"STT")
			{
				iBDAU=i+2;
				break;
			}
		}

		int stt=1;
		for (int i = iBDAU; i < iEnd;i++ )
		{
			szval = pRange->GetItem(i,3);
			if(szval!=L"")
			{
				pRange->PutItem(i,2,stt);
				stt++;
			}
		}

		if(iRow+1==iEnd)
		{
			// Chèn dòng
			PRS = pSheet->GetRange(pRange->GetItem(iEnd,1),pRange->GetItem(iEnd,12));
			PRS->Insert(xlShiftDown);
		}

		// Định dạng lại bảng
		PRS = pSheet->GetRange(pRange->GetItem(iBDAU,2),pRange->GetItem(iEnd-1,12));
		PRS->Borders->GetItem(XlBordersIndex::xlEdgeLeft)->Weight=xlThin;
		PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->Weight=xlThin;
		PRS->Borders->GetItem(XlBordersIndex::xlEdgeRight)->Weight=xlThin;
		PRS->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->Weight=xlThin;
		PRS->Borders->GetItem(XlBordersIndex::xlInsideVertical)->Weight=xlThin;

		PRS->Borders->GetItem(XlBordersIndex::xlInsideHorizontal)->PutWeight(xlThin);
		PRS->Borders->GetItem(XlBordersIndex::xlInsideHorizontal)->PutLineStyle(xlDot);

		PRS->Font->PutBold(0);
	}
	catch(...){}

}


void cls_main::f_keyboard()
{
	try
	{
		if(iHotk==-1)
		{
			shTS = name_sheet((_bstr_t)"shTS");
			psTS = xl->Sheets->GetItem(shTS);
			iHotk = _wtoi(getVCell(psTS,L"TS_HOTK"));
		}

		// Nếu trạng thái khác 1 (=0) thì sẽ không sử dụng phím tắt
		if(iHotk!=1) return;

		int _icol = pSheet->Application->ActiveCell->Column;
		if(_icol!=1) return;

		// Kiểm tra sheet activate
		_bstr_t cd = pSheet->CodeName;
		CString szval = (LPCTSTR)cd;
		szval = szval.Left(8);
		CString szAdd[13] = {
			L"shHSNTVLLAYM",L"shHSNTVLNBVL",L"shHSNTVLNVL",L"shHSNTVLYVL",
			L"shHSNTCVLMTN",L"shHSNTCVNBCV",L"shHSNTCVNBCV1",L"shHSNTCVNBCV2",
			L"shHSNTCVNTCV",L"shHSNTCVYNCV",L"shHSNTGDNTBGD1",L"shHSNTGDNTGD1",L"shHSNTGDYCNT1"};

		int iChek=0;
		for (int i = 0; i < 13; i++)
		{
			if(szAdd[i].GetLength()>8 && szAdd[i].Left(8)==szval)
			{
				iChek=1;
				break;
			}
		}

		if(iChek==0) return;

		int istyle=0;
		int _irow = pSheet->Application->ActiveCell->Row;
		RangePtr pRg1 = pSheet->Cells;
		RangePtr PRS = pRg1->GetItem(_irow, _icol);

		CString sKey = pRg1->GetItem(_irow,_icol);
		sKey.MakeLower();
		pRg1->PutItem(_irow,_icol,L"");

		int jSelect=0;
		int iLen = sKey.GetLength();
		CString sk1 = sKey.Left(1);
		if(sk1==L"z")
		{
			// Xóa định dạng
			PRS->EntireRow->ClearFormats();
		}
		else if(sk1==L"b")
		{
			// Chữ đậm
			int _pos = sKey.Find(L";");
			if(_pos>0)
			{
				sKey = sKey.Right(iLen-2);
				if(_wtoi(sKey)>0) PRS = pRg1->GetItem(_irow, _wtoi(sKey));
				else
				{
					szval.Format(L"%s%d",sKey,_irow);
					PRS = pRg1->GetRange((_bstr_t)szval,(_bstr_t)szval);
				}

				istyle = PRS->Font->GetBold();
				if(istyle==0) PRS->Font->PutBold(1);
				else PRS->Font->PutBold(0);
			}
			else
			{
				istyle = PRS->Font->GetBold();
				if(istyle==0) PRS->EntireRow->Font->PutBold(1);
				else PRS->EntireRow->Font->PutBold(0);
			}		
		}
		else if(sk1==L"i")
		{
			// Chữ nghiêng
			int _pos = sKey.Find(L";");
			if(_pos>0)
			{
				sKey = sKey.Right(iLen-2);
				if(_wtoi(sKey)>0) PRS = pRg1->GetItem(_irow, _wtoi(sKey));
				else
				{
					szval.Format(L"%s%d",sKey,_irow);
					PRS = pRg1->GetRange((_bstr_t)szval,(_bstr_t)szval);
				}

				istyle = PRS->Font->GetItalic();
				if(istyle==0) PRS->Font->PutItalic(1);
				else PRS->Font->PutItalic(0);
			}
			else
			{
				istyle = PRS->Font->GetItalic();
				if(istyle==0) PRS->EntireRow->Font->PutItalic(1);
				else PRS->EntireRow->Font->PutItalic(0);
			}			
		}
		else if(sk1==L"w" || sk1==L"ư" || sk1==L"Ư")
		{
			// Wrap text
			int _pos = sKey.Find(L";");
			if(_pos>0)
			{
				sKey = sKey.Right(iLen-2);
				if(_wtoi(sKey)>0) PRS = pRg1->GetItem(_irow, _wtoi(sKey));
				else
				{
					szval.Format(L"%s%d",sKey,_irow);
					PRS = pRg1->GetRange((_bstr_t)szval,(_bstr_t)szval);
				}

				istyle = PRS->GetWrapText();
				if(istyle==0) PRS->PutWrapText(1);
				else PRS->PutWrapText(0);
			}
			else
			{
				istyle = PRS->GetWrapText();
				if(istyle==0) PRS->EntireRow->PutWrapText(1);
				else PRS->EntireRow->PutWrapText(0);
			}			
		}
		else if(sk1==L"u")
		{
			if(sKey==L"uh")
			{
				// Unhide
				// Xác định vùng được lựa chọn
				RangePtr PRc = pSheet->Application->Selection;
				_bstr_t _bsArr = PRc->GetAddress(1,1,XlReferenceStyle::xlR1C1);
				int _nSelection = _fGetSelection((LPCTSTR)_bsArr);
				if(_nSelection>0)
				{
					int jbd = _irow, jkt= _irow;
					int _pos = _arrSLS[1].Find(L"-");
					if(_pos>0)				
					{
						// Có nhiều lựa chọn
						jbd = _wtoi(_arrSLS[1].Left(_pos));
						jkt = _wtoi(_arrSLS[1].Right(_arrSLS[1].GetLength()-_pos-1));
					}

					PRS = pSheet->GetRange(pRg1->GetItem(jbd,1),pRg1->GetItem(jkt,1));
					PRS->EntireRow->PutHidden(false);
				}
			}
			else
			{
				// Gạch chân
				int _pos = sKey.Find(L";");
				if(_pos>0)
				{
					sKey = sKey.Right(iLen-2);
					if(_wtoi(sKey)>0) PRS = pRg1->GetItem(_irow, _wtoi(sKey));
					else
					{
						szval.Format(L"%s%d",sKey,_irow);
						PRS = pRg1->GetRange((_bstr_t)szval,(_bstr_t)szval);
					}

					istyle = PRS->Font->GetUnderline();
					if(istyle==2) PRS->Font->PutUnderline(xlUnderlineStyleNone);
					else PRS->Font->PutUnderline(xlUnderlineStyleSingle);
				}
				else
				{
					istyle = PRS->Font->GetUnderline();
					if(istyle==2) PRS->EntireRow->Font->PutUnderline(xlUnderlineStyleNone);
					else PRS->EntireRow->Font->PutUnderline(xlUnderlineStyleSingle);
				}
			}			
		}
		else if(sk1==L"c")
		{
			// Chèn dòng
			int num=1;
			if(iLen>1) num = _wtoi(sKey.Right(iLen-1));
			if(num>=1)
			{
				PRS = pSheet->GetRange(pRg1->GetItem(_irow,1),pRg1->GetItem(_irow+num-1,1));
				PRS->EntireRow->Insert(xlShiftDown);
			}
		}
		else if(sk1==L"x")
		{
			// Xóa dòng
			int num=1;
			if(iLen>1) num = _wtoi(sKey.Right(iLen-1));
			if(num>=1)
			{
				PRS = pSheet->GetRange(pRg1->GetItem(_irow,1),pRg1->GetItem(_irow+num-1,1));
				PRS->EntireRow->Delete(xlShiftUp);
			}
		}
		else if(sk1==L"h")
		{
			// Hide
			// Xác định vùng được lựa chọn
			RangePtr PRc = pSheet->Application->Selection;
			_bstr_t _bsArr = PRc->GetAddress(1,1,XlReferenceStyle::xlR1C1);
			int _nSelection = _fGetSelection((LPCTSTR)_bsArr);
			if(_nSelection>0)
			{
				int jbd = _irow, jkt= _irow;
				int _pos = _arrSLS[1].Find(L"-");
				if(_pos>0)				
				{
					// Có nhiều lựa chọn
					jbd = _wtoi(_arrSLS[1].Left(_pos));
					jkt = _wtoi(_arrSLS[1].Right(_arrSLS[1].GetLength()-_pos-1));
				}

				PRS = pSheet->GetRange(pRg1->GetItem(jbd,1),pRg1->GetItem(jkt,1));
				PRS->EntireRow->PutHidden(true);
			}
		}
		else if(sk1==L"t")
		{
			// Tab
			int iktr=0;
			for (int i = 2; i < 100; i++)
			{
				szval = GIOR(_irow,i,pRg1,L"");
				szval.Trim();
				if(szval!=L"")
				{
					iktr=i;
					break;
				}
			}

			if(iktr>0)
			{
				int num=1;
				if(iLen>1) num = _wtoi(sKey.Right(iLen-1));
				if(num>=1)
				{
					CString sgan = L"'      ";	// 6 khoảng trống
					if(szval.Left(1)==L"'") szval = szval.Right(szval.GetLength()-1);					
					for (int i = 1; i <= num; i++) szval = sgan + szval;
					pRg1->PutItem(_irow,iktr,(_bstr_t)szval);
				}
			}
		}
		else if(sk1==L"+" || sk1==L"0" || _wtoi(sk1)>0)
		{
			// Thay đổi độ cao
			// Xác định vùng được lựa chọn
			RangePtr PRc = pSheet->Application->Selection;
			_bstr_t _bsArr = PRc->GetAddress(1,1,XlReferenceStyle::xlR1C1);
			int _nSelection = _fGetSelection((LPCTSTR)_bsArr);
			if(_nSelection>0)
			{
				sKey.Replace(L"+",L"");
				int iHg = 18;
				int num = _wtoi(sKey);
				int _pos = sKey.Find(L";");
				if(_pos>0)
				{
					iHg = _wtoi(sKey.Right(sKey.GetLength()-_pos-1)); 
					num = _wtoi(sKey.Left(_pos));
				}

				int jbd = _irow, jkt= _irow;
				_pos = _arrSLS[1].Find(L"-");
				if(_pos>0)				
				{
					// Có nhiều lựa chọn
					jbd = _wtoi(_arrSLS[1].Left(_pos));
					jkt = _wtoi(_arrSLS[1].Right(_arrSLS[1].GetLength()-_pos-1));
				}

				if(iHg>=0 && iHg<410 && num>=0 && num<=100)
				{
					PRS = pSheet->GetRange(pRg1->GetItem(jbd,1),pRg1->GetItem(jkt,1));
					PRS->EntireRow->PutRowHeight((int)(num*iHg));
				}				
			}
		}

		// Vị trí activate
		if(_irow>1)
		{
			RangePtr pRg2 = pRg1->GetItem(_irow-1, _icol);
			pRg2->Select();
		}
	}
	catch(...){_s(L"Lỗi nhập ký tự");}
}


void cls_main::f_luu_noidung()
{
	try
	{
		// Định nghĩa các sheet
		shNDNK = name_sheet("shMauNKY4");
		psNDNK=xl->Sheets->GetItem(shNDNK);
		prNDNK = psNDNK->Cells;

		int jc1 = getRow(psNDNK,"NKBS_1");		
		int jr1 = getColumn(psNDNK,"NKBS_1");
		int ir = getRow(psNDNK,"HK_NGAY");
		int ic = getColumn(psNDNK,"HK_NGAY");
		CString sNgay = prNDNK->GetItem(ir,ic);

////////  LƯU NỘI DUNG NHẬT KÝ (19.01.2016) /////////////////////////////////////

//////// 1 - ĐIỀU KIỆN THỜI TIẾT VÀ NHIỆT ĐỘ
		CString szval=L"",sc1=L"",sc2=L"";
		for (int i = 1; i <= 4; i++)
		{
			// Thời tiết
			szval = GIOR(jc1+2,jr1+i,prNDNK,L"");
			if(szval==L"--") szval=L"";
			sc1+=szval;
			sc1+=L";";

			// Nhiệt độ
			szval = GIOR(jc1+3,jr1+i,prNDNK,L"");
			if(szval==L"--") szval=L"";
			sc2 +=szval;
			sc2 +=L";";
		}

		sc1 = ReplateChamphay(sc1);
		sc2 = ReplateChamphay(sc2);

//////// 4 - CÔNG TÁC VỆ SINH
		CString sVSinh[3]={L"Tốt",L"Bình thường",L"Kém"};
		sVSinh[0] = GIOR(1,jr1+16,prNDNK,L"Tốt");
		sVSinh[1] = GIOR(2,jr1+16,prNDNK,L"Bình thường");
		sVSinh[2] = GIOR(3,jr1+16,prNDNK,L"Kém");

		// sc4=1 tốt	=2 bình thường	=3 kém	=0 chưa tích chọn
		int jc4 = getRow(psNDNK,"NKBS_4");
		szval = GIOR(jc4+1,jr1+10,prNDNK,L"");
		CString sc4 = L"";
		if(szval==L"TRUE" || szval==L"-1") sc4=sVSinh[0];
		else
		{
			szval = GIOR(jc4+1,jr1+11,prNDNK,L"");
			if(szval==L"TRUE" || szval==L"-1") sc4=sVSinh[1];
			else
			{
				szval = GIOR(jc4+1,jr1+12,prNDNK,L"");
				if(szval==L"TRUE" || szval==L"-1") sc4=sVSinh[2];
			}
		}

//////// 5 - CÔNG TÁC AN TOÀN LAO ĐỘNG
		CString sATLD[3]={L"Tốt",L"Bình thường",L"Kém"};
		sATLD[0] = GIOR(1,jr1+17,prNDNK,L"Tốt");
		sATLD[1] = GIOR(2,jr1+17,prNDNK,L"Bình thường");
		sATLD[2] = GIOR(3,jr1+17,prNDNK,L"Kém");

		// sc5=1 tốt	=2 bình thường	=3 kém	=0 chưa tích chọn
		int jc5 = getRow(psNDNK,"NKBS_5");
		szval = GIOR(jc5+1,jr1+10,prNDNK,L"");
		CString sc5 = L"";
		if(szval==L"TRUE" || szval==L"-1") sc5=sATLD[0];
		else
		{
			szval = GIOR(jc5+1,jr1+11,prNDNK,L"");
			if(szval==L"TRUE" || szval==L"-1") sc5=sATLD[1];
			else
			{
				szval = GIOR(jc5+1,jr1+12,prNDNK,L"");
				if(szval==L"TRUE" || szval==L"-1") sc5=sATLD[2];
			}
		}

// ----------------------> Thêm mới hoặc cập nhật vào tbNdung
		CString sDem=L"";
		SqlString.Format(L"SELECT COUNT(*) AS iDem FROM tbNdung WHERE ngayghi LIKE '%s';",sNgay);
		
		// Kiểm tra sNgay đã tồn tại trong CSDL chưa?
		if(getPathNHKY()==L"") return;
		if(connectDsn(pathNK)==-1) return;
		ObjConn.OpenAccessDB(pathNK, L"", L"");
		CRecordset* Recset = ObjConn.Execute(SqlString);
		Recset->GetFieldValue(L"iDem",sDem);
		Recset->Close();
		int iUpdt = _wtoi(sDem);

		CString sboqua= GIOR(ir+3,ic,prNDNK,L"");
		sboqua.Trim();
		if(sboqua!=L"") sboqua=L"1";
		
		if(iUpdt>0)
		{
			sDem.Format(L"Nội dung ngày %s đã tồn tại. Bạn có muốn cập nhật?",sNgay);
			if(_yn(sDem,0)!=6)
			{
				ObjConn.CloseDatabase();
				return;
			}

			SqlString.Format(
				L"UPDATE tbNdung SET dkthoitiet='%s', nhietdo='%s', ctvsinh='%s', ctatldong='%s',boqua='%s' "
				L"WHERE ngayghi LIKE '%s';",sc1,sc2,sc4,sc5,sboqua,sNgay);
			ObjConn.ExecuteRB(SqlString);
		}
		else
		{
			SqlString.Format(
				L"INSERT INTO tbNdung (ngayghi,dkthoitiet,nhietdo,ctvsinh,ctatldong,boqua) "
				L"VALUES ('%s','%s','%s','%s','%s','%s');",sNgay,sc1,sc2,sc4,sc5,sboqua);
			ObjConn.ExecuteRB(SqlString);
		}

//////// LƯU CÁC NỘI DUNG CHÍNH
		int jc32 = getRow(psNDNK,"NKBS_32");
		int jc33 = getRow(psNDNK,"NKBS_33");
		int jc34 = getRow(psNDNK,"NKBS_34");
		int jc35 = getRow(psNDNK,"NKBS_35");

		int jc6 = getRow(psNDNK,"NKBS_6");
		int jc7 = getRow(psNDNK,"NKBS_7");
		int jc8 = getRow(psNDNK,"NKBS_8");
		int jc9 = getRow(psNDNK,"NKBS_9");

		// Biện pháp thi công
		f_SQL_AddNdung(sNgay,jc32,jc33,jr1,1,L"tbBptc",12,iUpdt);

		// Tình trạng thực tế
		f_SQL_AddNdung(sNgay,jc33,jc34,jr1,1,L"tbTTTTe",13,iUpdt);

		// Chất lượng thi công
		f_SQL_AddNdung(sNgay,jc34,jc35,jr1,1,L"tbCltc",14,iUpdt);

		// Tiến độ thực hiện
		f_SQL_AddNdung(sNgay,jc35,jc4,jr1,1,L"tbTdth",15,iUpdt);

		// Nhận xét của cán bộ giám sát thi công
		f_SQL_AddNdung(sNgay,jc6,jc7,jr1,1,L"tbNxGstc",0,iUpdt);

		// Nhận xét của chủ đầu tư
		f_SQL_AddNdung(sNgay,jc7,jc8,jr1,1,L"tbNxCdt",0,iUpdt);

		// Ý kiến của nhà thầu thi công
		f_SQL_AddNdung(sNgay,jc8,jc9-1,jr1,1,L"tbYkNttc",0,iUpdt);

///////////// Bổ sung 23.06.2017 --> Lưu nhân công và máy thi công --------------------------------------------
		CString szv=L"";		
		int jc2 = getRow(psNDNK,"NKBS_2");
		int jc3 = getRow(psNDNK,"NKBS_3");	
		int stt=0;
		for (int i = jc2+2; i < jc3; i++)
		{
			szv = GIOR(i,jr1,prNDNK,L"");
			if(szv==L"STT")
			{
				stt=i;	// vị trí chứa STT nhân công
				break;
			}
		}

		// Lưu dữ liệu nhân công
		if(stt>0) f_SQL_AddNdung(sNgay,jc2+1,stt,jr1,1,L"tbNcong",0,iUpdt);

		// Xác định vị trí tiêu đề thiết bị thi công
		int pos=-1;
		int jc22 = 0;
		if(stt==0) stt = jc2+1;
		for (int i = stt+1; i < jc3; i++)
		{
			szv = GIOR(i,jr1,prNDNK,L"");
			pos = szv.Find(L"công");
			if(pos>0)
			{
				jc22=i;		// vị trí chứa tiêu đề thiết bị thi công
				break;
			}
		}

		// Lưu chi tiết danh sách nhân công
		if(jc22>0) f_SQL_Add_NC_MTC(sNgay,stt,jc22,jr1,1,L"tbSTTNc");
		
		stt=0;
		for (int i = jc22+1; i < jc3; i++)
		{
			szv = GIOR(i,jr1,prNDNK,L"");
			if(szv==L"STT")
			{
				stt=i;
				break;
			}
		}

		if(stt>0)
		{
			f_SQL_AddNdung(sNgay,jc22,stt,jr1,1,L"tbMtcong",0,iUpdt);		// Lưu dữ liệu máy thi công
			f_SQL_Add_NC_MTC(sNgay,stt,jc3,jr1,1,L"tbSTTMtc");				// Lưu chi tiết danh sách máy thi công
		}

/////////////////////////////////////////////////////////////////////

		_s(L"Cập nhật nội dung thành công!");
		ObjConn.CloseDatabase();
	
	}
	catch(...){_s(L"[3] Lỗi lưu dữ liệu nhật ký");}
}


void cls_main::f_SQL_Add_NC_MTC(CString sNgayNK,int num1,int num2,int cot,int cong,CString frTable)
{
	try
	{
		int dem=0;
		CString sLLu[5]={L"",L"",L"",L"",L""};
		CString szval=L"";

		// Xóa dữ liệu cũ
		SqlString.Format(L"DELETE FROM %s WHERE ngayghi LIKE '%s';",frTable,sNgayNK);
		ObjConn.ExecuteRB(SqlString);

		if(num2>num1+cong)
		for (int i = num1+cong; i < num2; i++)
		{
			szval = GIOR(i,cot+1,prNDNK,L"");
			if(szval!=L"")
			{
				dem++;
				if(dem<=9) sLLu[0].Format(L"G0%d",dem);
				else sLLu[0].Format(L"G%d",dem);

				for (int j = 1; j <= 4; j++) sLLu[j] = GIOR(i,cot+j,prNDNK,L"");
				SqlString.Format(L"INSERT INTO %s (ngayghi,maso,noidung,ftr1,ftr2,ftr3) "
					L"VALUES ('%s','%s','%s','%s','%s','%s');",
						frTable,sNgayNK,sLLu[0],sLLu[1],sLLu[2],sLLu[3],sLLu[4]);
				ObjConn.ExecuteRB(SqlString);
			}
		}
	}
	catch(...){}
}

// Hàm lưu nội dung nhật ký
// sNgayNK= ngày lưu nhật ký
// num1,num2 = dòng bắt đầu và kết thúc lưu; gt = số dòng được cộng thêm để bắt đầu tính vị trí lưu
// frTable = tên bảng truy vấn;
void cls_main::f_SQL_AddNdung(CString sNgayNK,int num1,int num2,int cot,int cong,CString frTable,int cotcheck,int iUpd)
{
	try
	{
		CString szval=L"", sLuuND=L"",sLuucheck=L"";

		// Xác định checkbox (nếu có)
		if(cotcheck>0)
		{
			CString sCheck[3]={L"Tốt",L"Bình thường",L"Kém"};
			sCheck[0] = GIOR(1,cot+cotcheck,prNDNK,L"Tốt");
			sCheck[1] = GIOR(2,cot+cotcheck,prNDNK,L"Bình thường");
			sCheck[2] = GIOR(3,cot+cotcheck,prNDNK,L"Kém");

			szval = GIOR(num1+1,cot+10,prNDNK,L"");
			if(szval==L"TRUE" || szval==L"-1") sLuucheck=sCheck[0];
			else
			{
				szval = GIOR(num1+1,cot+11,prNDNK,L"");
				if(szval==L"TRUE" || szval==L"-1") sLuucheck=sCheck[1];
				else
				{
					szval = GIOR(num1+1,cot+12,prNDNK,L"");
					if(szval==L"TRUE" || szval==L"-1") sLuucheck=sCheck[2];
				}
			}
		}		

		// Xác định nội dung (nếu có)
		if(num2>num1+cong)
		for (int i = num1+cong; i < num2; i++)
		{
			szval = GIOR(i,cot,prNDNK,L"");
			sLuuND+=szval;
			sLuuND+=L";";
		}

		sLuuND = ReplateChamphay(sLuuND);

		// Kiểm tra có dữ liệu chưa? Chưa có thì thêm mới / ngược lại  thì update
		if(iUpd>0) SqlString.Format(L"UPDATE %s SET noidung='%s', checknew='%s' WHERE ngayghi LIKE '%s';",frTable,sLuuND,sLuucheck,sNgayNK);
		else SqlString.Format(L"INSERT INTO %s (ngayghi,noidung,checknew) VALUES ('%s','%s','%s');",frTable,sNgayNK,sLuuND,sLuucheck);
		ObjConn.ExecuteRB(SqlString);
	}
	catch(...)
	{
		CString szval=L"";
		szval.Format(L"Lỗi truy vấn nội dung bảng %s",frTable);
		_s(szval);
	}
}


void cls_main::f_AddImage_nky()
{
	try
	{
		CFileDialog dlgOpenFile(TRUE);
		// add OFN_ALLOWMULTISELECT flag
		dlgOpenFile.GetOFN().Flags |= OFN_ALLOWMULTISELECT;
 
		// Thiết lập một bộ đệm để giữ ít nhất 100 đầy đủ đường dẫn và tên tập tin
		const int nBufferSize = 100 * (MAX_PATH + 1) + 1;
		CString strBuffer=L"";
		LPTSTR pBuffer = strBuffer.GetBufferSetLength(nBufferSize);
		dlgOpenFile.GetOFN().lpstrFile = pBuffer;
		dlgOpenFile.GetOFN().nMaxFile = nBufferSize;
 
		if(IDOK == dlgOpenFile.DoModal())
		{
			// Định nghĩa các sheet
			shNDNK = name_sheet("shMauNKY4");
			psNDNK=xl->Sheets->GetItem(shNDNK);
			prNDNK = psNDNK->Cells;

			int ir = getRow(psNDNK,"HK_NGAY");
			int ic = getColumn(psNDNK,"HK_NGAY");
			CString sNgay = prNDNK->GetItem(ir,ic);
		

			int xde = FindComment(psNDNK,1,"END2");
			int vtri=0,slg=0,dem=0;
			CString val=L"",_spath=L"";

			for (int i = xde+1; i <= xde+100; i++)
			{
				_spath = prNDNK->GetItem(i,1);
				_spath.Trim();
				if(_spath!=L"")
				{
					slg++;
					dem=0;
				}
				else dem++;

				if(dem==5)
				{
					vtri=i-4;
					break;					
				}
			}

			if(slg>=100) 
				MessageBox(
					NULL,L"Số lượng ảnh lưu trữ đang được giới hạn là 100 ảnh.",
					L"Cảnh báo",MB_OK | MB_ICONWARNING);

			if(slg>0)
			{
				val.Format(L"Đã tồn tại dữ liệu ảnh ngày %s.\nBạn có muốn thêm tiếp vào ngay dưới?"
					L"\nChọn YES để thêm tiếp - NO để thay thế dữ liệu ảnh cũ",sNgay);

				int N = _yn(val,1);
				if(N==6 || N==7)
				{
					if(N==7)
					{
						// Xóa dữ liệu cũ
						PRS = psNDNK->GetRange(prNDNK->GetItem(xde+1,1),prNDNK->GetItem(vtri,1));
						PRS->EntireRow->Delete(xlShiftUp);
						vtri=xde+1;
						slg=0;
					}
				}
				else
				{
					// release buffer
					strBuffer.ReleaseBuffer();
					return;
				}
			}

			// Xác định các file được lựa chọn
			int slg2=0;
			POSITION _pos = dlgOpenFile.GetStartPosition();
			while(NULL != _pos)
			{
				_spath = dlgOpenFile.GetNextPathName(_pos);
				val = _spath.Right(4);
				val.MakeLower();
				if(val==L".jpg")
				{
					prNDNK->PutItem(vtri+slg2,1,(_bstr_t)_spath);
					slg2++;
				}			

				if(slg2==100) break;
			}

			if(slg2>0)
			{
				_s(L"Thêm mới ảnh thành công!\nBạn có thể mô tả về hình ảnh đó tại cột B tương ứng!");
				
				if(slg+slg2>=100) 
				MessageBox(NULL,L"Số lượng ảnh lưu trữ đang được giới hạn là 100 ảnh.",
					L"Cảnh báo",MB_OK | MB_ICONWARNING);

				//xl->ActiveWindow->PutScrollRow(vtri);
				//PRS = prNDNK->GetItem(vtri,1);
				//PRS->EntireRow->Select();
			}
			else _s(L"Không có hình ảnh nào được thêm mới. Vui lòng kiểm tra lại định dạng tệp tin sử dụng (*.JPG)");			
		}

		// release buffer
		strBuffer.ReleaseBuffer();
	}
	catch(CException* e)
	{
		e->ReportError();
		e->Delete();
	}
}


// Tạo bảng dự toán Man Month
// iCopy = dòng bắt đầu copy
// iColE = cột cuối cùng bắt đầu copy
int cls_main::f_Create_table(_WorksheetPtr pSh1, int iCopy, int iColE)
{
	try
	{
		_bstr_t bsCPTV = wsTHCPtv->Name;
		_bstr_t bs1 = pSh1->Name;
		RangePtr prCPTV = wsTHCPtv->Cells;
		RangePtr pRg1 = pSh1->Cells;

		// Kiểm tra xem Man Month thứ mấy?
		CString szval = GIOR(iCopy,1,pRg1,"");
		szval.Trim();

		int iBdau=1,iEnd=1,iVtri=iCopy;
		if(szval!=L"")
		{
			// Tạo bảng tính mới ở dưới
			// Xác định vị trí END dưới cùng
			int dem=0;
			while (iEnd>0)
			{
				iEnd = FindCEnd(pSh1,2,"END",iEnd+1);
				if(iEnd>0)
				{
					if(dem==0) iBdau = iEnd;
					iVtri=iEnd;
				}

				dem++;
			}

			// Xác định được vị trí sẽ đổ dữ liệu
			iVtri+=3;

			// Copy - paste
			PRS = pSh1->GetRange(pRg1->GetItem(iCopy,2),pRg1->GetItem(iBdau,iColE));
			PRS->Copy(pRg1->GetItem(iVtri,2));
			xl->PutCutCopyMode(XlCutCopyMode(false));
		}

		// STT
		szval.Format(L"='%s'!%s",(LPCTSTR)bsCPTV,GAR(iMMRow,2,prCPTV,3));
		pRg1->PutItem(iVtri,1,(_bstr_t)szval);

		// Tên nội dung chi phí
		szval.Format(L"=\"Dự toán: \"&'%s'!%s",(LPCTSTR)bsCPTV,GAR(iMMRow,3,prCPTV,3));
		pRg1->PutItem(iVtri,2,(_bstr_t)szval);	

		return iVtri;
	}
	catch(...){_s(L"Lỗi tạo bảng.");}
}


void cls_main::f_Link_manmonth(_WorksheetPtr pSh1,_WorksheetPtr pSh2,_WorksheetPtr pSh3)
{
	try
	{
		_bstr_t bsCPTV = wsTHCPtv->Name;	// Tổng hợp CPTV
		_bstr_t bs1 = pSh1->Name;	// Dự toán MM
		_bstr_t bs2 = pSh2->Name;	// Chi phí chuyên gia
		_bstr_t bs3 = pSh3->Name;	// Chi phí khác

		RangePtr prCPTV = wsTHCPtv->Cells;
		RangePtr pRg1 = pSh1->Cells;
		RangePtr pRg2 = pSh2->Cells;
		RangePtr pRg3 = pSh3->Cells;

		CString szval=L"";
		int iVtri = _mmvt[1];

		// Liên kết MM với tổng hợp cptv (ps1-wsTHCPtv)
		// Giá trị trước thuế
		szval.Format(L"=%s-%s",GAR(iMMRow,9,prCPTV,0),GAR(iMMRow,8,prCPTV,0));
		prCPTV->PutItem(iMMRow,7,(_bstr_t)szval);

		// Thuế giá trị gia tăng
		szval.Format(L"='%s'!%s",(LPCTSTR)bs1,GAR(iVtri+6,5,pRg1,0));
		prCPTV->PutItem(iMMRow,8,(_bstr_t)szval);

		// Giá trị sau thuế
		szval.Format(L"='%s'!%s-'%s'!%s",(LPCTSTR)bs1,GAR(iVtri+8,5,pRg1,0),(LPCTSTR)bs1,GAR(iVtri+7,5,pRg1,0));
		prCPTV->PutItem(iMMRow,9,(_bstr_t)szval);


////////////////////////////////////////
		// Liên kết MM với chi phí chuyên gia (ps1-ps2)
		int iEnd = FindCEnd(pSh2,2,"END",_mmvt[2]+1);
		szval.Format(L"='%s'!%s",(LPCTSTR)bs2,GAR(iEnd-1,8,pRg2,0));
		pRg1->PutItem(iVtri+2,5,(_bstr_t)szval);

		// Liên kết MM với chi phí khác (ps1-ps3)
		iEnd = FindCEnd(pSh3,2,"END",_mmvt[4]+1);
		szval.Format(L"='%s'!%s",(LPCTSTR)bs3,GAR(iEnd-1,5,pRg3,0));
		pRg1->PutItem(iVtri+4,5,(_bstr_t)szval);
	}
	catch(...){}
}


void cls_main::f_Delete_table(_bstr_t num)
{
	try
	{
		_WorksheetPtr pSh1;
		RangePtr pRg1;
		CString szval=L"";
		_bstr_t sh1,bsx[5] = {"shDTMM","shCPCGia","shCPTLuong","shCPKTV","shTDCViec"};
		int iGbd[5] = {4,4,4,4,5};
		int iBdau=0,iEnd=0,iChk=0;

		for (int i = 0; i < 5; i++)
		{
			sh1 = name_sheet(bsx[i]);
			pSh1 = xl->Sheets->GetItem(sh1);
			pRg1 = pSh1->Cells;

			iBdau = FindValue(pSh1,1,1,0,num);
			if(iBdau>0)			
			{
				iEnd = FindCEnd(pSh1,2,"END",iBdau+1);
				if(iEnd>0)
				{
					// Kiểm tra nếu vị trí muốn xóa là đầu tiên và duy nhất thì không xóa, chỉ ClearContents dòng đầu tiên
					iChk = FindCEnd(pSh1,2,"END",iEnd+1);
					if(iBdau==iGbd[i] && iChk==0)
					{
						PRS = pSh1->GetRange(pRg1->GetItem(iBdau,1),pRg1->GetItem(iBdau,2));
						PRS->ClearContents();
					}
					else
					{
						for (int j = iEnd+1; j < iEnd+100; j++)
						{
							szval = GIOR(j,2,pRg1,L"");
							if(szval!=L"")
							{
								iEnd = j-1;
								break;
							}
						}

						PRS = pSh1->GetRange(pRg1->GetItem(iBdau,1),pRg1->GetItem(iEnd,1));
						PRS->EntireRow->Delete(xlShiftUp);
					}					
				}	
			}
		}
	}
	catch(...){_s(L"Lỗi xóa bảng tính.");}
}


void cls_main::f_select_gtri()
{
	try
	{
		_WorksheetPtr pSh1 = pWb->ActiveSheet;
		RangePtr pRg1 = pSh1->Cells;
		_bstr_t _bsh1 = pSh1->CodeName;
		if(_bsh1!=(_bstr_t)L"shDTXD" && _bsh1!=(_bstr_t)L"shTHXD") return;

		slvchon=0,ivtdchuyen=0;

		// Xác định vùng được lựa chọn
		int iSetHM=0,cotid=1;
		PRS = pSh1->Application->Selection;
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
					slvchon++;
					iVtchon[slvchon] = _wtoi(_arrSLS[i]);

					// Kiểm tra lựa chọn đó có phải là hạng mục hay không?
					if(_bsh1==(_bstr_t)L"shDTXD")
					{
						cotid = getColumn(pSh1,L"BIEUGIA_ID");
						CString sktr = GIOR(iVtchon[slvchon],cotid,pRg1,L"");
						if(sktr.Left(2)==L"HM") iSetHM=1;
						else if(sktr.Left(2)==L"GD") iSetHM=2;
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
						slvchon++;
						iVtchon[slvchon] = j;
					}
				}
			}

			if(slvchon==0) return;

			if(iSetHM==1 || iSetHM==2)
			{
				CString sTen=L"hạng mục";
				if(iSetHM==2) sTen=L"giai đoạn";
				CString szval = L"";
				szval.Format(L"Bạn có chắc chắn đưa toàn bộ dữ liệu của %s được chọn sang danh mục công việc?",sTen);
				if(_yn(szval,0)!=6) return;

				sTen=L"HM";
				if(iSetHM==2) sTen=L"GD";

				// Xác định các vị trí được chọn (iVtchon[0] = HM)
				iVtchon[0] = iVtchon[1];
				slvchon=0;
				CString shmc = L"";			
				int iEnd = FindComment(pSh1,cotid,"END");
				for (int i = iVtchon[0]+1; i < iEnd; i++)
				{
					shmc = GIOR(i,cotid,pRg1,L"");
					if(shmc.Left(2)==L"HM" || shmc.Left(2)==sTen) break;

					if(_wtoi(shmc)>0 || shmc.Left(2)==L"HM" || shmc.Left(2)==L"GD")
					{
						slvchon++;
						iVtchon[slvchon] = i;
					}					
				}
			}

			// Xác định vị trí sẽ di chuyển đến
			AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
			Dlg_move_dulieu _dlg;
			_dlg.DoModal();

			if(ivtdchuyen==0) return;

			// Di chuyển
			if(_bsh1==(_bstr_t)L"shDTXD") f_move_bieugia(pSh1,iSetHM);
			else f_move_bangvattu(pSh1,iSetHM);
		}
	}
	catch(...){_s(L"Lỗi xác định vùng chọn");}
}

// itype=1 HM - itype=2 GD
void cls_main::f_move_bieugia(_WorksheetPtr psh1,int itype)
{
	try
	{
		_bstr_t bsn = psh1->Name;
		RangePtr prg1 = psh1->Cells;
		CString szval = L"", sMSCV = L"", sLLink = L"";

		_bstr_t shBTC = name_sheet("shNhomTC");
		_WorksheetPtr psBTC = xl->Sheets->GetItem(shBTC);
		RangePtr prBTC = psBTC->GetCells();
		int iVisible = psBTC->GetVisible();
		if (iVisible!=-1) psBTC->PutVisible(XlSheetVisibility::xlSheetVisible);

		// Đổ dữ liệu sang sheet Bang tieu chuan
		int icSTT = getColumn(psBTC,L"NHTC_STT");
		int icTen = getColumn(psBTC,L"NHTC_TEN");
		int icMa = getColumn(psBTC,L"NHTC_MA");
		int icTC = getColumn(psBTC,L"NHTC_TC");
		int icGC = getColumn(psBTC,L"NHTC_GC");
		int irSTT = getRow(psBTC,L"NHTC_STT");

		shTS = name_sheet("shTS");
		psTS=xl->Sheets->GetItem(shTS);
		prTS=psTS->GetCells();
		jNhomTC = _wtoi(getVCell(psTS,L"TS_GRPTC"));

		// Xác định cột
		int icot1 = getColumn(psh1,L"BIEUGIA_ID");
		int icot3 = getColumn(psh1, L"BIEUGIA_MVUA");
		int icot4 = getColumn(psh1,L"BIEUGIA_MDM");		
		int icot5 = getColumn(psh1,L"BIEUGIA_MHDG");
		int icot6 = getColumn(psh1,L"BIEUGIA_TEN");
		int icot7 = getColumn(psh1,L"BIEUGIA_DV");
		int icot8 = getColumn(psh1,L"BIEUGIA_KL");		

		// Xác định vị trí cần di chuyển đến danh mục công việc
		shDMCV = name_sheet((_bstr_t)"shHSNTCV");
		psDMCV = xl->Sheets->GetItem(shDMCV);
		prDMCV = psDMCV->Cells;
		
		int iCdm1 = getColumn(psDMCV,L"DMBB_STT");
		int iCdm2 = getColumn(psDMCV,L"DMBB_MADM");
		int iCdm3 = getColumn(psDMCV,L"DMBB_MAVUA");
		int iCdm4 = getColumn(psDMCV,L"DMBB_MACV");
		int iCdm6 = getColumn(psDMCV,L"DMBB_MAHS");
		int iCdm7 = getColumn(psDMCV,L"DMBB_ND");
		int iCdm8 = getColumn(psDMCV,L"DMBB_DV");
		int iCdm9 = getColumn(psDMCV,L"DMBB_KL");
		int iCdm10 = getColumn(psDMCV,L"DMBB_TC");

		int iCdm27 = getColumn(psDMCV,L"DMBB_DOT");
		int iCdm28 = getColumn(psDMCV,L"DMBB_DTLMAU");		

		int num=0,iCtg[9],iCtm[6];
		iCtg[1] = getColumn(psDMCV,L"DMBB_TN_NGAY");
		iCtg[2] = getColumn(psDMCV,L"DMBB_NB_NGAY");
		iCtg[3] = getColumn(psDMCV,L"DMBB_PHIEUYC");
		iCtg[4] = getColumn(psDMCV,L"DMBB_AB_NGAY");
		iCtg[5] = getColumn(psDMCV,L"DMBB_HK_NGAYBD");
		iCtg[6] = getColumn(psDMCV,L"DMBB_HK_NGAYKT");
		iCtg[7] = getColumn(psDMCV,L"DMBB_DKDBT_NGAY");
		iCtg[8] = getColumn(psDMCV,L"DMBB_TDDBT_NGAY");

		iCtm[1] = getColumn(psDMCV,L"DMBB_TN_GIO");
		iCtm[2] = getColumn(psDMCV,L"DMBB_NB_GIO");
		iCtm[3] = getColumn(psDMCV,L"DMBB_AB_GIO");
		iCtm[4] = getColumn(psDMCV,L"DMBB_DKDBT_GIO");
		iCtm[5] = getColumn(psDMCV,L"DMBB_TDDBT_GIO");

		int _iCot[50];
		_iCot[6] = iCdm9;
		_iCot[30] = getColumn(psDMCV, L"DMBB_HK_TONGNGAY");
		_iCot[35] = getColumn(psDMCV, L"DMBB_HK_TNMAY");
		_iCot[36] = getColumn(psDMCV, L"DMBB_HK_DMNC");
		_iCot[37] = getColumn(psDMCV, L"DMBB_HK_HESO");
		_iCot[38] = getColumn(psDMCV, L"DMBB_HK_TONGCONG");
		_iCot[39] = getColumn(psDMCV, L"DMBB_HK_NCTN");

		int vt=ivtdchuyen;

		if(itype==1 || itype==2)
		{
			PRS = prDMCV->GetItem(ivtdchuyen,1);
			PRS->EntireRow->Insert(xlShiftDown);

			PRS = prDMCV->GetItem(ivtdchuyen,1);
			PRS->EntireRow->Font->PutBold(1);
			PRS->EntireRow->Font->PutColor(BLACK);

			szval.Format(L"='%s'!%s",(LPCTSTR)bsn,GAR(iVtchon[0],icot6,prg1,0));
			prDMCV->PutItem(ivtdchuyen,iCdm7,(_bstr_t)szval);

			if(itype==1)
			{
				PRS->EntireRow->Interior->PutColor(RGB(217,217,217));

				prDMCV->PutItem(ivtdchuyen,iCdm4,(_bstr_t)L"HM");

				szval.Format(L"=COUNTIF(%s:%s,\"HM\")",GAR(7,iCdm4,prDMCV,0),GAR(ivtdchuyen-1,iCdm4,prDMCV,0));
				prDMCV->PutItem(ivtdchuyen,iCdm6,(_bstr_t)szval);
				szval = GIOR(ivtdchuyen,iCdm6,prDMCV,L"");
				num = _wtoi(szval);
				num++;
				szval.Format(L"HM%d",num);
				prDMCV->PutItem(ivtdchuyen,iCdm6,(_bstr_t)szval);

				PRS = psDMCV->GetRange(prDMCV->GetItem(ivtdchuyen,1),prDMCV->GetItem(ivtdchuyen,iCtm[5]));
				PRS->Borders->GetItem(XlBordersIndex::xlInsideVertical)->PutLineStyle(xlNone);				
				PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->PutWeight(xlThin);
				PRS->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->PutWeight(xlThin);

				PRS = prDMCV->GetItem(ivtdchuyen,iCdm7);
				PRS->PutHorizontalAlignment(xlCenter);
			}
			else
			{
				PRS->EntireRow->Interior->PutColor(xlNone);

				prDMCV->PutItem(ivtdchuyen,iCdm4,(_bstr_t)L"GD");

				szval.Format(L"=COUNTIF(%s:%s,\"GD\")",GAR(7,iCdm4,prDMCV,0),GAR(ivtdchuyen-1,iCdm4,prDMCV,0));
				prDMCV->PutItem(ivtdchuyen,iCdm6,(_bstr_t)szval);
				szval = GIOR(ivtdchuyen,iCdm6,prDMCV,L"");
				num = _wtoi(szval);
				num++;
				if(num<10) szval.Format(L"'0%dGD",num);
				else szval.Format(L"'%dGD",num);
				prDMCV->PutItem(ivtdchuyen,iCdm6,(_bstr_t)szval);

				PRS = psDMCV->GetRange(prDMCV->GetItem(ivtdchuyen,1),prDMCV->GetItem(ivtdchuyen,iCtm[5]));
				PRS->Borders->GetItem(XlBordersIndex::xlInsideVertical)->PutWeight(xlThin);				

				PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->PutWeight(xlThin);
				PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->PutLineStyle(xlDot);

				PRS->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->PutWeight(xlThin);
				PRS->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->PutLineStyle(xlDot);

				PRS = prDMCV->GetItem(ivtdchuyen,iCdm7);
				PRS->PutHorizontalAlignment(xlLeft);
			}					

			ivtdchuyen++;
		}

		if(getPathQLCL()==L"") return;
		if(connectDsn(pathMDB)==-1) return;
		ObjConn.OpenAccessDB(pathMDB,L"",L"");
		CRecordset *RecsetTC, *RecsetND;

		if (ibgDMNC == 1) f_Danh_sachfile();
	
		int iChen=0;
		CString stemp=L"",szv[7]={L"",L"",L"",L"",L"",L"",L"" };
		for (int i = 1; i <= slvchon; i++)
		{
			szval.Format(L"Đang tiến hành di chuyển công việc. "
				L"Vui lòng đợi trong giây lát... (%d%s)", (int)(100 * i / slvchon), L"%");
			xl->PutStatusBar((_bstr_t)szval);

			// Kiểm tra vị trí lựa chọn có tồn tại công tác không?
			szv[0] = GIOR(iVtchon[i],icot1,prg1,L"");
			if(szv[0].Left(2)==L"HM" || szv[0].Left(2)==L"GD")
			{
				PRS = prDMCV->GetItem(ivtdchuyen,1);
				PRS->EntireRow->Insert(xlShiftDown);

				PRS = prDMCV->GetItem(ivtdchuyen,1);
				PRS->EntireRow->Font->PutBold(1);
				PRS->EntireRow->Font->PutColor(BLACK);

				szval.Format(L"='%s'!%s",(LPCTSTR)bsn,GAR(iVtchon[i],icot6,prg1,0));
				prDMCV->PutItem(ivtdchuyen,iCdm7,(_bstr_t)szval);

				if(szv[0].Left(2)==L"HM")
				{
					PRS->EntireRow->Interior->PutColor(RGB(217,217,217));				

					prDMCV->PutItem(ivtdchuyen,iCdm4,(_bstr_t)L"HM");

					szval.Format(L"=COUNTIF(%s:%s,\"HM\")",GAR(7,iCdm4,prDMCV,0),GAR(ivtdchuyen-1,iCdm4,prDMCV,0));
					prDMCV->PutItem(ivtdchuyen,iCdm6,(_bstr_t)szval);
					szval = GIOR(ivtdchuyen,iCdm6,prDMCV,L"");
					num = _wtoi(szval);
					num++;
					szval.Format(L"HM%d",num);
					prDMCV->PutItem(ivtdchuyen,iCdm6,(_bstr_t)szval);

					PRS = psDMCV->GetRange(prDMCV->GetItem(ivtdchuyen,1),prDMCV->GetItem(ivtdchuyen,iCtm[5]));
					PRS->Borders->GetItem(XlBordersIndex::xlInsideVertical)->PutLineStyle(xlNone);				
					PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->PutWeight(xlThin);
					PRS->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->PutWeight(xlThin);

					PRS = prDMCV->GetItem(ivtdchuyen,iCdm7);
					PRS->PutHorizontalAlignment(xlCenter);
				}
				else
				{
					PRS->EntireRow->Interior->PutColor(xlNone);

					prDMCV->PutItem(ivtdchuyen,iCdm4,(_bstr_t)L"GD");
					
					szval.Format(L"=COUNTIF(%s:%s,\"GD\")",GAR(7,iCdm4,prDMCV,0),GAR(ivtdchuyen-1,iCdm4,prDMCV,0));
					prDMCV->PutItem(ivtdchuyen,iCdm6,(_bstr_t)szval);
					szval = GIOR(ivtdchuyen,iCdm6,prDMCV,L"");
					num = _wtoi(szval);
					num++;
					if(num<10) szval.Format(L"'0%dGD",num);
					else szval.Format(L"'%dGD",num);
					prDMCV->PutItem(ivtdchuyen,iCdm6,(_bstr_t)szval);

					PRS = psDMCV->GetRange(prDMCV->GetItem(ivtdchuyen,1),prDMCV->GetItem(ivtdchuyen,iCtm[5]));
					PRS->Borders->GetItem(XlBordersIndex::xlInsideVertical)->PutWeight(xlThin);				

					PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->PutWeight(xlThin);
					PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->PutLineStyle(xlDot);

					PRS->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->PutWeight(xlThin);
					PRS->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->PutLineStyle(xlDot);

					PRS = prDMCV->GetItem(ivtdchuyen,iCdm7);
					PRS->PutHorizontalAlignment(xlLeft);
				}						

				ivtdchuyen++;

				continue;
			}

			szv[5] = GIOR(iVtchon[i], icot3, prg1, L"");		// Mã vữa
			szv[6] = GIOR(iVtchon[i], icot4, prg1, L"");		// Mã ĐM
			szv[1] = GIOR(iVtchon[i], icot5, prg1, L"");		// MHĐG

			szv[2].Format(L"='%s'!%s", (LPCTSTR)bsn, GAR(iVtchon[i], icot6, prg1, 0));
			szv[3].Format(L"='%s'!%s", (LPCTSTR)bsn, GAR(iVtchon[i], icot7, prg1, 0));
			szv[4].Format(L"='%s'!%s", (LPCTSTR)bsn, GAR(iVtchon[i], icot8, prg1, 0));

			if (_wtoi(szv[0]) > 0 && szv[1] != L"")
			{
				sMSCV = szv[1];					// Mã hiệu đơn giá
				sLLink = sFpathdulieu[1];		// Link connect

				slMaTC=0;
				if(ibgTC==1)
				{
					// Xác định tiêu chuẩn nếu có					
					SqlString.Format(L"SELECT *FROM tbMaCV_Tieuchuan where MaCV='%s';",szv[1]);
					RecsetTC = ObjConn.Execute(SqlString);
					//Fill TC

					while( !RecsetTC->IsEOF())
					{
						RecsetTC->GetFieldValue(L"MaTC",stemp);
						if(f_Check_tieuchuan(stemp)==0)
						{
							slMaTC++;
							sKTMaTC[slMaTC] = stemp;
							sKTTenTC[slMaTC] = stemp + L": Không tìm thấy tên tiêu chuẩn";

							SqlString.Format(L"SELECT *FROM tbTentieuchuan where MaTC='%s';",stemp);
							RecsetND = ObjConn.Execute(SqlString);
							while( !RecsetND->IsEOF())
							{
								RecsetND->GetFieldValue(L"TenTC",sKTTenTC[slMaTC]);
								break;
							}
							RecsetND->Close();
						}
						RecsetTC->MoveNext();
					}
					RecsetTC->Close();
				}

				if(jNhomTC==0) iChen = slMaTC;
				else iChen=1;

				if(iChen==0) iChen++;

				if(iChen>1)
				{
					// Có nhiều hơn 1 tiêu chuẩn
					PRS = psDMCV->GetRange(prDMCV->GetItem(ivtdchuyen,1),prDMCV->GetItem(ivtdchuyen+iChen-1,1));
					PRS->EntireRow->Insert(xlShiftDown);
				}
				else
				{
					// iChen=0 hoặc 1 đều chèn dòng
					PRS = prDMCV->GetItem(ivtdchuyen,1);
					PRS->EntireRow->Insert(xlShiftDown);
				}

				// Đổ dữ liệu
				prDMCV->PutItem(ivtdchuyen,iCdm2,(_bstr_t)szv[5]);	// Mã ĐM
				prDMCV->PutItem(ivtdchuyen,iCdm3,(_bstr_t)szv[6]);	// Mã vữa
				prDMCV->PutItem(ivtdchuyen,iCdm4,(_bstr_t)szv[1]);	// MHDG
				prDMCV->PutItem(ivtdchuyen,iCdm7,(_bstr_t)szv[2]);				
				prDMCV->PutItem(ivtdchuyen,iCdm8,(_bstr_t)szv[3]);				
				prDMCV->PutItem(ivtdchuyen,iCdm9,(_bstr_t)szv[4]);

				PRS = prDMCV->GetItem(ivtdchuyen,1);
				PRS->EntireRow->Font->PutBold(0);
				PRS->EntireRow->Font->PutItalic(0);

				int iktrNT=0;
				if(ibgTG==1 || ibgTM==1)
				{
					int vtcpy=0;
					// Xác định vị trí cần sao chép
					for (int j = ivtdchuyen-1; j >= 8; j--)
					{
						stemp = prDMCV->GetItem(j,iCdm1);
						if(_wtoi(stemp)>0)
						{
							vtcpy=j;
							break;
						}
					}

					// Sao chép thời gian
					if(vtcpy>0)
					{
						if(ibgTG==1)
						{
							for (int j = 1; j < 9; j++)
							{
								szval = GIOR(vtcpy,iCtg[j],prDMCV,L"");
								f_ktra_date(szval,prDMCV,ivtdchuyen,iCtg[j]);
							}
						}

						if(ibgTM==1)
						{
							for (int j = 1; j < 6; j++)
							{
								szval = GIOR(vtcpy,iCtm[j],prDMCV,L"");
								prDMCV->PutItem(ivtdchuyen,iCtm[j],(_bstr_t)szval);
							}
						}

						iktrNT=1;
					}

					if(iktrNT==0)
					{
						prDMCV->PutItem(ivtdchuyen,iCtg[2],(_bstr_t)L"=NOW()");
						CString val = GIOR(ivtdchuyen,iCtg[2],prDMCV,L"");
						int pos2= val.Find(L" ");
						if(pos2>0) val = val.Left(pos2);
						val.Trim();

						f_ktra_date(val,prDMCV,ivtdchuyen,iCtg[2]);

						val.Format(L"=%s",GAR(ivtdchuyen,iCtg[2],prDMCV,0));
						prDMCV->PutItem(ivtdchuyen,iCtg[3],(_bstr_t)val);

						val.Format(L"=%s+1",GAR(ivtdchuyen,iCtg[2],prDMCV,0));
						prDMCV->PutItem(ivtdchuyen,iCtg[4],(_bstr_t)val);

						val.Format(L"=%s",GAR(ivtdchuyen,iCtg[2],prDMCV,0));
						prDMCV->PutItem(ivtdchuyen,iCtg[5],(_bstr_t)val);

						val.Format(L"=%s",GAR(ivtdchuyen,iCtg[5],prDMCV,0));
						prDMCV->PutItem(ivtdchuyen,iCtg[6],(_bstr_t)val);
					}
				}

				// Đổ tiêu chuẩn
				if(slMaTC>0)
				{
					if(jNhomTC==0)
					{
						for (int j = 0; j < iChen; j++)
						{
							prDMCV->PutItem(ivtdchuyen + j, iCdm10, (_bstr_t)sKTTenTC[j + 1]);

							stemp = prDMCV->GetItem(ivtdchuyen + j, iCdm10 + 1);
							if (stemp == L"") prDMCV->PutItem(ivtdchuyen + j, iCdm10 + 1, (_bstr_t)L" ");
						}
					}
					else
					{
						// Đây là trường hợp group tiêu chuẩn
						// Kiểm tra xem nhóm tiêu chuẩn đã tồn tại bên sheet BTC chưa?
						int ivtr=1,iOk=0,icheck=0;
						int iEndBTC = prBTC->SpecialCells(xlCellTypeLastCell)->GetRow();
						for (int j = 1; j <= iEndBTC; j++)
						{
							szval = GIOR(j,icTen,prBTC,L"");
							if(szval!=L"")
							{
								szval = GIOR(j,icTC,prBTC,L"");
								if(szval==sKTTenTC[1])
								{
									ivtr=j;
									icheck=1;
									if(slMaTC>1)
									{
										for (int k = 2; k <= slMaTC; k++)
										{
											szval = GIOR(j+k-1,icTC,prBTC,L"");
											if(szval==sKTTenTC[k]) icheck++;
											else break;
										}
									}

									// Đã tồn tại tên nhóm tiêu chuẩn
									if(icheck==slMaTC)
									{							
										iOk=1;
										break;
									}
								}
							}
						}

						if(iOk==1)
						{
							szval.Format(L"='%s'!%s",(LPCTSTR)shBTC,GAR(ivtr,icTen,prBTC,0));
							prDMCV->PutItem(ivtdchuyen,iCdm10,(_bstr_t)szval);
						}
						else
						{
							// Chưa tồn tại --> Thêm mới vào đây
							ivtr=irSTT;
							for (int j = iEndBTC; j >= irSTT; j--)
							{
								szval = GIOR(j,icTC,prBTC,L"");
								szval.Trim();
								if(szval!=L"")
								{
									ivtr=j;
									break;
								}
							}

							PRS = psBTC->GetRange(prBTC->GetItem(ivtr+1,1),prBTC->GetItem(ivtr+1,icGC));
							PRS->Interior->PutColor(RGB(217,217,217));

							szval=L"";
							time_t now = time(0);
							tm *ltm = localtime(&now);
							szval.Format(L"%d%d",ltm->tm_min,ltm->tm_sec);
							int _nRand = _wtoi(szval);

							szval.Format(L"Nhóm TC%d%d%d",_nRand,rand()%10,rand()%10);
							prBTC->PutItem(ivtr+1,icTen,(_bstr_t)szval);

							szval.Format(L"=COUNTA(%s:%s)",GAR(irSTT+1,icTen,prBTC,3),GAR(ivtr+1,icTen,prBTC,0));
							prBTC->PutItem(ivtr+1,icSTT,(_bstr_t)szval);

							szval.Format(L"='%s'!%s",(LPCTSTR)shBTC,GAR(ivtr+1,icTen,prBTC,3));
							prDMCV->PutItem(ivtdchuyen,iCdm10,(_bstr_t)szval);

							for (int j = 1; j <= slMaTC; j++)
							{
								prBTC->PutItem(ivtr+j,icMa,(_bstr_t)sKTMaTC[j]);
								prBTC->PutItem(ivtr+j,icTC,(_bstr_t)sKTTenTC[j]);
							}
						}
					}
				}

/****************************************************************************************************/

				if (ibgDMNC == 1)
				{
					// Bổ sung 10.12.2019
					szfDMNhancong = L"", szfTNMay = L"";
					_getDinhmucNhancong(sMSCV, sLLink);

					if (szfDMNhancong == L"") szfDMNhancong = L"0";
					prDMCV->PutItem(ivtdchuyen, _iCot[36], (_bstr_t)szfDMNhancong);		// Định mức nhân công  <-- Lấy trong CSDL
					prDMCV->PutItem(ivtdchuyen, _iCot[35], (_bstr_t)szfTNMay);			// Tài nguyên máy

					prDMCV->PutItem(ivtdchuyen, _iCot[37], 1);							// Hệ số năng suất
					prDMCV->PutItem(ivtdchuyen, _iCot[30], 1);							// Tổng thời gian thực hiện

					CString szTongcong = L"";
					szTongcong.Format(L"=%s*%s*%s",
						GAR(ivtdchuyen, _iCot[6], prDMCV, 0), GAR(ivtdchuyen, _iCot[36], prDMCV, 0), GAR(ivtdchuyen, _iCot[37], prDMCV, 0));
					prDMCV->PutItem(ivtdchuyen, _iCot[38], (_bstr_t)szTongcong);			// Tổng công

					CString szNCNgay = L"";
					szNCNgay.Format(L"=IFERROR(ROUNDUP(%s/%s,0),0)", GAR(ivtdchuyen, _iCot[38], prDMCV, 0), GAR(ivtdchuyen, _iCot[30], prDMCV, 0));
					prDMCV->PutItem(ivtdchuyen, _iCot[39], (_bstr_t)szNCNgay);			// Nhân công trong ngày
				}
				
/****************************************************************************************************/

				int vtcu = ivtdchuyen;
				if(iChen==0) ivtdchuyen++;
				else ivtdchuyen+=iChen;

				// Định dạng khung
				PRS = psDMCV->GetRange(prDMCV->GetItem(vtcu,1),prDMCV->GetItem(ivtdchuyen-1,iCtm[5]));
				PRS->Interior->PutColor(xlNone);
				PRS->Font->PutColor(BLACK);
				PRS->Font->PutBold(0);
				PRS->Font->PutItalic(0);

				PRS->Borders->GetItem(XlBordersIndex::xlInsideVertical)->PutWeight(xlThin);				

				PRS->Borders->GetItem(XlBordersIndex::xlInsideHorizontal)->PutWeight(xlThin);
				PRS->Borders->GetItem(XlBordersIndex::xlInsideHorizontal)->PutLineStyle(xlDot);

				PRS->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->PutWeight(xlThin);
				PRS->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->PutLineStyle(xlDot);

				PRS = psDMCV->GetRange(prDMCV->GetItem(vtcu,iCdm10),prDMCV->GetItem(ivtdchuyen-1,iCdm10));
				PRS->PutWrapText(0);
				PRS->PutHorizontalAlignment(xlLeft);

				PRS = psDMCV->GetRange(prDMCV->GetItem(vtcu,iCdm7),prDMCV->GetItem(ivtdchuyen-1,iCdm7));
				PRS->PutHorizontalAlignment(xlLeft);

				PRS = psDMCV->GetRange(prDMCV->GetItem(vtcu,1),prDMCV->GetItem(ivtdchuyen-1,1));
				PRS->EntireRow->Rows->AutoFit();
			}
		}

		ObjConn.CloseDatabase();

		// Sắp xếp lại stt
		int iEnd = FindComment(psDMCV,iCdm1,"END");
		int ntype = _wtoi(getVCell(psTS,L"TS_ATOSTT"));
		if(ntype!=0 && ntype!=1 && ntype!=2) ntype=0;
		f_Auto_stt_dmuc(prDMCV,ntype,0,8,iEnd,iCdm1,iCdm4,iCdm6);

		//int pos = _yn(L"Chuyển dữ liệu thành công! "
		//	L"\nChọn YES để đi đến danh mục NT công việc.",0);

		//if(pos==6)
		//{
		//	// ...
		//}

		xl->PutStatusBar((_bstr_t)L"Ready");

		int pos = psDMCV->GetVisible();
		if (pos!=-1) psDMCV->PutVisible(XlSheetVisibility::xlSheetVisible);
		psDMCV->Activate();

		// Hiển thị 2 cột đơn vị và khối lượng
		PRS = prDMCV->GetItem(1,iCdm9);
		int nCW = PRS->ColumnWidth;
		if(nCW==0)
		{
			psDMCV->GetRange(prDMCV->GetItem(1,iCdm8),prDMCV->GetItem(1,iCdm8))->EntireColumn->Hidden=false;
			psDMCV->GetRange(prDMCV->GetItem(1,iCdm9),prDMCV->GetItem(1,iCdm9))->EntireColumn->Hidden=false;
		}

		//xl->ActiveWindow->PutScrollRow(vt);
		//PRS = prDMCV->GetItem(vt,1);
		//PRS->EntireRow->Select();
	}
	catch(...){_s(L"Lỗi di chuyển biểu giá hợp đồng");}
}


void cls_main::f_move_bangvattu(_WorksheetPtr psh1,int itype)
{
	try
	{
		_bstr_t bsn = psh1->Name;
		RangePtr prg1 = psh1->Cells;
		CString szval=L"";

		_bstr_t shBTC = name_sheet("shNhomTC");
		_WorksheetPtr psBTC = xl->Sheets->GetItem(shBTC);
		RangePtr prBTC = psBTC->GetCells();
		int iVisible = psBTC->GetVisible();
		if (iVisible!=-1) psBTC->PutVisible(XlSheetVisibility::xlSheetVisible);

		// Đổ dữ liệu sang sheet Bang tieu chuan
		int icSTT = getColumn(psBTC,L"NHTC_STT");
		int icTen = getColumn(psBTC,L"NHTC_TEN");
		int icMa = getColumn(psBTC,L"NHTC_MA");
		int icTC = getColumn(psBTC,L"NHTC_TC");
		int icGC = getColumn(psBTC,L"NHTC_GC");
		int irSTT = getRow(psBTC,L"NHTC_STT");

		shTS = name_sheet("shTS");
		psTS=xl->Sheets->GetItem(shTS);
		prTS=psTS->GetCells();
		jNhomTC = _wtoi(getVCell(psTS,L"TS_GRPTC"));

		// Xác định cột
		int icot1 = getColumn(psh1,L"THXD_STT");
		int icot2 = getColumn(psh1,L"THXD_MSVT");
		int icot3 = getColumn(psh1,L"THXD_TEN");
		int icot4 = getColumn(psh1,L"THXD_DV");

		// Xác định vị trí cần di chuyển đến danh mục vật liệu
		shDMVL = name_sheet((_bstr_t)"shHSNTVL");
		psDMVL = xl->Sheets->GetItem(shDMVL);
		prDMVL = psDMVL->Cells;
		int iCdm1 = getColumn(psDMVL,L"DMVL_STT");
		int iCdm2 = getColumn(psDMVL,L"DMVL_MAVL");
		int iCdm3 = getColumn(psDMVL,L"DMVL_MAHS");
		int iCdm4 = getColumn(psDMVL,L"DMVL_ND");
		int iCdm5 = getColumn(psDMVL,L"DMVL_TC");

		int iCtg[7],iCtm[6];
		iCtg[1] = getColumn(psDMVL,L"DMVL_NK_NG");
		iCtg[2] = getColumn(psDMVL,L"DMVL_LM_NG");
		iCtg[3] = getColumn(psDMVL,L"DMVL_LM_KQ");
		iCtg[4] = getColumn(psDMVL,L"DMVL_NTNB_NG");
		iCtg[5] = getColumn(psDMVL,L"DMVL_YC");
		iCtg[6] = getColumn(psDMVL,L"DMVL_AB_NG");

		iCtm[1] = getColumn(psDMVL,L"DMVL_LM_GIO");
		iCtm[2] = getColumn(psDMVL,L"DMVL_NTNB_GIO");
		iCtm[3] = getColumn(psDMVL,L"DMVL_AB_GIO");
		iCtm[4] = iCtm[3];
		iCtm[5] = getColumn(psDMVL,L"DMVL_KBB");

		int vt=ivtdchuyen;

		if(getPathQLCL()==L"") return;
		if(connectDsn(pathMDB)==-1) return;
		ObjConn.OpenAccessDB(pathMDB,L"",L"");
		CRecordset *RecsetTC, *RecsetND;		

		int iChen=0;
		CString stemp=L"",szv[5]={L"",L"",L"",L"",L""};
		for (int i = 1; i <= slvchon; i++)
		{
			szval.Format(L"Đang tiến hành di chuyển vật tư. "
				L"Vui lòng đợi trong giây lát... (%d%s)", (int)(100 * i / slvchon), L"%");
			xl->PutStatusBar((_bstr_t)szval);

			// Kiểm tra vị trí lựa chọn có tồn tại công tác không?
			szv[0] = GIOR(iVtchon[i],icot1,prg1,L"");
			szv[1] = GIOR(iVtchon[i],icot2,prg1,L"");
			szv[2].Format(L"='%s'!%s",(LPCTSTR)bsn,GAR(iVtchon[i],icot3,prg1,0));

			if(_wtoi(szv[0])>0 && szv[1]!=L"")
			{
				slMaTC=0;
				if(ibgTC==1)
				{
					// Xác định tiêu chuẩn nếu có					
////////////////////////////////////////////////////////////////////////////////////////

					SqlString.Format(L"SELECT * FROM tbMaVL_Tieuchuan WHERE MaVL LIKE '%s' ORDER BY ID ASC;",szv[1]);
					RecsetTC = ObjConn.Execute(SqlString);

					while( !RecsetTC->IsEOF())
					{
						RecsetTC->GetFieldValue(L"MaTC",stemp);
						if(f_Check_tieuchuan(stemp)==0)
						{
							slMaTC++;
							sKTMaTC[slMaTC] = stemp;
							sKTTenTC[slMaTC] = stemp + L": Không tìm thấy tên tiêu chuẩn";

							SqlString.Format(L"SELECT *FROM tbTentieuchuan where MaTC='%s';",stemp);
							RecsetND = ObjConn.Execute(SqlString);
							while( !RecsetND->IsEOF())
							{
								RecsetND->GetFieldValue(L"TenTC",sKTTenTC[slMaTC]);
								break;
							}
							RecsetND->Close();
						}
						RecsetTC->MoveNext();
					}
					RecsetTC->Close();

////////////////////////////////////////////////////////////////////////////////////////

				}

				if(jNhomTC==0) iChen = slMaTC;
				else iChen=1;

				if(iChen==0) iChen++;

				if(iChen>1)
				{
					// Có nhiều hơn 1 tiêu chuẩn
					PRS = psDMVL->GetRange(prDMVL->GetItem(ivtdchuyen,1),prDMVL->GetItem(ivtdchuyen+iChen-1,1));
					PRS->EntireRow->Insert(xlShiftDown);					
				}
				else
				{
					PRS = prDMVL->GetItem(ivtdchuyen,1);
					PRS->EntireRow->Insert(xlShiftDown);
				}

				// Đổ dữ liệu
				prDMVL->PutItem(ivtdchuyen,iCdm2,(_bstr_t)szv[1]);
				prDMVL->PutItem(ivtdchuyen,iCdm4,(_bstr_t)szv[2]);
				
				PRS = prDMVL->GetItem(ivtdchuyen,1);
				PRS->EntireRow->Font->PutBold(0);
				PRS->EntireRow->Font->PutItalic(0);

				// Đổ tiêu chuẩn (nếu có)
				if(slMaTC>0)
				{
					if(jNhomTC==0)
					{
						for (int j = 0; j < iChen; j++)
						{
							prDMVL->PutItem(ivtdchuyen+j,iCdm5,(_bstr_t)sKTTenTC[j+1]);

							stemp = prDMVL->GetItem(ivtdchuyen+j,iCdm5+1);
							if(stemp==L"") prDMVL->PutItem(ivtdchuyen+j,iCdm5+1,(_bstr_t)L" ");
						}
					}
					else
					{
						// Đây là trường hợp group tiêu chuẩn
						// Kiểm tra xem nhóm tiêu chuẩn đã tồn tại bên sheet BTC chưa?
						int ivtr=1,iOk=0,icheck=0;
						int iEndBTC = prBTC->SpecialCells(xlCellTypeLastCell)->GetRow();
						for (int k = 1; k <= iEndBTC; k++)
						{
							szval = GIOR(k,icTen,prBTC,L"");
							if(szval!=L"")
							{
								szval = GIOR(k,icTC,prBTC,L"");
								if(szval==sKTTenTC[1])
								{
									ivtr=k;
									icheck=1;
									if(slMaTC>1)
									{
										for (int j = 2; j <= slMaTC; j++)
										{
											szval = GIOR(k+j-1,icTC,prBTC,L"");
											if(szval==sKTTenTC[j]) icheck++;
											else break;
										}
									}

									// Đã tồn tại tên nhóm tiêu chuẩn
									if(icheck==slMaTC)
									{							
										iOk=1;
										break;
									}
								}
							}
						}

						if(iOk==1)
						{
							szval.Format(L"='%s'!%s",(LPCTSTR)shBTC,GAR(ivtr,icTen,prBTC,0));
							prDMVL->PutItem(ivtdchuyen,iCdm5,(_bstr_t)szval);
						}
						else
						{
							// Chưa tồn tại --> Thêm mới vào đây
							ivtr=irSTT;
							for (int k = iEndBTC; k >= irSTT; k--)
							{
								szval = GIOR(k,icTC,prBTC,L"");
								szval.Trim();
								if(szval!=L"")
								{
									ivtr=k;
									break;
								}
							}

							PRS = psBTC->GetRange(prBTC->GetItem(ivtr+1,1),prBTC->GetItem(ivtr+1,icGC));
							PRS->Interior->PutColor(RGB(217,217,217));

							szval=L"";
							time_t now = time(0);
							tm *ltm = localtime(&now);
							szval.Format(L"%d%d",ltm->tm_min,ltm->tm_sec);
							int _nRand = _wtoi(szval);

							szval.Format(L"Nhóm TC%d%d%d",_nRand,rand()%10,rand()%10);
							prBTC->PutItem(ivtr+1,icTen,(_bstr_t)szval);

							szval.Format(L"=COUNTA(%s:%s)",GAR(irSTT+1,icTen,prBTC,3),GAR(ivtr+1,icTen,prBTC,0));
							prBTC->PutItem(ivtr+1,icSTT,(_bstr_t)szval);

							szval.Format(L"='%s'!%s",(LPCTSTR)shBTC,GAR(ivtr+1,icTen,prBTC,3));
							prDMVL->PutItem(ivtdchuyen,iCdm5,(_bstr_t)szval);

							for (int k = 1; k <= slMaTC; k++)
							{
								prBTC->PutItem(ivtr+k,icMa,(_bstr_t)sKTMaTC[k]);
								prBTC->PutItem(ivtr+k,icTC,(_bstr_t)sKTTenTC[k]);
							}
						}
					}
				}

				if(ibgTG==1 || ibgTM==1)
				{
					int vtcpy=0;
					// Xác định vị trí cần sao chép
					for (int j = ivtdchuyen-1; j >= 8; j--)
					{
						stemp = prDMVL->GetItem(j,iCdm1);
						if(_wtoi(stemp)>0)
						{
							vtcpy=j;
							break;
						}
					}

					// Sao chép thời gian
					if(vtcpy>0)
					{
						if(ibgTG==1)
						{
							for (int j = 1; j < 7; j++)
							{
								szval = GIOR(vtcpy,iCtg[j],prDMVL,L"");
								f_ktra_date(szval,prDMVL,ivtdchuyen,iCtg[j]);
							}
						}

						if(ibgTM==1)
						{
							for (int j = 1; j < 4; j++)
							{
								szval = GIOR(vtcpy,iCtm[j],prDMVL,L"");
								prDMVL->PutItem(ivtdchuyen,iCtm[j],(_bstr_t)szval);
							}
						}
					}
				}

				int vtcu = ivtdchuyen;
				if(iChen==0) ivtdchuyen++;
				else ivtdchuyen+=iChen;

				// Định dạng khung
				PRS = psDMVL->GetRange(prDMVL->GetItem(vtcu,1),prDMVL->GetItem(ivtdchuyen-1,iCtm[5]));
				PRS->Interior->PutColor(xlNone);
				PRS->Font->PutColor(BLACK);
				PRS->Font->PutBold(0);
				PRS->Font->PutItalic(0);

				PRS->Borders->GetItem(XlBordersIndex::xlInsideVertical)->PutWeight(xlThin);				

				PRS->Borders->GetItem(XlBordersIndex::xlInsideHorizontal)->PutWeight(xlThin);
				PRS->Borders->GetItem(XlBordersIndex::xlInsideHorizontal)->PutLineStyle(xlDot);

				PRS->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->PutWeight(xlThin);
				PRS->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->PutLineStyle(xlDot);

				PRS = psDMVL->GetRange(prDMVL->GetItem(vtcu,iCdm5),prDMVL->GetItem(ivtdchuyen-1,iCdm5));
				PRS->PutWrapText(0);
				PRS->PutHorizontalAlignment(xlLeft);

				PRS = psDMVL->GetRange(prDMVL->GetItem(vtcu,iCdm4),prDMVL->GetItem(ivtdchuyen-1,iCdm4));
				PRS->PutHorizontalAlignment(xlLeft);

				PRS = psDMVL->GetRange(prDMVL->GetItem(vtcu,1),prDMVL->GetItem(ivtdchuyen-1,1));
				PRS->EntireRow->Rows->AutoFit();
			}
		}

		ObjConn.CloseDatabase();

		// Sắp xếp lại stt
		int iEnd = FindComment(psDMVL,iCdm1,"END");
		int ntype = _wtoi(getVCell(psTS,L"TS_ATOSTT"));
		if(ntype!=0 && ntype!=1 && ntype!=2) ntype=0;
		f_Auto_stt_dmuc(prDMVL,ntype,0,8,iEnd,iCdm1,iCdm2,iCdm3);

		//int pos = _yn(L"Chuyển dữ liệu thành công! "
		//	L"\nChọn YES để đi đến danh mục NT vật liệu.",0);

		//if(pos==6)
		//{
		//	// ...
		//}

		xl->PutStatusBar((_bstr_t)L"Ready");

		int pos = psDMVL->GetVisible();
		if (pos!=-1) psDMVL->PutVisible(XlSheetVisibility::xlSheetVisible);
		psDMVL->Activate();

		//xl->ActiveWindow->PutScrollRow(vt);
		//PRS = prDMVL->GetItem(vt,1);
		//PRS->EntireRow->Select();
	}
	catch(...){_s(L"Lỗi kết nối bảng vật tư");}
}


// 09.06.2017: Kiểm tra tiêu chuẩn trùng
int cls_main::f_Check_tieuchuan(CString sKtraTC)
{
	try
	{
		if(slMaTC==0) return 0;
		for (int i = 1; i <= slMaTC; i++)
			if(sKtraTC==sKTMaTC[i]) return 1;
	}
	catch(...){return 0;}

	return 0;
}


// Bổ sung 05.05.2018
void cls_main::f_enter_giaidoan()
{
	try
	{
		psDMGD = pWb->ActiveSheet;
		prDMGD = psDMGD->Cells;

		int iRow = psDMGD->Application->ActiveCell->Row;
		int iColumn = psDMGD->Application->ActiveCell->Column;

		int iId = getColumn(psDMGD,L"DMGD_STT");
		int iMa = getColumn(psDMGD,L"DMGD_HS");
		int iND = getColumn(psDMGD,L"DMGD_ND");
		int iKBB = getColumn(psDMGD,L"DMGD_KBB");
		int iEnd =FindComment(psDMGD,iId,L"END");

		int ngay1 = getColumn(psDMGD,L"DMGD_NTNB_NG");
		int ngay2 = getColumn(psDMGD,L"DMGD_YC");
		int ngay3 = getColumn(psDMGD,L"DMGD_AB_NG");

		CString szval = L"";
		if(iRow>=8 && iRow<iEnd)
		{
			if(iColumn==ngay1 || iColumn==ngay2 || iColumn==ngay3)
			{
				// Leo edit 01.06.2018
				szval= GIOR(iRow,1,prDMGD,L"");
				if(_wtoi(szval)>0) CheckNhomKy(psDMGD,iRow,iKBB);
			}
			else if(iColumn==iMa)
			{
				szval = GIOR(iRow, iMa, prDMGD, L"");
				szval.MakeUpper();

				if (szval != L"")
				{
					PRS = psDMGD->GetRange(prDMGD->GetItem(iRow, iId), prDMGD->GetItem(iRow, iKBB));
					if (szval.Left(2) == L"HM")
					{						
						szval = GIOR(iRow, iND, prDMGD, L"");
						if (szval == L"")
						{
							PRS->Font->PutBold(1);
							PRS->Interior->PutColor(LIGHTGRAY);
							PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->Weight = xlThin;
							PRS->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->Weight = xlThin;
							prDMGD->PutItem(iRow, iMa, (_bstr_t)szval);
							prDMGD->PutItem(iRow, iND, (_bstr_t)L"HẠNG MỤC SỐ ...");
							prDMGD->PutItem(iRow, iId, (_bstr_t)L"");
						}
					}
					else
					{
						PRS->EntireRow->Font->PutBold(0);
						PRS->Font->PutItalic(0);
						PRS->Interior->PutColor(xlNone);

						PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->PutWeight(xlThin);
						PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->PutLineStyle(xlDot);
						PRS->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->PutWeight(xlThin);
						PRS->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->PutLineStyle(xlDot);
					}
				}

				int dem = 0;
				for (int i = 8; i < iEnd; i++)
				{
					szval = GIOR(i, iMa, prDMGD, L"");
					szval.MakeUpper();
					if (szval != L"")
					{
						if (szval.Left(2) != L"HM")
						{
							dem++;
							prDMGD->PutItem(i, iId, dem);
						}
					}
					else prDMGD->PutItem(i, iId, (_bstr_t)L"");
				}
			}			
		}
	}
	catch(...){}
}

void cls_main::_getDinhmucNhancong(CString szMaCV, CString szLLk)
{
	try
	{
		CRecordset* Recset;
		CRecordset* RCS2;

		CString szMaDM = L"", szMSVT = L"", szHPhi = L"", szTenmay = L"";
		CString szTruyvan = L"";
		CString szMa[50], szLink[50];

		szMaCV.Replace(L"\n", L"");
		_TackTukhoaChuan(szMaCV, L"|", L" ", 0, 0);
		for (int i = 1; i <= iKey; i++) szMa[i] = pKey[i];

		szLLk.Replace(L"\n", L"");
		_TackTukhoaChuan(szLLk, L"|", L";", 0, 0);
		for (int i = 1; i <= iKey; i++) szLink[i] = pKey[i];

		for (int i = 1; i <= iKey; i++)
		{
			szconDMNC[i] = L"";
			szconTNMay[i] = L"";

			// Xác định định mức nhân công
			// Xác định mã định mức
			if (_FileExists(szLink[i], 0) == 1)
			{
				szMaDM = L"";
				ObjConnCV.OpenAccessDB(szLink[i], L"", L"");
				SqlString.Format(L"SELECT * FROM tbDonGia WHERE MHDG LIKE '%s';", szMa[i]);
				Recset = ObjConnCV.Execute(SqlString);
				while (!Recset->IsEOF())
				{
					Recset->GetFieldValue(L"MHDM", szMaDM);
					break;
				}
				Recset->Close();

				if (szMaDM != L"")
				{
					szMSVT = L"", szHPhi = L"", szTenmay = L"";
					SqlString.Format(L"SELECT * FROM tbDinhMuc WHERE MHDM LIKE '%s';", szMaDM);
					Recset = ObjConnCV.Execute(SqlString);
					while (!Recset->IsEOF())
					{
						Recset->GetFieldValue(L"MSVT", szMSVT);
						if (szMSVT.Left(1) == L"N")
						{
							Recset->GetFieldValue(L"HPVT", szHPhi);
							szHPhi.Replace(L",", L".");
							szHPhi.Trim();
							if (szHPhi == L"") szHPhi = L"0";

							szconDMNC[i] = szHPhi;
							szconDMNC[i] += L"+";

							szfDMNhancong += szHPhi;
							szfDMNhancong += L"+";
						}
						else if (szMSVT.Left(1) == L"M")
						{
							// Tìm máy trong từ điển và giá ca máy
							SqlString.Format(L"SELECT * FROM tbTuDienVatTu WHERE MSVT LIKE '%s';", szMSVT);
							RCS2 = ObjConnCV.Execute(SqlString);
							while (!RCS2->IsEOF())
							{
								RCS2->GetFieldValue(L"TVT", szTenmay);
								break;
							}
							RCS2->Close();

							if (szTenmay == L"")
							{
								SqlString.Format(L"SELECT * FROM tbGiaCaMay WHERE MH LIKE '%s';", szMSVT);
								RCS2 = ObjConnCV.Execute(SqlString);
								while (!RCS2->IsEOF())
								{
									RCS2->GetFieldValue(L"LM_TB", szTenmay);
									break;
								}
								RCS2->Close();
							}

							if (szTenmay != L"" && szTenmay != L"Máy khác")
							{
								szTenmay.Replace(L";", L",");
								szTenmay += L" [1];";

								szconTNMay[i] += szTenmay;
								szfTNMay += szTenmay;
							}
						}

						Recset->MoveNext();
					}
					Recset->Close();
				}

				ObjConnCV.CloseDatabase();
			}
		}

		CString szval = L"";
		if (szfDMNhancong != L"")
		{
			szval.Format(L"=%s", szfDMNhancong.Left(szfDMNhancong.GetLength() - 1));
			szfDMNhancong = szval;
			if (szfDMNhancong.Find(L"+") == -1) szfDMNhancong.Replace(L"=", L"");
		}

		if (szfTNMay != L"")
		{
			szval.Format(L"'%s", szfTNMay);
			szfTNMay = szval;
		}

		for (int i = 1; i <= iKey; i++)
		{
			if (szconDMNC[i] != L"")
			{
				szval.Format(L"=%s", szconDMNC[i].Left(szconDMNC[i].GetLength() - 1));
				szconDMNC[i] = szval;
				if (szconDMNC[i].Find(L"+") == -1) szconDMNC[i].Replace(L"=", L"");
			}

			if (szconTNMay[i] != L"")
			{
				szval.Format(L"'%s", szconTNMay[i]);
				szconTNMay[i] = szval;
			}
		}
	}
	catch (...) {}
}


void cls_main::f_Danh_sachfile()
{
	try
	{
		CString spath = GIOR(iRowDLGoc, iColumnDLGoc, prTS, L"");
		CString szfolder = GIOR(iRowDLGoc + 1, iColumnDLGoc, prTS, L"");
		if (spath.Right(1) != L"\\") spath += L"\\";
		sDDanfile = spath + szfolder;

		slfileDL = 0;
		CString szval = L"", scheck = L"";
		for (int i = iRowDLGoc + 1; i < iRowDLGoc + 100; i++)
		{
			szval = GIOR(i, iColumnDLGoc + 2, prTS, L"");
			if (szval == L"") break;

			scheck = GIOR(i, iColumnDLGoc + 3, prTS, L"");
			if (_wtoi(scheck) == 1)
			{
				if (szval.Find(L"\\") == -1) scheck.Format(L"%s\\%s", sDDanfile, szval);
				else
				{
					scheck = szval;
					szval = _TackTenFile(szval, 1);
				}

				if (_FileExists(scheck, 0) == 1)
				{
					slfileDL++;
					sFiledulieu[slfileDL] = szval;
					sFpathdulieu[slfileDL] = scheck;
				}
			}
		}

		if (slfileDL == 0)
		{
			_s(L"Không tồn tại tệp tin dữ liệu. Vui lòng kiểm tra lại đường dẫn.");
			return;
		}
	}
	catch (...) {}
}