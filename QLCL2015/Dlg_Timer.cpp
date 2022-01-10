#include "Dlg_Timer.h"
#include "cls_main.h"

// Dlg_Timer dialog

IMPLEMENT_DYNAMIC(Dlg_Timer, CDialogEx)

Dlg_Timer::Dlg_Timer(CWnd* pParent /*=NULL*/)
	: CDialogEx(Dlg_Timer::IDD, pParent)
{

}

Dlg_Timer::~Dlg_Timer()
{
}

void Dlg_Timer::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, TM_MC1, cmc1);	
	DDX_Control(pDX, TM_CK1, cck1);	
}


BEGIN_MESSAGE_MAP(Dlg_Timer, CDialogEx)
	ON_WM_TIMER()
	ON_BN_CLICKED(TM_B2, &Dlg_Timer::OnBnClickedB2)
	ON_BN_CLICKED(TM_B1, &Dlg_Timer::OnBnClickedB1)
	ON_NOTIFY(MCN_SELECT, TM_MC1, &Dlg_Timer::OnMcnSelectMc1)
END_MESSAGE_MAP()


// Dlg_Timer message handlers
BOOL Dlg_Timer::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	// Biến xác định độ trễ click đúp
	iDbclick=0;
	sDbcl[1]=L"1",sDbcl[2]=L"2";
	SetTimer(1, 800, NULL);	// 1000ms = 1 second

	// Xác định vị trí hiển thị hộp thoại
	f_Set_location();

	// Lấy ngày hiện hành
	f_set_date();

	if(_iTimeDlg>0) cck1.EnableWindow(0);

	return TRUE; 
}


// Bắt sự kiện click Enter
BOOL Dlg_Timer::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if (pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(TM_MC1))
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


void Dlg_Timer::f_Set_location()
{
	try
	{
		// Xác định vị trí chỉ chuột
		POINT p;
		if (GetCursorPos(&p))
		{
			CRect r;
			GetWindowRect(&r);
			ScreenToClient(&r);

			// Xác định kích thước hộp thoại
			int h = r.Height();
			int w = r.Width();

			// Xác định kích thước màn hình
			int isx = GetSystemMetrics(SM_CXSCREEN);
			int isy = GetSystemMetrics(SM_CYSCREEN);

			// Xác định vị trí hiển thị hộp thoại
			if(p.x+w+10 >= isx) p.x-=(w+10);
			if(p.y+h-100 >= isy) p.y-=50;

			if(_iTimeDlg>0) MoveWindow(p.x + 20, p.y - 30, r.right - r.left, r.bottom - r.top, TRUE);
			else MoveWindow(p.x + 10, p.y - 200 , r.right - r.left, r.bottom - r.top, TRUE);
		}
	}
	catch(...){}
}


void Dlg_Timer::OnTimer(UINT_PTR nIDEvent)
{
	iDbclick=0;
	CDialog::OnTimer(nIDEvent);
}


void Dlg_Timer::OnBnClickedB2()
{
	CDialog::OnCancel();
}


void Dlg_Timer::OnBnClickedB1()
{
	try
	{
		CTime dateTime;
		cmc1.GetCurSel(dateTime);

		CString sday,smonth,syear;
		int d = dateTime.GetDay();
		if(d<=9) sday.Format(L"0%d",d);
		else sday.Format(L"%d",d);

		int m = dateTime.GetMonth();
		if(m<=9) smonth.Format(L"0%d",m);
		else smonth.Format(L"%d",m);

		int y = dateTime.GetYear();
		if(y<=99) syear.Format(L"20%d",y);
		else syear.Format(L"%d",y);

		if(_iNgay==8) syear=syear.Right(2);
		_bNThang.Format(L"%s/%s/%s",sday,smonth,syear);
		if(_iTimeDlg>0)
		{
			// _iTimeDlg>0 là thao tác trên các hộp thoại khác
			m_timeDlg[_iTimeDlg-1].SetWindowText(_bNThang);
			CDialog::OnOK();
			return; 
		}

		// Nếu trạng thái khác 1 (=0) thì sẽ không sử dụng chạy hàm f_Autocheck_...
		CString sgt = getVCell(psTS,L"TS_NGNHAP");	// sgt=TRUE (-1) hoặc FALSE (0)
		if(_wtoi(sgt)==0 || sgt==_T("FALSE")) iATdate=0;
		else iATdate=1;		// -1 LẤY LÀ 1

		pSh1 = pWb->ActiveSheet;
		RangePtr pRg1 = pSh1->Cells;
		int _irow = pSh1->Application->ActiveCell->Row;
		int _icol = pSh1->Application->ActiveCell->Column;

		if(iATdate==1) f_Autocheck_date(pRg1,_irow,_icol);
		else f_ktra_date(_bNThang,pRg1,_irow,_icol);

		f_Check_nghile(pRg1,_irow,_icol);

		_bstr_t sh = pSh1->CodeName;

		// Leo edit 01.06.2018
		sgt= GIOR(_irow,1,pRg1,L"");
		if(_wtoi(sgt)>0)
		{
			CString szcotKBB=L"";					
			if(sh==(_bstr_t)L"shHSNTVL") szcotKBB = L"DMVL_KBB";
			else if(sh==(_bstr_t)L"shHSNTCV") szcotKBB = L"DMBB_KBB";
			else if(sh==(_bstr_t)L"shHSNTGD") szcotKBB = L"DMGD_KBB";
			if(szcotKBB!=L"") CheckNhomKy(pSh1,_irow,getColumn(pSh1,szcotKBB));
		}

		int nck = cck1.GetCheck();
		if(nck==0) CDialog::OnOK();
		
		if(sh!=(_bstr_t)L"shHSNTCV") return;

// Phần này dành cho danh mục công việc
		_WorksheetPtr psDMCV = pSh1;
		RangePtr prDMCV = psDMCV->Cells;
		int iRow = _irow;

		int jColumnEnd = (int)psDMCV->Cells->SpecialCells(xlCellTypeLastCell)->GetColumn();
		CString szval = GIOR(iRow, jColumnEnd, prDMCV, L"");
		if (szval != L"") jColumnEnd++;

		RangePtr PRS = prDMCV->GetItem(iRow, jColumnEnd);
		PRS->PutNumberFormat(L"General");

		int jTime = getColumn(pSh1,L"DMBB_HK_TONGNGAY");
		int jStart = getColumn(pSh1,L"DMBB_HK_NGAYBD");
		int jFinish = getColumn(pSh1,L"DMBB_HK_NGAYKT");
		int jLienhe = getColumn(pSh1, L"DMBB_HK_LIENKET");
		int jTongnghi = getColumn(pSh1, L"DMBB_HK_TONGNGHI");

		PRS = prDMCV->GetItem(iRow, jTime);
		PRS->PutNumberFormat(L"General");

		int num = 0;
		int iColumn = _icol;
		CString szTong = L"", szBD = L"", szKT = L"", szLienhe = L"";

		szLienhe = GIOR(iRow, jLienhe, prDMCV, L"");
		szLienhe.MakeUpper();

		if (iColumn == jStart)
		{
			_Change_Start(prDMCV, iRow, jTime, jStart, jFinish, jLienhe, jTongnghi, jColumnEnd);
		}
		else if (iColumn == jFinish)
		{
			_Change_Finish(prDMCV, iRow, jTime, jStart, jFinish, jLienhe, jTongnghi, jColumnEnd);
		}
	}
	catch(...){}
}


void Dlg_Timer::_Change_Start(RangePtr prDMCV, int iRow, int jTime, int jStart, int jFinish, int jLienhe, int jTongnghi, int jColumnEnd)
{
	try
	{
		int num = 0;
		CString szBD = L"", szKT = L"", szLienhe = L"";
		RangePtr PRS;

		CString szTong = GIOR(iRow, jTime, prDMCV, L"");
		if (szTong != L"")
		{
			num = _wtoi(szTong);
			if (num <= 0) num = 1;
			prDMCV->PutItem(iRow, jTime, num);

			szKT.Format(L"=IF(%s>0,%s+%s-1,%s)+%s", GAR(iRow, jTime, prDMCV, 0),
				GAR(iRow, jStart, prDMCV, 0), GAR(iRow, jTime, prDMCV, 0),
				GAR(iRow, jStart, prDMCV, 0), GAR(iRow, jTongnghi, prDMCV, 0));
			prDMCV->PutItem(iRow, jFinish, (_bstr_t)szKT);

			if (szLienhe.Find(L"FF") >= 0 || szLienhe.Find(L"SF") >= 0)
			{
				PRS = prDMCV->GetItem(iRow, jLienhe);
				PRS->ClearContents();
			}
		}
		else
		{
			szBD = GIOR(iRow, jStart, prDMCV, L"");
			szKT = GIOR(iRow, jFinish, prDMCV, L"");
			if (szKT != L"")
			{
				// return =0 --> szBD > szKT
				int ilen = szKT.GetLength();
				if (compare_date(ilen, szBD, szKT) == 0)
				{
					prDMCV->PutItem(iRow, jTime, 1);

					szKT.Format(L"=IF(%s>0,%s+%s-1,%s)+%s", GAR(iRow, jTime, prDMCV, 0),
						GAR(iRow, jStart, prDMCV, 0), GAR(iRow, jTime, prDMCV, 0),
						GAR(iRow, jStart, prDMCV, 0), GAR(iRow, jTongnghi, prDMCV, 0));
					prDMCV->PutItem(iRow, jFinish, (_bstr_t)szKT);

					if (szLienhe.Find(L"FF") >= 0 || szLienhe.Find(L"SF") >= 0)
					{
						PRS = prDMCV->GetItem(iRow, jLienhe);
						PRS->ClearContents();
					}
				}
				else
				{
					szTong.Format(L"=%s-%s-%s+1", GAR(iRow, jFinish, prDMCV, 0),
						GAR(iRow, jStart, prDMCV, 0), GAR(iRow, jTongnghi, prDMCV, 0));
					prDMCV->PutItem(iRow, jColumnEnd, (_bstr_t)szTong);
					szTong = GIOR(iRow, jColumnEnd, prDMCV, L"");
					prDMCV->PutItem(iRow, jTime, (_bstr_t)szTong);

					PRS = prDMCV->GetItem(iRow, jColumnEnd);
					PRS->ClearContents();
				}
			}
		}

		CString szval = L"";
		szval.Format(L"%d", _wtoi(szLienhe));
		if ((_wtoi(szval) > 0 && szval == szLienhe) || szLienhe.Find(L"FS") >= 0 || szLienhe.Find(L"SS") >= 0)
		{
			PRS = prDMCV->GetItem(iRow, jLienhe);
			PRS->ClearContents();
		}
	}
	catch (...) {}
}


void Dlg_Timer::_Change_Finish(RangePtr prDMCV, int iRow, int jTime, int jStart, int jFinish, int jLienhe, int jTongnghi, int jColumnEnd)
{
	try
	{
		int num = 0;
		CString szBD = L"", szKT = L"", szLienhe = L"", szTong = L"";
		RangePtr PRS = prDMCV->GetItem(iRow, jFinish);
		szKT = PRS->GetFormula();
		szKT.Replace(L"$", L"");
		CString szval = GAR(iRow, jStart, prDMCV, 0);
		if (szKT.Find(szval) >= 0)
		{
			szBD = GIOR(iRow, jStart, prDMCV, L"");
			szKT = GIOR(iRow, jFinish, prDMCV, L"");
			if (szBD != L"")
			{
				// return =0 --> szBD > szKT
				int ilen = szKT.GetLength();
				if (compare_date(ilen, szBD, szKT) == 0)
				{
					prDMCV->PutItem(iRow, jTime, 1);
					prDMCV->PutItem(iRow, jStart, (_bstr_t)szKT);

					szval.Format(L"%d", _wtoi(szLienhe));
					if ((_wtoi(szval) > 0 && szval == szLienhe) || szLienhe.Find(L"FS") >= 0 || szLienhe.Find(L"SS") >= 0)
					{
						PRS = prDMCV->GetItem(iRow, jLienhe);
						PRS->ClearContents();
					}
				}
				else
				{
					szTong.Format(L"=%s-%s-%s+1", GAR(iRow, jFinish, prDMCV, 0),
						GAR(iRow, jStart, prDMCV, 0), GAR(iRow, jTongnghi, prDMCV, 0));
					prDMCV->PutItem(iRow, jColumnEnd, (_bstr_t)szTong);
					szTong = GIOR(iRow, jColumnEnd, prDMCV, L"");
					prDMCV->PutItem(iRow, jTime, (_bstr_t)szTong);

					PRS = prDMCV->GetItem(iRow, jColumnEnd);
					PRS->ClearContents();
				}
			}
			else
			{
				num = _wtoi(szTong);
				if (num <= 0) num = 1;
				prDMCV->PutItem(iRow, jTime, num);

				szBD.Format(L"=%s-%s-%s+1", GAR(iRow, jFinish, prDMCV, 0),
					GAR(iRow, jTime, prDMCV, 0), GAR(iRow, jTongnghi, prDMCV, 0));
				prDMCV->PutItem(iRow, jColumnEnd, (_bstr_t)szBD);
				szBD = GIOR(iRow, jColumnEnd, prDMCV, L"");
				f_ktra_date(szBD, prDMCV, iRow, jStart);

				PRS = prDMCV->GetItem(iRow, jColumnEnd);
				PRS->ClearContents();

				szval.Format(L"%d", _wtoi(szLienhe));
				if ((_wtoi(szval) > 0 && szval == szLienhe) || szLienhe.Find(L"FS") >= 0 || szLienhe.Find(L"SS") >= 0)
				{
					PRS = prDMCV->GetItem(iRow, jLienhe);
					PRS->ClearContents();
				}
			}
		}
		else
		{
			szTong = GIOR(iRow, jTime, prDMCV, L"");
			if (szTong != L"")
			{
				num = _wtoi(szTong);
				if (num <= 0) num = 1;
				prDMCV->PutItem(iRow, jTime, num);

				szBD.Format(L"=%s-%s-%s+1", GAR(iRow, jFinish, prDMCV, 0),
					GAR(iRow, jTime, prDMCV, 0), GAR(iRow, jTongnghi, prDMCV, 0));
				prDMCV->PutItem(iRow, jColumnEnd, (_bstr_t)szBD);
				szBD = GIOR(iRow, jColumnEnd, prDMCV, L"");
				f_ktra_date(szBD, prDMCV, iRow, jStart);

				PRS = prDMCV->GetItem(iRow, jColumnEnd);
				PRS->ClearContents();

				szval.Format(L"%d", _wtoi(szLienhe));
				if ((_wtoi(szval) > 0 && szval == szLienhe) || szLienhe.Find(L"FS") >= 0 || szLienhe.Find(L"SS") >= 0)
				{
					PRS = prDMCV->GetItem(iRow, jLienhe);
					PRS->ClearContents();
				}
			}
			else
			{
				szBD = GIOR(iRow, jStart, prDMCV, L"");
				szKT = GIOR(iRow, jFinish, prDMCV, L"");
				if (szBD != L"")
				{
					// return =0 --> szBD > szKT
					int ilen = szBD.GetLength();
					if (compare_date(ilen, szBD, szKT) == 0)
					{
						prDMCV->PutItem(iRow, jTime, 1);
						prDMCV->PutItem(iRow, jStart, (_bstr_t)szKT);

						szval.Format(L"%d", _wtoi(szLienhe));
						if ((_wtoi(szval) > 0 && szval == szLienhe) || szLienhe.Find(L"FS") >= 0 || szLienhe.Find(L"SS") >= 0)
						{
							PRS = prDMCV->GetItem(iRow, jLienhe);
							PRS->ClearContents();
						}
					}
					else
					{
						szTong.Format(L"=%s-%s-%s+1", GAR(iRow, jFinish, prDMCV, 0),
							GAR(iRow, jStart, prDMCV, 0), GAR(iRow, jTongnghi, prDMCV, 0));
						prDMCV->PutItem(iRow, jColumnEnd, (_bstr_t)szTong);
						szTong = GIOR(iRow, jColumnEnd, prDMCV, L"");
						prDMCV->PutItem(iRow, jTime, (_bstr_t)szTong);

						PRS = prDMCV->GetItem(iRow, jColumnEnd);
						PRS->ClearContents();
					}
				}
			}
		}

		if (szLienhe.Find(L"FF") >= 0 || szLienhe.Find(L"SF") >= 0)
		{
			PRS = prDMCV->GetItem(iRow, jLienhe);
			PRS->ClearContents();
		}
	}
	catch (...) {}
}



void Dlg_Timer::f_Autocheck_date(RangePtr pr1,int ir,int ic)
{
	try
	{
		CString sNgay = GIOR(ir,ic,pr1,L"");
		sNgay.Trim();

		if(sNgay!=_bNThang)
		{
			f_ktra_date(_bNThang,pr1,ir,ic);
			sNgay = GIOR(ir,ic,pr1,L"");

			// Số ký tự ngày tháng (mặc định =10)
			_iNgay = _bNThang.GetLength();

			// Kiểm tra sheet nào?
			_bstr_t sh1 = pSh1->CodeName;
			if(sh1==(_bstr_t)L"shHSNTVL")
			{
				// Danh mục vật liệu
				// Xác định các cột ngày tháng cần thiết
				int jcot[6]={7,9,11,13,15,16};

				jcot[0] = getColumn(pSh1,"DMVL_NK_NG");
				jcot[1] = getColumn(pSh1,"DMVL_LM_NG");
				jcot[2] = getColumn(pSh1,"DMVL_LM_KQ");
				jcot[3] = getColumn(pSh1,"DMVL_NTNB_NG");
				jcot[4] = getColumn(pSh1,"DMVL_YC");
				jcot[5] = getColumn(pSh1,"DMVL_AB_NG");

				// Kiểm tra xem iVtri nằm ở đâu
				int iKtr=-1;
				int ivtri=-1;
				for (int i = 0; i < 6; i++)
				{
					if(ic==jcot[i])
					{
						ivtri=i;
						break;
					}
				}

				CString val = L"";
				if(ivtri>=0)
				{
					// ivtri = số vị trí cột chứa ngày tháng
					// Nếu ivtri>0 -> Kiểm tra vị trí trước đó
					if(ivtri>0)
					{
						iKtr=-1;
						for (int i = ivtri-1; i >= 0; i--)
						{
							val = GIOR(ir,jcot[i],pr1,L"");
							val.Trim();
							if(val!=L"")
							{
								iKtr=i;
								break;
							}
						}

						if(iKtr>=0) f_SetGtriDate(pr1,sNgay,ir,ic,jcot[iKtr],0);
					}

					// Kiểm tra tiếp ngày kế sau (nếu có): ivtri<5
					if(ivtri<5)
					{
						iKtr=-1;
						for (int i = ivtri+1; i <= 5; i++)
						{
							val = GIOR(ir,jcot[i],pr1,L"");
							val.Trim();
							if(val!=L"")
							{
								iKtr=i;
								break;
							}
						}

						if(iKtr>=0) f_SetGtriDate(pr1,sNgay,ir,ic,jcot[iKtr],1);
					}
				}
			}
			else if(sh1==(_bstr_t)L"shHSNTCV")
			{
				// Danh mục công việc
				// Xác định các cột ngày tháng cần thiết
				int jcot[7]={9,11,12,14,15,18,19};

				jcot[0] = getColumn(pSh1,"DMBB_TN_NGAY");
				jcot[1] = getColumn(pSh1,"DMBB_TN_KQ");
				jcot[2] = getColumn(pSh1,"DMBB_NB_NGAY");
				jcot[3] = getColumn(pSh1,"DMBB_PHIEUYC");
				jcot[4] = getColumn(pSh1,"DMBB_AB_NGAY");
				jcot[5] = getColumn(pSh1,"DMBB_HK_NGAYBD");
				jcot[6] = getColumn(pSh1,"DMBB_HK_NGAYKT");

				// Kiểm tra xem vị trí ic nằm ở đâu
				int iKtr=-1;
				CString val=L"";

				if(ic==jcot[0])
				{
					// Cột I (9) = Ngày lấy mẫu thí nghiệm
					// I <= K
					f_SetGtriDate(pr1,sNgay,ir,ic,jcot[1],1);

					// Kiểm tra tiếp điều kiện cột I <= L(N,O)
					iKtr=-1;
					for (int i = 2; i <= 4; i++)
					{
						val = GIOR(ir,jcot[i],pr1,L"");
						val.Trim();
						if(val!=L"")
						{
							iKtr=i;
							break;
						}
					}

					if(iKtr>=2) f_SetGtriDate(pr1,sNgay,ir,ic,jcot[iKtr],1);
				}
				else if(ic==jcot[1])
				{
					// Cột K (11) = Ngày lấy kết quả mẫu thí nghiệm
					f_SetGtriDate(pr1,sNgay,ir,ic,jcot[0],0);
				}
				else if(ic==jcot[2])
				{
					// Cột L (12) = Ngày nghiệm thu nội bộ
					// I <= L <= N (O)
					// Kiểm tra cột I
					int N = f_SetGtriDate(pr1,sNgay,ir,ic,jcot[0],0);

					// Kiểm tra tiếp cột N (O)
					if(N==1)
					{
						iKtr=-1;
						for (int i = 3; i <= 4; i++)
						{
							val = GIOR(ir,jcot[i],pr1,L"");
							val.Trim();
							if(val!=L"")
							{
								iKtr=i;
								break;
							}
						}

						if(iKtr>=3) f_SetGtriDate(pr1,sNgay,ir,ic,jcot[iKtr],1);
					}
				}
				else if(ic==jcot[3])
				{
					// Cột N (14) = Ngày ghi phiếu yêu cầu
					// (I) L <= N <= O
					// Kiểm tra cột (I) L
					iKtr=2;
					val = GIOR(ir,jcot[iKtr],pr1,L"");
					val.Trim();
					if(val==L"")
					{
						iKtr=0;
						val = GIOR(ir,jcot[iKtr],pr1,L"");
						val.Trim();
						if(val==L"") iKtr=-1;
					}

					int N = 1;
					if(iKtr!=-1) N = f_SetGtriDate(pr1,sNgay,ir,ic,jcot[iKtr],0);

					// Kiểm tra tiếp cột O
					if(N==1) f_SetGtriDate(pr1,sNgay,ir,ic,jcot[4],1);
				}
				else if(ic==jcot[4])
				{
					// Cột O (15) = Ngày nghiệm thu công việc
					f_SetGtriDate(pr1,sNgay,ir,ic,jcot[3],0);
				}
				else if(ic==jcot[5])
				{
					// Cột R (18) = Ngày bắt đầu nhật ký
					f_SetGtriDate(pr1,sNgay,ir,ic,jcot[6],1);
				}
				else if(ic==jcot[6])
				{
					// Cột S (19) = Ngày kết thúc nhật ký
					f_SetGtriDate(pr1,sNgay,ir,ic,jcot[5],0);
				}
			}
		}
	}
	catch(...){_s(L"#QL7.12: Lỗi kiểm tra ngày tháng.");}
}


// Bổ sung 25.08.2017
void Dlg_Timer::f_Check_nghile(RangePtr pr1,int ir,int ic)
{
	try
	{
		long gtri=0;
		if(ichkLe==1)
		{
			// Đọc từ file MDB
			gtri = f_Read_Nhatky();

			// Đọc file dữ liệu csv
			CString spath=L"";
			spath.Format(L"%s\\%s",_fGPathFolder(L"ToolQLCL.dll"),_dFileNle);
			spath.Replace(L"\\\\",L"\\");
			gtri += f_Read_file_CSV(spath,gtri);
		}

		RangePtr PRS = pr1->GetItem(ir,ic);
		PRS->ClearComments();

		int _iktra = 0;
		CString sthumay = L"";
		CString sngay = GIOR(ir, ic, pr1, L"");
		if (sngay != L"")
		{
			if (ichk7 == 1 || ichkCN == 1)
			{
				/*// Nghỉ cuối tuần
				if(ichk7==0 && ichkCN==1)
					szval.Format(L"=IFERROR(CHOOSE(WEEKDAY(%s),\"CN\"),\"\")",GAR(ir,ic,pr1,0));
				else if(ichk7==1 && ichkCN==0)
					szval.Format(
						L"=CHOOSE(WEEKDAY(%s),\"\",\"\",\"\",\"\",\"\",\"\",\"T7\")",
						GAR(ir,ic,pr1,0));
				else if(ichk7==1 && ichkCN==1)
					szval.Format(
						L"=CHOOSE(WEEKDAY(%s),\"CN\",\"\",\"\",\"\",\"\",\"\",\"T7\")",
						GAR(ir,ic,pr1,0));

				pr1->PutItem(1,99,(_bstr_t)szval);
				szval = GIOR(1,99,pr1,L"");*/

				sthumay = fThuMay(sngay);
				if ((ichk7 == 1 && sthumay == L"T7") || (ichkCN == 1 && sthumay == L"CN"))
				{
					_iktra = 1;
					PRS->Interior->PutColor(clrKT_Bg);
					PRS->Font->PutColor(clrKT_Txt);
				}
			}

			if (_iktra == 0)
			{
				// Kiểm tra tiếp xem có phải ngày nghỉ lễ tết không?
				if (gtri > 0)
				{
					for (int k = 1; k <= gtri; k++)
					{
						if (sngay == sNgayle[k])
						{
							_iktra = 1;
							PRS->Interior->PutColor(clrKT_Bg);
							PRS->Font->PutColor(clrKT_Txt);
							if (ichkCmt == 1 && sGchu[k] != L"") PRS->AddComment((_bstr_t)sGchu[k]);
							break;
						}
					}
				}
			}

			if (_iktra == 0)
			{
				PRS->Interior->PutColor(xlNone);
				PRS->Font->PutColor(RGB(0, 0, 0));
			}
		}

		//pr1->PutItem(1,99,(_bstr_t)L"");
	}
	catch(...){_s(L"Lỗi kiểm tra ngày nghỉ lễ.");}	
}


long Dlg_Timer::f_Read_Nhatky()
{
	try
	{
		szCheckNgay=L"";
		if(getPathNHKY()==L"") return 0;
		if(connectDsn(pathNK)==-1) return 0;
		
		CConnection ObjConn;
		ObjConn.OpenAccessDB(pathNK, L"", L"");
		CRecordset* Recset = ObjConn.Execute(L"SELECT * FROM tbBoqua;");

		int dem=0;
		int icheck=0;
		CString szval=L"";
		while( !Recset->IsEOF() )
		{
			icheck=0;
			Recset->GetFieldValue(L"ngayghi",szval);

			if(szCheckNgay!=L"")
			{
				if(szCheckNgay.Find(szval)>=0) icheck=1;
			}

			if(icheck==0)
			{
				dem++;
				sNgayle[dem] = szval;
				Recset->GetFieldValue(L"ghichu",sGchu[dem]);

				szCheckNgay+=sNgayle[dem];
				szCheckNgay+=L"|";
			}

			Recset->MoveNext();
		}
		Recset->Close();
		ObjConn.CloseDatabase();
		
		return dem;
	}
	catch(...){}
	return 0;
}


long Dlg_Timer::f_Read_file_CSV(CString pathCSV, long num)
{
	int icheck=0;
	long cong=num;

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
			_s(L"Không thể mở được file dữ liệu. Vui lòng chọn lại file dữ liệu.");
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

			int fdem=0;
			CString ft[2]={L"",L""};

			for(int i= 1;i<numberOfLine;i++)
			{
				wcscpy_s(*(dArray+i), wcslen(*(p+i))+1, *(p+i));
				wchar_t *textString = new wchar_t[SIZE_LINE];
				wcscpy_s(textString,wcslen(*(dArray+i))+1, *(dArray+i));
				szVarString= textString;

				long length;
				CString szLeft;
				int nToken= szVarString.Find(L'\t');

				fdem=0;
				while(nToken>=0)
				{
					length= szVarString.GetLength();
					szLeft= szVarString.Left(nToken);
					szLeft.TrimLeft();szLeft.TrimRight();
					szVarString=szVarString.Right(length-nToken-1);
					nToken= szVarString.Find(L'\t');

					fdem++;
					if(fdem==1) ft[0]=szLeft;

					// Thêm vào để lấy cột cuối cùng (by Leo 23.06.15)
					ft[1]=L"";
					if(nToken<0) ft[1]=szVarString;
				}

				ft[0].Trim();
				ft[0].Replace(L" ",L"");
				ft[0].Replace(L"N",L"");
				if(ft[0]!=L"")
				{
					icheck=0;
					if(szCheckNgay!=L"")
					{
						if(szCheckNgay.Find(ft[0])>=0) icheck=1;
					}

					if(icheck==0)
					{
						cong++;
						sNgayle[cong] = ft[0];
						sGchu[cong] = ft[1];
						sGchu[cong].Trim();
					}
				}
				free(textString);
			}
			fclose(pReadFile);
		}
	
		//free memory
		for (i = 0; i < _size; i++)
		{
			free(*(dArray+i));
			free(*(p+i));
		}
		free(dArray);
		free(p);
	}
	catch(_com_error & error){_s(L"#QL4.126: Lỗi đọc dữ liệu ngày tháng nghỉ lễ.");}

	return cong;
}


long Dlg_Timer::f_GetLineFile(CString szPath)
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


int Dlg_Timer::f_SetGtriDate(RangePtr pRx,CString csdate,int iR,int iC,int iTG,int type)
{
	try
	{
		CString sTp[2]={L"sau",L"trước"} ;
		CString val = GIOR(iR,iTG,pRx,L"");
		val.Trim();

		if(val!=L"" && val!=csdate)
		{
			if(compare_date(_iNgay,val,csdate)==type)
			{
				CString szval=L"";
				szval.Format(L"Bạn đã nhập vào ngày [%s]."
					L"\nGiá trị ngày nhập phải %s ngày %s",csdate,sTp[type],val);
				MessageBox(szval,L"Giá trị nhập không hợp lệ",MB_OK | MB_ICONWARNING);
				f_ktra_date(_bNThang,pRx,iR,iC);
				return 0;
			}
		}
	}
	catch(...){}

	return 1;
}



void Dlg_Timer::f_set_date()
{
	try
	{
		CString sn=L"";
		if(_iTimeDlg>0) m_timeDlg[_iTimeDlg-1].GetWindowTextW(sn);
		if(sn!=L"") _bNThang = sn;
		if(_bNThang!=L"")
		{
			int d,m,y;
			this->SetWindowTextW(L"Ngày " + _bNThang);
			d = _wtoi(_bNThang.Left(2));
			m = _wtoi(_bNThang.Mid(3,2));
			y = _wtoi(_bNThang.Right(4));
			COleDateTime now(y, m, d, 0, 0, 0);
			cmc1.SetCurSel(now);
		}
		else
		{
			this->SetWindowTextW(L"Lựa chọn ngày tháng");
			COleDateTime now = COleDateTime::GetCurrentTime();			
			cmc1.SetCurSel(now);
		}
	}
	catch(...)
	{
		this->SetWindowTextW(L"Lựa chọn ngày tháng");
		COleDateTime now = COleDateTime::GetCurrentTime();
		cmc1.SetCurSel(now);
	}
}


void Dlg_Timer::OnMcnSelectMc1(NMHDR *pNMHDR, LRESULT *pResult)
{
	try
	{
		// Tạo trạng thái kích đúp bằng cách xác định khoảng thời gian kích chuột 2 lần tại cùng vị trí (ngày)
		// xem thời gian đó có nhỏ hơn GetDoubleClickTime() = 500ms không?
		// Nếu nhỏ hơn thì coi như là click đúp
		// Ở đây ta set timer cho biến đếm iDbclick++ --> Tại OnInitDialog settimer = 800ms
		// Mỗi lần timer chạy iDbclick=0, do đó trong 1 khoảng thời gian định sẵn (800ms)
		// Nếu iDbclick=2 thì ta coi như đó là click đúp

		iDbclick++;

		*pResult = 0;
		UpdateData();
		CString szval=L"";
		CTime dateTime;
		cmc1.GetCurSel(dateTime);
		szval.Format(L"Ngày %d/%d/%d",
				(int)dateTime.GetDay(),(int)dateTime.GetMonth(),(int)dateTime.GetYear());
		this->SetWindowTextW(szval);
		UpdateData(FALSE);

		if(iDbclick<=2) sDbcl[iDbclick] = szval;
		if(iDbclick==2 && sDbcl[1]==sDbcl[2]) OnBnClickedB1();
	}
	catch(...){_s(L"Lỗi kích đúp chọn ngày");}
}

