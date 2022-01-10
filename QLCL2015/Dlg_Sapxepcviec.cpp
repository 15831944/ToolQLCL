#include "Dlg_Sapxepcviec.h"

IMPLEMENT_DYNAMIC(Dlg_Sapxepcviec, CDialog)

Dlg_Sapxepcviec::Dlg_Sapxepcviec(CWnd* pParent /*=NULL*/)
	: CDialog(Dlg_Sapxepcviec::IDD, pParent)
{

}

Dlg_Sapxepcviec::~Dlg_Sapxepcviec()
{
}

void Dlg_Sapxepcviec::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, SCV_CB1, m_cbb1);
	DDX_Control(pDX, SCV_CB2, m_cbb2);
	DDX_Control(pDX, SCV_CB3, m_cbb3);
	DDX_Control(pDX, SCV_B1, m_btn);
	DDX_Control(pDX, SCV_CK1, m_chk);
}

BEGIN_MESSAGE_MAP(Dlg_Sapxepcviec, CDialog)
	ON_BN_CLICKED(SCV_B2, &Dlg_Sapxepcviec::OnBnClickedB2)
	ON_BN_CLICKED(SCV_B1, &Dlg_Sapxepcviec::OnBnClickedB1)
	ON_CBN_SELCHANGE(SCV_CB2, &Dlg_Sapxepcviec::OnCbnSelchangeCb2)
	ON_CBN_SELCHANGE(SCV_CB3, &Dlg_Sapxepcviec::OnCbnSelchangeCb3)
	ON_CBN_SELCHANGE(SCV_CB1, &Dlg_Sapxepcviec::OnCbnSelchangeCb1)
END_MESSAGE_MAP()

// Dlg_Sapxepcviec message handlers
BOOL Dlg_Sapxepcviec::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	// Xác định sheet hiện hành
	pSheet = pWb->ActiveSheet;
	_bsCNM = pSheet->CodeName;	// shHSNTVL hoặc shHSNTCV
	_bsName = pSheet->Name;
	sNameSh = (LPCTSTR)_bsName;
	pRange = pSheet->Cells;

	// Thiết lập mặc định	
	m_chk.SetCheck(iAtoSTT);

	// Load dữ liệu combobox
	f_Load_combobx();

	// Xác định số lượng và vị trí từng hạng mục
	f_xac_dinh_HM();

	return TRUE;
}


void Dlg_Sapxepcviec::f_Load_combobx()
{
	// Ưu tiên 1	
	if (_bsCNM == (_bstr_t)L"shHSNTVL")
	{
		m_cbb1.SetCueBanner(L"Kích chọn kiểu sắp xếp...");
		m_cbb1.AddString(L"Ngày nhập kho");
		m_cbb1.AddString(L"Ngày lấy mẫu thi nghiệm");
		m_cbb1.AddString(L"Ngày nghiệm thu nội bộ");
		m_cbb1.AddString(L"Ngày ghi phiếu yêu cầu");
		m_cbb1.AddString(L"Ngày nghiệm thu vật liệu");

		// iKieuSTT = kiểu vật liệu
		if (iKieuSTT >= 0) m_cbb1.SetCurSel(iKieuSTT);
		else m_btn.EnableWindow(0);
	}
	else
	{
		m_cbb1.SetCueBanner(L"Kích chọn kiểu sắp xếp...");
		m_cbb1.AddString(L"Ngày lấy mẫu thí nghiệm");
		m_cbb1.AddString(L"Ngày nghiệm thu nội bộ");
		m_cbb1.AddString(L"Ngày ghi phiếu yêu cầu");
		m_cbb1.AddString(L"Ngày nghiệm thu công việc");
		m_cbb1.AddString(L"Ngày bắt đầu nhật ký");
		m_cbb1.AddString(L"Ngày kiểm tra điều kiện bê tông");
		m_cbb1.AddString(L"Ngày theo dõi bê tông");

		// iKieuST2 = kiểu công việc
		if (iKieuST2 >= 0) m_cbb1.SetCurSel(iKieuST2);
		else m_btn.EnableWindow(0);
	}
}


void Dlg_Sapxepcviec::xl_xdcot_DMCV(_WorksheetPtr pSh)
{
	if (_bsCNM == (_bstr_t)L"shHSNTVL")
	{
		_iCot[1] = getColumn(pSh, "DMVL_STT");
		_iCot[2] = getColumn(pSh, "DMVL_MAVL");
		_iCot[4] = getColumn(pSh, "DMVL_MAHS");
		_iCot[5] = getColumn(pSh, "DMVL_ND");
		_iCot[8] = getColumn(pSh, "DMVL_TC");
		_iCot[21] = getColumn(pSh, "DMVL_KBB");
		_iCot[26] = _iCot[21];

		iNgay[0] = getColumn(pSh, "DMVL_NK_NG");
		iNgay[1] = getColumn(pSh, "DMVL_LM_NG");
		iNgay[2] = getColumn(pSh, "DMVL_NTNB_NG");
		iNgay[3] = getColumn(pSh, "DMVL_YC");
		iNgay[4] = getColumn(pSh, "DMVL_AB_NG");
		iNgay[5] = iNgay[4];
		iNgay[6] = iNgay[4];
	}
	else
	{
		_iCot[1] = getColumn(pSh, "DMBB_STT");
		_iCot[2] = getColumn(pSh, "DMBB_MACV");
		_iCot[4] = getColumn(pSh, "DMBB_MAHS");
		_iCot[5] = getColumn(pSh, "DMBB_ND");
		_iCot[8] = getColumn(pSh, "DMBB_TC");
		_iCot[21] = getColumn(pSh, "DMBB_PS");
		_iCot[26] = getColumn(pSh, "DMBB_TDDBT_GIO");

		iNgay[0] = getColumn(pSh, "DMBB_TN_NGAY");
		iNgay[1] = getColumn(pSh, "DMBB_NB_NGAY");
		iNgay[2] = getColumn(pSh, "DMBB_PHIEUYC");
		iNgay[3] = getColumn(pSh, "DMBB_AB_NGAY");
		iNgay[4] = getColumn(pSh, "DMBB_HK_NGAYBD");
		iNgay[5] = getColumn(pSh, "DMBB_DKDBT_NGAY");
		iNgay[6] = getColumn(pSh, "DMBB_TDDBT_NGAY");
	}
}


void Dlg_Sapxepcviec::f_xac_dinh_HM()
{
	try
	{
		// Xác định các cột sheet danh mục công việc
		xl_xdcot_DMCV(pSheet);

		// Vùng sắp xếp
		m_cbb2.AddString(L"Tất cả các hạng mục");

		CString szval = L"";
		iEnd = FindComment(pSheet, _iCot[1], "END");
		ibd = 8, ikt = iEnd;

		int dem = 0;
		for (int i = 8; i < iEnd; i++)
		{
			// Xác định hạng mục
			szval = GIOR(i, _iCot[2], pRange, L"");
			if (szval == L"HM")
			{
				dem++;
				Msrt[dem] = i;
				szval = GIOR(i, _iCot[5], pRange, L"");
				m_cbb2.AddString(szval);
			}
		}

		Msrt[dem + 1] = iEnd;
		m_cbb2.SetCurSel(0);

		// Không tồn tại hạng mục
		if (dem == 0) m_cbb2.EnableWindow(0);

		// Xác định các giai đoạn
		f_xacdinh_gdoan(ibd, ikt);
	}
	catch (...) { _s(L"Lỗi xác định hạng mục"); }
}


void Dlg_Sapxepcviec::f_xacdinh_gdoan(int num1, int num2)
{
	try
	{
		m_cbb3.ResetContent();
		ASSERT(m_cbb3.GetCount() == 0);
		m_cbb3.AddString(L"Tất cả các giai đoạn");

		int dem = 0;
		CString szval = L"";
		for (int i = num1 + 1; i < num2; i++)
		{
			szval = GIOR(i, _iCot[2], pRange, L"");
			if (szval == L"GD")
			{
				dem++;
				Msgd[dem] = i;
				szval = GIOR(i, _iCot[5], pRange, L"");
				m_cbb3.AddString(szval);
			}
		}

		Msgd[dem + 1] = ikt;
		m_cbb3.SetCurSel(0);
	}
	catch (...) { _s(L"Lỗi xác định giai đoạn"); }
}


void Dlg_Sapxepcviec::OnBnClickedB2()
{
	CDialog::OnCancel();
}


void Dlg_Sapxepcviec::OnBnClickedB1()
{
	try
	{
		iAtoSTT = m_chk.GetCheck();
		int num = m_cbb1.GetCurSel();
		if (num == -1)
		{
			_s(L"Vui lòng kích chọn kiểu sắp xếp!");
			return;
		}

		/////////////////////////////////////////////////////////////////
				// Bước 1: Xác định các vị trí có chứa công tác & sắp xếp
		if (_bsCNM == (_bstr_t)L"shHSNTVL") iKieuSTT = num;
		else  iKieuST2 = num;

		int dem = 0;
		CString szval = L"";
		for (int i = ibd; i < ikt; i++)
		{
			szval = GIOR(i, _iCot[2], pRange, L"");
			if (szval == L"HM" || szval == L"GD")
			{
				szval = GIOR(i + 1, _iCot[2], pRange, L"");
				if (szval != L"HM" && szval != L"GD")
				{
					dem++;
					Msrt[dem] = i;
				}
			}
		}

		// Chốt cuối cùng
		Msrt[dem + 1] = ikt;

		// Tồn tại hạng mục hoặc giai đoạn thì sắp xếp
		CDialog::OnOK();
		if (dem >= 1)
		{
			xl->PutScreenUpdating(VARIANT_FALSE);
			xl->StatusBar = (_bstr_t)L"Đang tiến hành sắp xếp dữ liệu. Vui lòng đợi trong giây lát...";

			// Bước 2: Chèn thêm 1 cột vào trước cột A		
			RangePtr PRS = pRange->GetItem(1, 1);
			PRS->EntireColumn->Insert(xlToRight);

			// Tăng vị trí cột lên 1 đơn vị
			_iCot[1]++;	_iCot[2]++;	_iCot[4]++;	_iCot[5]++;	_iCot[8]++;	_iCot[21]++; _iCot[26]++;
			for (int i = 0; i <= 6; i++) iNgay[i]++;

			// Bước 3: Sắp xếp dữ liệu (Quan trọng)
			for (int i = 1; i <= dem; i++) f_Sap_xep(Msrt[i], Msrt[i + 1], iNgay[num]);

			// Bước 4: Tạo Autofilter cho dòng tiêu đề mặc định
			int ktra = pSheet->GetAutoFilterMode();
			if (ktra == -1) pSheet->PutAutoFilterMode(FALSE);
			PRS = pSheet->GetRange(pRange->GetItem(7, _iCot[1]), pRange->GetItem(7, _iCot[26]));
			PRS->EntireRow->AutoFilter(1, vtMissing, XlAutoFilterOperator::xlAnd, vtMissing, TRUE);

			// Bước 5: Xóa cột A vừa thêm vào
			PRS = pRange->GetItem(1, 1);
			PRS->EntireColumn->Delete(xlToLeft);

			// Bước 6: Đánh lại stt
			_iCot[1]--; _iCot[2]--; _iCot[4]--;
			int ntype = _wtoi(getVCell(psTS, L"TS_ATOSTT"));
			if (ntype != 0 && ntype != 1 && ntype != 2) ntype = 0;
			f_Auto_stt_dmuc(pRange, ntype, iAtoSTT, 8, iEnd, _iCot[1], _iCot[2], _iCot[4]);
			xl->StatusBar = (_bstr_t)L"Ready";
		}

		/////////////////////////////////////////////////////////////////

	}
	catch (...) {}
}


void Dlg_Sapxepcviec::f_Sap_xep(int iNum1, int iNum2, int iCSort)
{
	try
	{
		// Kiểm tra iNum2 có phải vị trí END không? --> Xác định lại iNum2
		if (iNum2 == iEnd)
		{
			CString s5 = L"", s8 = L"";
			for (int i = iEnd - 1; i >= iNum1 + 1; i--)
			{
				s5 = GIOR(i, _iCot[5], pRange, L"");
				s8 = GIOR(i, _iCot[8], pRange, L"");
				if (s5 != L"" || s8 != L"")
				{
					iNum2 = i + 1;
					break;
				}
			}
		}

		// Đổ công thức
		CString szval = L"";
		szval.Format(L"=IF(%s<>\"\",TEXT(%s,\"General\"),%s)",
			GAR(iNum1 + 1, _iCot[1], pRange, 0), GAR(iNum1 + 1, iCSort, pRange, 0), GAR(iNum1, 1, pRange, 0));
		pRange->PutItem(iNum1 + 1, 1, (_bstr_t)szval);

		// Sao chép công thức
		RangePtr pRselect = pSheet->GetRange(pRange->GetItem(iNum1 + 1, 1), pRange->GetItem(iNum2 - 1, 1));
		PRS = pRange->GetItem(iNum1 + 1, 1);
		PRS->AutoFill(pRselect, XlAutoFillType::xlFillDefault);

		// Kiểm tra sort
		int ktra = pSheet->GetAutoFilterMode();
		if (ktra == -1) pSheet->PutAutoFilterMode(FALSE);

		// Call xla (Quan trọng)
		CString sRsort = L"", sRange = L"";
		sRsort.Format(L"'%d:%d", iNum1, iNum1);
		sRange.Format(L"%s", GAR(iNum1, 1, pRange, 0));
		f_Sort_danhmuc(sNameSh, sRsort, sRange);

		// Xóa công thức
		PRS = pRange->GetItem(1, 1);
		PRS->EntireColumn->ClearContents();
	}
	catch (...) { _s(L"Lỗi sắp xếp dữ liệu"); }
}

void Dlg_Sapxepcviec::OnCbnSelchangeCb2()
{
	try
	{
		int num = m_cbb2.GetCurSel();
		if (num == 0)
		{
			ibd = 8;
			ikt = iEnd;
		}
		else
		{
			ibd = Msrt[num];
			ikt = Msrt[num + 1];
		}

		// Xác định các giai đoạn tương ứng
		f_xacdinh_gdoan(ibd, ikt);
	}
	catch (...) {}
}


void Dlg_Sapxepcviec::OnCbnSelchangeCb3()
{
	int num = m_cbb3.GetCurSel();
	if (num == 0)
	{
		num = m_cbb2.GetCurSel();
		if (num == 0)
		{
			ibd = 8;
			ikt = iEnd;
		}
		else
		{
			ibd = Msrt[num];
			ikt = Msrt[num + 1];
		}
	}
	else
	{
		ibd = Msgd[num];
		ikt = Msgd[num + 1];
	}
}


void Dlg_Sapxepcviec::OnCbnSelchangeCb1()
{
	m_btn.EnableWindow(1);
}
