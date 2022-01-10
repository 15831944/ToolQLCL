// Dg_In_ho_so message handlers
#include "Dg_In_ho_so.h"
#include "Dlg_Setup_Print.h"
#include "Dlg_Cpys.h"
#include "Dlg_luu_Hso.h"
#include "Dlg_sapxepIn.h"

Dg_In_ho_so::Dg_In_ho_so(CWnd* pParent /*=NULL*/)
	: CDialog(Dg_In_ho_so::IDD, pParent)
{
	
}

void Dg_In_ho_so::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IHS_C1, cbb[1]);
	DDX_Control(pDX, IHS_C2, cbb[2]);
	DDX_Control(pDX, IHS_C3, cbb[3]);
	DDX_Control(pDX, IHS_C4, cbb[4]);
	DDX_Control(pDX, IHS_T1, txt[1]);
	DDX_Control(pDX, IHS_T2, txt[2]);
	DDX_Control(pDX, IHS_CK1, chk[1]);
	DDX_Control(pDX, IHS_CK2, chk[2]);
	DDX_Control(pDX, IHS_CK3, chk[3]);
	DDX_Control(pDX, IHS_CK4, chk[4]);
	DDX_Control(pDX, IHS_CK5, chk[5]);
	DDX_Control(pDX, IHS_CK6, chk[6]);
	DDX_Control(pDX, IHS_CK7, chk[7]);
	DDX_Control(pDX, IHS_CK8, chk[8]);
	DDX_Control(pDX, IHS_B1, btn[1]);
	DDX_Control(pDX, IHS_B2, btn[2]);
	DDX_Control(pDX, IHS_B3, btn[3]);
	DDX_Control(pDX, IHS_L1, list1);
	DDX_Control(pDX, IHS_S2, kqgrp);
	DDX_Control(pDX, IHS_B5, btn55);
	DDX_Control(pDX, IHS_B7, btn77);
	DDX_Control(pDX, IHS_ST99, sttHuongdan);
}

BEGIN_MESSAGE_MAP(Dg_In_ho_so, CDialog)
	ON_WM_SIZE()
	ON_WM_SIZING()
	ON_WM_TIMER()
	ON_BN_CLICKED(IHS_B3, &Dg_In_ho_so::OnBnClickedB3)
	ON_BN_CLICKED(IHS_B2, &Dg_In_ho_so::OnBnClickedB2)
	ON_BN_CLICKED(IHS_B1, &Dg_In_ho_so::OnBnClickedB1)
	ON_CBN_SELCHANGE(IHS_C1, &Dg_In_ho_so::OnCbnSelchangeC1)
	ON_CBN_SELCHANGE(IHS_C2, &Dg_In_ho_so::OnCbnSelchangeC2)
	ON_CBN_SELCHANGE(IHS_C3, &Dg_In_ho_so::OnCbnSelchangeC3)
	ON_CBN_SELCHANGE(IHS_C4, &Dg_In_ho_so::OnCbnSelchangeC4)
	ON_BN_CLICKED(IHS_CK1, &Dg_In_ho_so::OnBnClickedCk1)
	ON_BN_CLICKED(IHS_CK2, &Dg_In_ho_so::OnBnClickedCk2)
	ON_NOTIFY(NM_DBLCLK, IHS_L1, &Dg_In_ho_so::OnNMDblclkL1)
	ON_NOTIFY(LVN_KEYDOWN, IHS_L1, &Dg_In_ho_so::OnLvnKeydownL1)
	ON_BN_CLICKED(IHS_CK3, &Dg_In_ho_so::OnBnClickedCk3)
	ON_EN_KILLFOCUS(IHS_T2, &Dg_In_ho_so::OnEnKillfocusT2)
	ON_BN_CLICKED(IHS_B4, &Dg_In_ho_so::OnBnClickedB4)
	ON_BN_CLICKED(IHS_CK5, &Dg_In_ho_so::OnBnClickedCk5)
	ON_BN_CLICKED(IHS_CK6, &Dg_In_ho_so::OnBnClickedCk6)
	ON_BN_CLICKED(IHS_CK7, &Dg_In_ho_so::OnBnClickedCk7)
	ON_BN_CLICKED(IHS_B5, &Dg_In_ho_so::OnBnClickedB5)	
	ON_BN_CLICKED(IHS_B6, &Dg_In_ho_so::OnBnClickedB6)
	ON_BN_CLICKED(IHS_CK4, &Dg_In_ho_so::OnBnClickedCk4)
	ON_BN_CLICKED(IHS_B7, &Dg_In_ho_so::OnBnClickedB7)	
	ON_BN_CLICKED(IHS_CK8, &Dg_In_ho_so::OnBnClickedCk8)	
END_MESSAGE_MAP()

BEGIN_EASYSIZE_MAP(Dg_In_ho_so)
	// EASYSIZE(control,left,top,right,bottom,options)
	// options = 1 -> không thay đổi, =0 có thay đổi
	// ES_BORDER = ghim (thay đổi theo)
	// ES_KEEPSIZE = giữ nguyên ban đầu

	// Trái
	EASYSIZE(IHS_S1,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)	
	EASYSIZE(IHS_ST1,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)	
	EASYSIZE(IHS_ST2,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)	
	EASYSIZE(IHS_ST3,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)	
	EASYSIZE(IHS_ST4,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)	
	EASYSIZE(IHS_ST5,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(IHS_ST99,ES_BORDER,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)

	EASYSIZE(IHS_C1,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)	
	EASYSIZE(IHS_C2,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)	
	EASYSIZE(IHS_C3,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)

	EASYSIZE(IHS_CK8,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(IHS_CK1,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(IHS_CK2,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(IHS_CK3,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(IHS_CK4,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(IHS_CK5,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(IHS_CK6,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(IHS_CK7,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)

	EASYSIZE(IHS_B1,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(IHS_B2,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(IHS_B3,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(IHS_B4,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(IHS_B5,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(IHS_B7,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)

	EASYSIZE(IHS_T1,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(IHS_T2,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)

	// Phải
	EASYSIZE(IHS_S2,ES_BORDER,ES_BORDER,ES_BORDER,ES_BORDER,0)
	EASYSIZE(IHS_L1,ES_BORDER,ES_BORDER,ES_BORDER,ES_BORDER,0)
END_EASYSIZE_MAP


BOOL Dg_In_ho_so::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	iSapxepIN=0;
	iSec=0,iSetTm=500;
	SetTimer(1, iSetTm, NULL);

	f_Load_Hoso();
	f_Style_ListCV();
	f_KTLaymau_VL_CV();
	f_Dot_thanh_toanVL();
	f_Dot_thanh_toanCV();
	f_Dot_thanh_toanGD();
	f_Load_ToolTip();
	
	for (int i = 1; i <= 6; i++) iLuuChk[i]=0;
	iLuuChk[8]=0;

	chk[4].SetCheck(_ihidepic);
	iLuuMDP = _ihidepic;

	if(iSetup==0) 
	{
		f_thietlap_giatri();
		iSetup=1;
	}

	// Xác định số lượng biên bản VL,CV & GD
	f_xacdinh_sl_bban();

	// Xác định số bảng tính bổ sung
	f_Load_all_sheet();

	for (int i = 1; i <=3; i++) igt[i]=0;
	txt[1].SetWindowText(L"");
	txt[2].SetWindowText(L"1");
	txt[1].EnableWindow(FALSE);
	cbb[2].EnableWindow(FALSE);
	cbb[3].EnableWindow(FALSE);
	list1.EnableWindow(FALSE);

	sttHuongdan.SubclassDlgItem(IHS_ST99,this);
	sttHuongdan.SetTextColor(BLUE);

	// Lưu hồ sơ (Dlg_luu_Hso)
	luuhsCB1=0;
	luuhsCB2=0;

	if(isxbs>0) btn55.EnableWindow(1);

	// Set công việc lẻ (24.02.2018)
	if(jInle==1) _SetLecongiviec();

	INIT_EASYSIZE;

	return TRUE; 
}

// Bắt sự kiện click Enter
BOOL Dg_In_ho_so::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if (pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(IHS_L1))
	{
		f_Add_congviec();
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(IHS_B3))
	{
		OnBnClickedB3();
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(IHS_T1))
	{
		GotoDlgCtrl(GetDlgItem(IHS_T2));
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(IHS_T2))
	{
		GotoDlgCtrl(GetDlgItem(IHS_B1));
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		_prt0 = list1.GetColumnWidth(1);
		_prt1 = list1.GetColumnWidth(2);
		_prt2 = list1.GetColumnWidth(3);

		/*try
		{
			CoInitialize(NULL);
			xl.GetActiveObject(L"Excel.Application");
			xl->Visible = true;
			xl->EnableCancelKey = XlEnableCancelKey(FALSE);
			CoUninitialize();
		}
		catch(_com_error & error){}*/

		OnCancel();
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

void Dg_In_ho_so::OnSize(UINT nType, int cx, int cy) 
{
	CDialog::OnSize(nType, cx, cy);
	UPDATE_EASYSIZE;
}

void Dg_In_ho_so::OnSizing(UINT fwSide, LPRECT pRect) 
{
	CDialog::OnSizing(fwSide, pRect);
	EASYSIZE_MINSIZE(338,360,fwSide,pRect);
}


void Dg_In_ho_so::f_Load_ToolTip()
{
		
	CEdit * pbt0 = (CEdit *)GetDlgItem(IHS_T1);
	CListCtrl * ls1 = (CListCtrl *)GetDlgItem(IHS_L1);
	CButton * btn2 = (CButton *)GetDlgItem(IHS_B1);
	CComboBox * cbb4 = (CComboBox *)GetDlgItem(IHS_C3);
	CButton * chk5 = (CButton *)GetDlgItem(IHS_CK3);
	//CButton * btn55 = (CButton *)GetDlgItem(IHS_B5);
	//CButton * btn77 = (CButton *)GetDlgItem(IHS_B7);
		
	m_ToolTips.Create(this);
	m_ToolTips.SetMaxTipWidth(600);
	m_ToolTips.SetDelayTime(900);

	m_ToolTips.AddTool(pbt0, L"- Nhập chuỗi công việc liên tiếp: 1-100,8-16..."
		L"\n- Nhập các công việc lẻ: 1,3,5,8..."
		L"\n- Ngăn cách các công việc bằng dấu phẩy (,)");
		
	m_ToolTips.AddTool(ls1, L"- Kích đúp để chọn 1 công việc."
		L"\n- Kết hợp phím Ctrl(Shift) và ấn Enter(Space) để chọn nhiều.");

	m_ToolTips.AddTool(btn2, L"Kích để thiết lập nhiều tùy chọn hơn."); 

	m_ToolTips.AddTool(cbb4, L"Sử dụng khi hồ sơ công việc được chia làm nhiêu giai đoạn."); 
	m_ToolTips.AddTool(chk5, L"Tự động căn chỉnh dòng khi bị tràn trang in."); 

	//m_ToolTips.AddTool(btn55, L"Tính năng sử dụng khi tồn tại các bảng tính được sao chép (nhân bản).");
	//m_ToolTips.AddTool(btn77, L"Cho phép sắp xếp lại thứ tự các biên bản in trước khi in."); 
}


// Bổ sung lệnh ngày 24.02.2018
void Dg_In_ho_so::_SetLecongiviec()
{
	try
	{
		if(sTukhoa==L"") return;
	
	
		_WorksheetPtr pSh1 = pWb->ActiveSheet;
		RangePtr pRg1 = pSh1->Cells;
		_bstr_t sh = pSh1->CodeName;

		int jtenHS = 0;
		if(sh==(_bstr_t)L"shHSNTVL") jtenHS=1;
		else if(sh==(_bstr_t)L"shHSNTCV") jtenHS=2;
		else if(sh==(_bstr_t)L"shHSNTGD") jtenHS=3;

		if(jtenHS>0)
		{
			cbb[1].SetCurSel(jtenHS);
			OnCbnSelchangeC1();

			txt[1].SetWindowText(sTukhoa);
		}		
	}
	catch(...){}
}


void Dg_In_ho_so::f_CheckNUM1(int a1,int a2)
{
	try
	{
		slInbs=0;
		_bstr_t bsVal = (_bstr_t)"";
		if(a1==1)
		{
			if(a2==1) bsVal = (_bstr_t)"shHSNTVLLAYM";
			else if(a2==2) bsVal = (_bstr_t)"shHSNTVLNBVL";
			else if(a2==3) bsVal = (_bstr_t)"shHSNTVLYVL";
			else if(a2==4) bsVal = (_bstr_t)"shHSNTVLNVL";
		}
		else if(a1==2)
		{
			if(a2==1) bsVal = (_bstr_t)"shHSNTCVLMTN";
			else if(a2==2) bsVal = (_bstr_t)"shHSNTCVNBCV";
			else if(a2==3) bsVal = (_bstr_t)"shHSNTCVYNCV";
			else if(a2==4) bsVal = (_bstr_t)"shHSNTCVNTCV";
			else if(a2==5) bsVal = (_bstr_t)"shHSNTCVNBCV1";
			else if(a2==6) bsVal = (_bstr_t)"shHSNTCVNBCV2";
		}
		else if(a1==3)
		{
			if(a2==1) bsVal = (_bstr_t)"shHSNTGDNTBGD1";
			else if(a2==2) bsVal = (_bstr_t)"shHSNTGDYCNT1";
			else if(a2==3) bsVal = (_bstr_t)"shHSNTGDNTGD1";
		}

		if(bsVal==(_bstr_t)"") return;

		// Kiểm tra xem có sheet nào tương ứng không?
		CString sGan=L"",val = (LPCTSTR)bsVal;
		int nLen = val.GetLength();
		for (int i = 1; i <= slcpy[a1]; i++)
		{
			sGan = (LPCTSTR)MCS[a1].sCN[i];
			if(val.Left(nLen) == sGan.Left(nLen))
			{
				slInbs++;
				csInbs[slInbs] = MCS[a1].sCN[i];
				csInnm[slInbs] = MCS[a1].sName[i];
				ctInbs[slInbs] = MCS[a1].iSLCV[i];
			}			
		}

		if(slInbs==0) btn55.EnableWindow(0);
		else btn55.EnableWindow(1);
	}
	catch(...){}
}


void Dg_In_ho_so::f_CheckBtn55(int a1)
{
	slInbs=0;
	if(a1==0)
	{
		// Gán dữ liệu
		for (int j = 1; j <=3; j++)
		{
			if(slcpy[j]>0)
			for (int i = 1; i <= slcpy[j]; i++)
			{
				slInbs++;
				csInbs[slInbs] = MCS[j].sCN[i];
				csInnm[slInbs] = MCS[j].sName[i];
				ctInbs[slInbs] = MCS[j].iSLCV[i];
			}
		}
	}
	else if(a1==4) btn55.EnableWindow(0);
	else
	{
		if(slcpy[a1]>0)
		for (int i = 1; i <= slcpy[a1]; i++)
		{
			slInbs++;
			csInbs[slInbs] = MCS[a1].sCN[i];
			csInnm[slInbs] = MCS[a1].sName[i];
			ctInbs[slInbs] = MCS[a1].iSLCV[i];
		}
	}

	if(slInbs==0) btn55.EnableWindow(0);
	else btn55.EnableWindow(1);
}


void Dg_In_ho_so::f_thietlap_giatri()
{
	_myPrint.ileft = 0;
	_myPrint.itop = 0;
	_myPrint.iright = 0;
	_myPrint.ibottom = 0;
	_myPrint.ihtop = 0;
	_myPrint.ifbot = 0;

	_myPrint.ihor = 1;
	_myPrint.iver = 0;

	_myPrint.isizehd = 12;
	_myPrint.isizeft = 12;
	_myPrint.istylehd = 0;
	_myPrint.istyleft = 0;
	_myPrint.isechd = 0;
	_myPrint.isecft = 0;
	_myPrint.isecpage = 1;
	_myPrint.inextpage = 0;

	_myPrint.ibreaks = 0;
	_myPrint.ibwhite = 1;
	_myPrint.ierror = 1;

	_myPrint.izoom = 0;

	_myPrint.ctxthd = L"";
	_myPrint.ctxtft = L"";

	_myPrint.iprivate=1;
}


// Name ngắt dòng các biên bản
//PR_LMVL2		PR_LMVL
//PR_NTNBVL2	PR_NTNBVL
//PR_YCVL2		PR_YCVL
//PR_NTVL2		PR_NTVL
//
//PR_LMTN2		PR_LMTN
//PR_BBKT2		PR_BBKT
//PR_BBTD2		PR_BBTD
//PR_NTNB2		PR_NTNB
//PR_YCCV2		PR_YCCV
//PR_NTCV2		PR_NTCV
//
//PR_NTNBGD2	PR_NTNBGD
//PR_YCGD2		PR_YCGD
//PR_NTGD2		PR_NTGD	

void Dg_In_ho_so::OnBnClickedB3()
{
	try
	{
		// Hàm check bản quyền
		if(f_Check_ban_quyen()!=1) return;

		CString szval=L"";
		luuhsCB2=0;
		luuhsCB1 = cbb[1].GetCurSel();
		if(luuhsCB1>0 && luuhsCB1<4)
		{
			luuhsCB2 = cbb[2].GetCurSel();
			txt[1].GetWindowTextW(szval);
			szval.Trim();
			if(szval==L"")
			{
				_s(L"Vui lòng lựa chọn công việc cần in");
				return;
			}
		}

		iLuuKEY=0;
		iKieuIn = 0;	// Kiểu in
		erp = 0;

		// In hồ sơ
		igt[1] = cbb[1].GetCurSel();
		cbb[1].GetLBText(igt[1],fchon[1]);

		if(igt[1]>0)
		{
			igt[2] = cbb[2].GetCurSel();
			cbb[2].GetLBText(igt[2],fchon[2]);
		}
		else fchon[2] = L"NULL";

		txt[1].GetWindowTextW(fchon[4]);		
		if(fchon[4].GetLength()>100) fchon[4] = fchon[4].Left(99) + L",...";

		txt[2].GetWindowTextW(fchon[5]);
		fchon[5].Trim();
		iSLBanIn=1;
		if(fchon[5]!=L"") iSLBanIn = _wtoi(fchon[5]);
		if(iSLBanIn<=0) iSLBanIn=1;

		fchon[6]  = L"  + Nghiệm thu nội bộ vật liệu.";
		fchon[7]  = L"  + Nghiệm thu nội bộ công việc.";
		fchon[10] = L"  + Kiểm tra và theo dõi bê tông.";
		fchon[12] = L"  + Yêu cầu nghiệm thu vật liệu.";
		fchon[13] = L"  + Yêu cầu nghiệm thu công việc.";

		int iTotal = iLuuChk[1]+iLuuChk[2]+iLuuChk[4]+iLuuChk[5]+iLuuChk[6];
		if(iTotal==0) fchon[11]=L"";
		else
		{
			if(iTotal>1) fchon[11] = L"- Không in các biên bản:";
			else fchon[11] = L"- Không in biên bản:";

			if(iLuuChk[1]==1) fchon[11].Format(L"%s \n%s",fchon[11],fchon[6]);
			if(iLuuChk[2]==1) fchon[11].Format(L"%s \n%s",fchon[11],fchon[7]);
			if(iLuuChk[4]==1) fchon[11].Format(L"%s \n%s",fchon[11],fchon[10]);
			if(iLuuChk[5]==1) fchon[11].Format(L"%s \n%s",fchon[11],fchon[12]);
			if(iLuuChk[6]==1) fchon[11].Format(L"%s \n%s",fchon[11],fchon[13]);
		}

		if(iLuuChk[8]==1)
		{
			if(fchon[11]==L"") fchon[11] = L"- Chỉ in biên bản vật liệu có lấy mẫu";
			else fchon[11].Format(L"%s \n%s",fchon[11],L"- Chỉ in biên bản vật liệu có lấy mẫu");
		}

		fchon[6] = L"- Tự động căn chỉnh dòng khi in.";
		fchon[7] = L"- Cho phép chèn hình ảnh vào nội dung hồ sơ.";

		if(iLuuChk[3]==1 && _ihidepic==1) fchon[8].Format(L"%s \n%s",fchon[6],fchon[7]);
		else if(iLuuChk[3]==1 && _ihidepic==0) fchon[8] = fchon[6];
		else if(iLuuChk[3]==0 && _ihidepic==1) fchon[8] = fchon[7];
		else fchon[8]=L"";

		if(fchon[11]!=L"" && fchon[8]!=L"") fchon[9].Format(L"%s \n%s",fchon[11],fchon[8]);
		else if(fchon[11]!=L"" && fchon[8]==L"") fchon[9] = fchon[11];
		else if(fchon[11]==L"" && fchon[8]!=L"") fchon[9] = fchon[8];
		else fchon[9] = L"";

		BOOLEAN f = FALSE;
		msg = L"";
		if(igt[1]==0)
		{
			msg = L"Bạn có chắc chắn in toàn bộ hồ sơ?";
			if(fchon[9]!=L"") msg.Format(L"%s \n- Số bản in = %s \n%s",msg,fchon[5],fchon[9]);
		}
		else if(igt[1]==4)
		{
			msg = L"Bạn có chắc chắn in toàn bộ danh mục?";
			if(fchon[9]!=L"") msg.Format(L"%s \n- Số bản in = %s \n%s",msg,fchon[5],fchon[9]);
		}
		else
		{
			iLuuKEY = f_Xac_dinh_sl_In();
			if(iLuuKEY>0)
			{
				msg.Format(L"Bạn sẽ in hồ sơ với các thông số dưới? \n"
					L"\n- Hồ sơ: %s \n- Biên bản: %s \n- Số công việc = %s \n- Số bản in = %s",
					fchon[1],fchon[2],fchon[4],fchon[5]);

				if(fchon[9]!=L"") msg.Format(L"%s \n%s",msg,fchon[9]);
			}
		}

		_bstr_t _prter = xl->Application->GetActivePrinter();
		msg.Format(L"%s \n- Máy in sử dụng: %s",msg,(LPCTSTR)_prter);
		msg.Format(L"%s \n* Chú ý: Thời gian in sẽ là tổng thời gian in từng biên bản và phụ thuộc tốc độ máy in.",msg);

		int result = MessageBox(msg,_T("Thông báo"), MB_YESNO | MB_ICONQUESTION);
		if (result == 6) f = TRUE;
		else f = FALSE;

		if(f==TRUE)
		{
			CDialog::OnOK();
			xl->PutScreenUpdating(VARIANT_FALSE);
			xl->StatusBar = (_bstr_t)L"Đang tiến hành in hồ sơ. Vui lòng chờ trong giây lát...";

			// Gán giá trị chèn ảnh
			shTS = name_sheet("shTS");
			psTS=xl->Sheets->GetItem(shTS);
			prTS=psTS->Cells;

			prTS->PutItem(getRow(psTS,"TS_PIC"),getColumn(psTS,"TS_PIC"),_ihidepic);

			_itrang = _myPrint.ifirst;		 // Trang bắt đầu khởi tạo
			if(igt[1]==0)
			{
				// In toàn bộ danh mục
				for (int i = 1; i <= iSLBanIn; i++)
					for (int j = 1; j <=4; j++) f_in_danh_muc_hs(j);

				// In toàn bộ hồ sơ
				In_toan_bo_ho_so();				
			}
			else if(igt[1]>0 && igt[1]<4)
			{
				// In toàn bộ 1 bộ hồ sơ nào đó (all biên bản)
				erp = 0;
				for(int i=1;i<=iSLBanIn;i++)
				{
					for(int j=1;j<=iLuuKEY;j++)
					{
						if(erp==1) return;
						f_Kiem_tra_inum(iCViec[j],igt[1],igt[2]);
					}
				}
			}
			else if(igt[1]==4)
			{
				// In toàn bộ danh mục
				if(igt[2]==0)
				{
					for (int i = 1; i <= iSLBanIn; i++)
						for (int j = 1; j <=4; j++) f_in_danh_muc_hs(j);
				}
				else
				{
					// In lẻ danh mục
					for (int i = 1; i <= iSLBanIn; i++) f_in_danh_muc_hs(igt[2]);
				}
			}
		
			// Trả lại giá trị ban đầu cho ảnh
			prTS->PutItem(getRow(psTS,"TS_PIC"),getColumn(psTS,"TS_PIC"),iLuuMDP);

			xl->StatusBar = (_bstr_t)L"Ready";
			xl->PutScreenUpdating(TRUE);
		}
	}
	catch(_com_error & error){_s(L"#QL4.01- Lỗi in hồ sơ 1.");}
}


void Dg_In_ho_so::OnBnClickedB2()
{
	// Hàm check bản quyền
	if(f_Check_ban_quyen()!=1) return;

	luuhsCB2=0;
	luuhsCB1 = cbb[1].GetCurSel();
	if(luuhsCB1>0 && luuhsCB1<4)
	{
		luuhsCB2 = cbb[2].GetCurSel();
		CString szval=L"";
		txt[1].GetWindowTextW(szval);
		szval.Trim();
		if(szval==L"")
		{
			_s(L"Vui lòng lựa chọn công việc cần in");
			return;
		}
	}

	iLuuKEY=0;
	iSLBanIn = 1;	// Số lượng bản in
	iKieuIn = 1;	// Kiểu xem trước
	erp = 0;

	// Xem trước hồ sơ
	igt[1] = cbb[1].GetCurSel();
	if(igt[1]>0) igt[2] = cbb[2].GetCurSel();
	else igt[2] = 0;

	try
	{
		if(igt[1]==0)
		{
			// Xem toàn bộ hồ sơ (istyle=1, iPrint=1)
			In_toan_bo_ho_so();

			// Xem toàn bộ danh mục
			for (int j = 1; j <=4; j++)
			{
				if(erp==1) return;
				f_in_danh_muc_hs(j);
			}
		}
		else if(igt[1]>0 && igt[1]<4) 
		{
			iLuuKEY = f_Xac_dinh_sl_In();

			erp = 0;
			for(int j=1;j<=iLuuKEY;j++)
			{
				if(erp==1) return;
				f_Kiem_tra_inum(iCViec[j],igt[1],igt[2]);
			}			
		}
		else if(igt[1]==4)
		{
			// Xem toàn bộ danh mục
			if(igt[2]==0)
			{
				for (int j = 1; j <=4; j++)
				{
					if(erp==1) return;
					f_in_danh_muc_hs(j);
				}
			}
			else f_in_danh_muc_hs(igt[2]);
		}
	}
	catch(_com_error & error){_s(L"#QL4.15- Lỗi xem trước hồ sơ.");}
}


void Dg_In_ho_so::OnBnClickedB1()
{
	iDlog=0;

	// Hiển thị hộp thoại tùy chọn
	AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
	Dlg_Setup_Print _dlg;
	_dlg.DoModal();
}


void Dg_In_ho_so::OnCbnSelchangeC1()
{
	slbsThuc=0;
	int gtri = cbb[1].GetCurSel();
	if(gtri>=0)	cbb[1].GetLBText(gtri,fchon[1]);

	cbb[4].ResetContent();
	ASSERT(cbb[4].GetCount() == 0);

	if(gtri==1)
	{
		if(slTTVL>0)
		{
			cbb[4].EnableWindow(1);
			cbb[4].AddString(L"Tất cả các đợt");
			for (int i = 1; i <= slTTVL; i++) cbb[4].AddString(sDotTT_VL[i]);
			cbb[4].SetCurSel(0);
		}
		else cbb[4].EnableWindow(0);
	}
	else if(gtri==2)
	{
		if(slTTCV>0)
		{
			cbb[4].EnableWindow(1);
			cbb[4].AddString(L"Tất cả các đợt");
			for (int i = 1; i <= slTTCV; i++) cbb[4].AddString(sDotTT_CV[i]);
			cbb[4].SetCurSel(0);
		}
		else cbb[4].EnableWindow(0);
	}
	else if(gtri==3)
	{
		if(slTTGD>0)
		{
			cbb[4].EnableWindow(1);
			cbb[4].AddString(L"Tất cả các đợt");
			for (int i = 1; i <= slTTGD; i++) cbb[4].AddString(sDotTT_GD[i]);
			cbb[4].SetCurSel(0);
		}
		else cbb[4].EnableWindow(0);
	}
	else cbb[4].EnableWindow(0);

	if(gtri==0)
	{
		kqgrp.SetWindowText(L"Lựa chọn công việc");

		cbb[2].ResetContent();
		ASSERT(cbb[2].GetCount() == 0);

		cbb[3].ResetContent();
		ASSERT(cbb[3].GetCount() == 0);

		txt[1].SetWindowText(L"");
		txt[1].EnableWindow(FALSE);
		cbb[2].EnableWindow(FALSE);
		cbb[3].EnableWindow(FALSE);
		btn[3].EnableWindow(TRUE);
		f_SetCheck_VLCV(1,1,1,1,1,1);

		f_Delete_ListCV();
		list1.EnableWindow(FALSE);
	}
	else
	{
		cbb[2].EnableWindow(TRUE);
		f_Load_Bienban();
	}
}


void Dg_In_ho_so::OnCbnSelchangeC2()
{
	slbsThuc=0;
	int n1 = cbb[1].GetCurSel();
	int n2 = cbb[2].GetCurSel();
	cbb[2].GetLBText(n1,fchon[2]);
	
	if(n1==1)
	{
		// Vật liệu
		if(n2==0) f_SetCheck_VLCV(1,0,0,1,0,1);			// all
		else if(n2==1) f_SetCheck_VLCV(0,0,0,0,0,1);	// LM
		else if(n2==2) f_SetCheck_VLCV(1,0,0,0,0,0);	// NB
		else if(n2==3) f_SetCheck_VLCV(0,0,0,1,0,0);	// YC
		else f_SetCheck_VLCV(0,0,0,0,0,0);
	}
	else if(n1==2)
	{
		// Công việc
		if(n2==0) f_SetCheck_VLCV(0,1,1,0,1,0);					// all
		else if(n2==2) f_SetCheck_VLCV(0,1,0,0,0,0);			// nb
		else if(n2==3) f_SetCheck_VLCV(0,0,0,0,1,0);			// yc
		else if(n2==5 || n2==6) f_SetCheck_VLCV(0,0,1,0,0,0);	// betong
		else f_SetCheck_VLCV(0,0,0,0,0,0);
	}
	else f_SetCheck_VLCV(0,0,0,0,0,0);	

	// Kiểm tra trường hợp chọn nghiệm thu công việc xây dựng
	if(n1==2)
	{
		if(n2==1 || n2==6) f_Xdinh_GiaiDoan(1);
		else if(n2==5) f_Xdinh_GiaiDoan(2);
		else f_Xdinh_GiaiDoan(0);
	}
}


void Dg_In_ho_so::OnCbnSelchangeC3()
{
	int gtri = cbb[3].GetCurSel();
	f_Load_GiaiDoan(gtri);
}


void Dg_In_ho_so::OnCbnSelchangeC4()
{
	f_Delete_ListCV();
	txt[1].SetWindowText(L"");

	CString szval=L"";
	int gtri = cbb[4].GetCurSel();	
	cbb[4].GetLBText(gtri,szval);
	if(gtri>0) f_Load_DotThanhtoan(szval);
}


void Dg_In_ho_so::f_Load_DotThanhtoan(CString sdotTT)
{
	try
	{
		int gtri = cbb[1].GetCurSel();
		if(gtri<1 || gtri>3) return;

		_bstr_t pWsh;
		_WorksheetPtr ps1;
		int iSTT,iMa,iND,iDOT,iEnd;

		if(gtri==1)
		{
			pWsh = name_sheet("shHSNTVL");
			ps1 = xl->Sheets->GetItem(pWsh);

			iSTT = getColumn(ps1,"DMVL_STT");
			iMa = getColumn(ps1,"DMVL_MAVL");
			iND = getColumn(ps1,"DMVL_ND");
			iDOT = getColumn(ps1,"DMVL_DOT");		
		}
		else if(gtri==2)
		{
			pWsh = name_sheet("shHSNTCV");
			ps1 = xl->Sheets->GetItem(pWsh);

			iSTT = getColumn(ps1,"DMBB_STT");
			iMa = getColumn(ps1,"DMBB_MACV");
			iND = getColumn(ps1,"DMBB_ND");
			iDOT = getColumn(ps1,"DMBB_DOT");
		}
		else if(gtri==3)
		{
			pWsh = name_sheet("shHSNTGD");
			ps1 = xl->Sheets->GetItem(pWsh);

			iSTT = getColumn(ps1,"DMGD_STT");
			iMa = getColumn(ps1,"DMGD_HS");
			iND = getColumn(ps1,"DMGD_ND");
			iDOT = getColumn(ps1,"DMGD_DOT");			
		}

		RangePtr pr1 = ps1->Cells;
		iEnd = FindComment(ps1,iSTT,"END");
		int dem=0;
		CString szval=L"";
		for (int i = 8; i < iEnd; i++)
		{
			szval = GIOR(i,iDOT,pr1,L"");
			szval.Trim();

			if(szval==sdotTT)
			{
				szval = GIOR(i,iSTT,pr1,L"");
				if(_wtoi(szval)>0)
				{
					list1.InsertItem(dem,L"",0);
					list1.SetItemText(dem,1,szval);

					szval = GIOR(i,iMa,pr1,L"");
					list1.SetItemText(dem,2,szval);

					szval = GIOR(i,iND,pr1,L"");
					list1.SetItemText(dem,3,szval);

					dem++;
				}					
			}				
		}

		msg.Format(L"Lựa chọn công việc (%d kết quả)",dem);
		kqgrp.SetWindowText(msg);
	}
	catch(...){}
}

void Dg_In_ho_so::OnBnClickedCk1()
{
	CButton *m_ctlCheck = (CButton*) GetDlgItem(IHS_CK1);
	if(m_ctlCheck->GetCheck() == BST_UNCHECKED) iLuuChk[1]=0;
	else iLuuChk[1]=1;
}


void Dg_In_ho_so::OnBnClickedCk2()
{
	CButton *m_ctlCheck = (CButton*) GetDlgItem(IHS_CK2);
	if(m_ctlCheck->GetCheck() == BST_UNCHECKED) iLuuChk[2]=0;
	else iLuuChk[2]=1;
}

void Dg_In_ho_so::OnBnClickedCk3()
{
	CButton *m_ctlCheck = (CButton*) GetDlgItem(IHS_CK3);
	if(m_ctlCheck->GetCheck() == BST_UNCHECKED) iLuuChk[3]=0;
	else iLuuChk[3]=1;
}

void Dg_In_ho_so::OnBnClickedCk5()
{
	CButton *m_ctlCheck = (CButton*) GetDlgItem(IHS_CK5);
	if(m_ctlCheck->GetCheck() == BST_UNCHECKED) iLuuChk[4]=0;
	else iLuuChk[4]=1;
}

void Dg_In_ho_so::f_Load_Hoso()
{
	cbb[1].ResetContent();
	ASSERT(cbb[1].GetCount() == 0);

	cbb[1].AddString(L"Tất cả hồ sơ");
	cbb[1].AddString(L"Nghiệm thu vật liệu");
	cbb[1].AddString(L"Nghiệm thu công việc xây dựng");
	cbb[1].AddString(L"NT hạng mục CT & giai đoạn TCXD");
	cbb[1].AddString(L"Các danh mục");
	cbb[1].SetCurSel(0);
}


void Dg_In_ho_so::f_Load_Bienban()
{
	cbb[2].ResetContent();
	ASSERT(cbb[2].GetCount() == 0);

	cbb[3].ResetContent();
	ASSERT(cbb[3].GetCount() == 0);

	int gtri = cbb[1].GetCurSel();
	if(gtri==1)
	{
		f_SetCheck_VLCV(1,0,0,1,0,1);
		cbb[3].EnableWindow(FALSE);
		if(slcv_VL==1)
		{
			txt[1].SetWindowText(L"1");
			f_Set_Duieu(1);
		}
		else if(slcv_VL>1)
		{
			msg.Format(L"1-%d",slcv_VL);
			txt[1].SetWindowText(msg);
			f_Set_Duieu(1);
		}
		else f_Set_Duieu(0);

		// Vật liệu
		cbb[2].AddString(L"Tất cả các biên bản");
		cbb[2].AddString(L"Lấy mẫu vật liệu");
		cbb[2].AddString(L"Nghiệm thu nội bộ vật liệu");
		cbb[2].AddString(L"Yêu cầu nghiệm thu vật liệu");
		cbb[2].AddString(L"Nghiệm thu vật liệu");
		list1.EnableWindow(TRUE);

		if(slcv_VL>0) f_Load_ListVLGD((_bstr_t)"shHSNTVL");
	}
	else if(gtri ==2)
	{
		f_SetCheck_VLCV(0,1,1,0,1,0);
		if(slcv_CV==1)
		{
			txt[1].SetWindowText(L"1");
			f_Set_Duieu(1);
		}
		else if(slcv_CV>1)
		{
			msg.Format(L"1-%d",slcv_CV);
			txt[1].SetWindowText(msg);
			f_Set_Duieu(1);
		}
		else f_Set_Duieu(0);

		// Công việc
		cbb[2].AddString(L"Tất cả các biên bản");
		cbb[2].AddString(L"Lấy mẫu thí nghiệm");
		cbb[2].AddString(L"Nghiệm thu nội bộ công việc");
		cbb[2].AddString(L"Yêu cầu nghiệm thu công việc");
		cbb[2].AddString(L"Nghiệm thu công việc");
		cbb[2].AddString(L"Biên bản kiểm tra bê tông");
		cbb[2].AddString(L"Biên bản theo dõi đổ bê tông");
		list1.EnableWindow(TRUE);
		
		if(slcv_CV>0)
		{
			cbb[3].EnableWindow(TRUE);
			f_Xdinh_GiaiDoan(0);
		}
		else cbb[3].EnableWindow(FALSE);

	}
	else if(gtri ==3)
	{
		f_SetCheck_VLCV(0,0,0,0,0,0);
		cbb[3].EnableWindow(FALSE);
		if(slcv_GD==1)
		{
			txt[1].SetWindowText(L"1");
			f_Set_Duieu(1);
		}
		else if(slcv_GD>1)
		{
			msg.Format(L"1-%d",slcv_GD);
			txt[1].SetWindowText(msg);
			f_Set_Duieu(1);
		}
		else f_Set_Duieu(0);

		// Giai đoạn
		cbb[2].AddString(L"Tất cả các biên bản");
		cbb[2].AddString(L"Nghiệm thu nội bộ giai đoạn");
		cbb[2].AddString(L"Yêu cầu nghiệm thu giai đoạn");
		cbb[2].AddString(L"Nghiệm thu giai đoạn");
		list1.EnableWindow(TRUE);

		if(slcv_GD>0) f_Load_ListVLGD((_bstr_t)"shHSNTGD");
	}
	else if(gtri ==4)
	{
		f_SetCheck_VLCV(0,0,0,0,0,0);
		cbb[3].EnableWindow(FALSE);
		txt[1].SetWindowText(L"");
		txt[1].EnableWindow(FALSE);
		btn[3].EnableWindow(TRUE);

		// Giai đoạn
		cbb[2].AddString(L"Tất cả các danh mục");
		cbb[2].AddString(L"Danh mục hồ sơ");
		cbb[2].AddString(L"Danh mục vật liệu");
		cbb[2].AddString(L"Danh mục công việc");
		cbb[2].AddString(L"Danh mục giai đoạn");

		f_Delete_ListCV();
		list1.EnableWindow(FALSE);
	}

	cbb[2].SetCurSel(0);

}


void Dg_In_ho_so::f_Xdinh_GiaiDoan(int istyle)
{
	try
	{
		_bstr_t pWsh = name_sheet("shHSNTCV");
		_WorksheetPtr pS1 = xl->Sheets->GetItem(pWsh);
		RangePtr pR1 = pS1->Cells;

		// Xác định cột chứa GD
		int iC0 = getColumn(pS1,"DMBB_STT");
		int iC1 = getColumn(pS1,"DMBB_MACV");
		int iC2 = getColumn(pS1,"DMBB_ND");

		// Xác định vị trí chứa comment END
		int xde = FindComment(pS1,iC0,"END");
		f_Delete_ListCV();

		// Thêm các giai đoạn vào combobo3
		cbb[3].ResetContent();
		ASSERT(cbb[3].GetCount() == 0);
		cbb[3].AddString(L"Tất cả giai đoạn");

		CString p0,c1,c2;
		int dem=0,cong=0;
		BOOLEAN f0 = FALSE;
		for (int i = 8; i < xde; i++)
		{
			p0 = pR1->GetItem(i,iC0);	// stt
			c1 = pR1->GetItem(i,iC1);	// mã
			if(_wtoi(p0)>0 && c1!=L"" && c1.Find(L"*")==-1)
			{
				f0 = FALSE;
				if(istyle==1)
				{
					// Kiểm tra lấy mẫu
					if(inLMCV[dem+1]==1) f0 = TRUE;
				}
				else if(istyle==2)
				{
					// Kiểm tra bê tông
					if(inKTBT[dem+1]==1) f0 = TRUE;
				}
				else f0 = TRUE;

				if(f0==TRUE)
				{
					c2 = pR1->GetItem(i,iC2);	// tên

					// Load lên list
					list1.InsertItem(cong,L"",0);
					list1.SetItemText(cong,1,p0);
					list1.SetItemText(cong,2,c1);
					list1.SetItemText(cong,3,c2);
					cong++;
				}
				dem++;
			}

			p0 = pR1->GetItem(i,iC1);
			if(p0.Left(2).MakeUpper()==L"GD")
			{
				// Thêm các giai đoạn vào combobo3
				p0 = pR1->GetItem(i,iC2);
				cbb[3].AddString(p0);
			}
		}
		
		txt[1].GetWindowTextW(msg);
		msg.Trim();
		if(msg==L"")
		{
			if(cong>0 && istyle==0) msg.Format(L"1-%d",cong);
			txt[1].SetWindowText(msg);
		}		

		msg.Format(L"Lựa chọn công việc (%d kết quả)",cong);
		kqgrp.SetWindowText(msg);

		cbb[3].SetCurSel(0);
	}
	catch(_com_error & error){_s(L"#QL4.12: Lỗi xác định giai đoạn công việc.");}
}


void Dg_In_ho_so::f_Load_GiaiDoan(int igdoan)
{
	try
	{
		_bstr_t pWsh = name_sheet("shHSNTCV");
		_WorksheetPtr pS1=xl->Sheets->GetItem(pWsh);
		RangePtr pR1 = pS1->Cells;

		// Xác định cột chứa GD
		int iC0 = getColumn(pS1,"DMBB_STT");
		int iC1 = getColumn(pS1,"DMBB_MACV");
		int iC2 = getColumn(pS1,"DMBB_ND");

		// Xác định vị trí chứa comment END
		int xde = FindComment(pS1,iC0,"END");

		f_Delete_ListCV();
		int n2 = cbb[2].GetCurSel();

		CString p0,c1,c2;
		int dem=0,gdn=0,cong=0,vtri=8,nbd=1,nkt=slcv_CV;
		BOOLEAN f0 = FALSE;
		if(igdoan>0)
		{
			// Xác định vtri = vị trí chứa giai đoạn cần hiển thị 
			for (int i = 8; i < xde; i++)
			{
				p0 = pR1->GetItem(i,iC0);	// stt
				c1 = pR1->GetItem(i,iC1);	// mã
				if(c1.MakeUpper()==L"GD") gdn++;
				if(_wtoi(p0)>0 && c1!=L"" && c1.Find(L"*")==-1) dem++;

				if(gdn==igdoan)
				{
					vtri=i;
					break;
				}
			}

			nbd=dem+1;
		}

		if(vtri<xde)
		{
			for (int i = vtri+1; i < xde; i++)
			{
				p0 = pR1->GetItem(i,iC0);	// stt
				c1 = pR1->GetItem(i,iC1);	// mã
				if(_wtoi(p0)>0 && c1!=L"")
				{
					f0 = FALSE;
					if(n2==1 || n2==6)
					{
						// Kiểm tra lấy mẫu
						if(inLMCV[dem+1]==1) f0 = TRUE;
					}
					else if(n2==5)
					{
						// Kiểm tra bê tông
						if(inKTBT[dem+1]==1) f0 = TRUE;
					}
					else f0 = TRUE;

					if(f0==TRUE)
					{
						c2 = pR1->GetItem(i,iC2);	// tên

						// Load lên list
						list1.InsertItem(cong,L"",0);
						list1.SetItemText(cong,1,p0);
						list1.SetItemText(cong,2,c1);
						list1.SetItemText(cong,3,c2);
						cong++;
					}
					dem++;
				}

				if(igdoan>0)
				{
					if(c1.MakeUpper()==L"GD")
					{
						if(dem>0) nkt=dem;
						break;
					}
				}
			}
		}

		msg = L"";
		if(cong>0 && n2!=1 && n2!=5 && n2!=6) msg.Format(L"%d-%d",nbd,nkt);
		txt[1].SetWindowText(msg);

		msg.Format(L"Lựa chọn công việc (%d kết quả)",cong);
		kqgrp.SetWindowText(msg);
	}
	catch(_com_error & error){_s(L"#QL4.12: Lỗi xác định giai đoạn công việc.");}
}

void Dg_In_ho_so::f_Load_ListVLGD(_bstr_t pSheet)
{
	try
	{
		_bstr_t pWsh = name_sheet(pSheet);
		_WorksheetPtr pS1=xl->Sheets->GetItem(pWsh);
		RangePtr pR1 = pS1->Cells;

		// Xác định cột
		int iC0,iC1,iC2,istart;

		if(pSheet==(_bstr_t)"shHSNTVL")
		{
			iC0 = getColumn(pS1,"DMVL_STT");
			iC1 = getColumn(pS1,"DMVL_MAVL");
			iC2 = getColumn(pS1,"DMVL_ND");
			istart = 8;
		}
		else if(pSheet==(_bstr_t)"shHSNTGD")
		{
			iC0 = getColumn(pS1,"DMGD_STT");
			iC1 = getColumn(pS1,"DMGD_HS");
			iC2 = getColumn(pS1,"DMGD_ND");
			istart = 8;
		}

		// Xác định vị trí chứa comment END
		int xde = FindComment(pS1,iC0,"END");

		f_Delete_ListCV();

		CString p0 = L"";
		CString c1,c2;
		int dem=0;

		for (int i = istart; i < xde; i++)
		{
			p0 = pR1->GetItem(i,iC0);
			if(_wtoi(p0)>0)
			{
				// Load lên list
				c1 = pR1->GetItem(i,iC1);	// mã
				c2 = pR1->GetItem(i,iC2);	// tên

				list1.InsertItem(dem,L"",0);
				list1.SetItemText(dem,1,p0);
				list1.SetItemText(dem,2,c1);
				list1.SetItemText(dem,3,c2);
				dem++;
			}
		}

		msg.Format(L"Lựa chọn công việc (%d kết quả)",dem);
		kqgrp.SetWindowText(msg);
	}
	catch(_com_error & error){_s(L"#QL4.13: Lỗi xác định giai đoạn vật liệu - giai đoạn.");}
}


void Dg_In_ho_so::f_Set_Duieu(int fset)
{
	if(fset==0) txt[1].SetWindowText(L"");
	txt[1].EnableWindow(fset);
	btn[3].EnableWindow(fset);
}


void Dg_In_ho_so::f_KTLaymau_VL_CV()
{
	try
	{
		// Xác định hiển thị hình ảnh
		_ihidepic = _wtoi(getVCell(psTS,L"TS_PIC"));

		// Định nghĩa lấy mẫu
		keyLMTN = getVCell(psTS,L"CF_LMTN");
		keyLMTN.Trim();
		if(keyLMTN==L"") keyLMTN=L"LM";

/////////////////////////////////////////////////////////////////////////////////////////////

		_bstr_t pWsh = name_sheet("shHSNTVL");
		_WorksheetPtr pS1=xl->Sheets->GetItem(pWsh);
		RangePtr pR1 = pS1->Cells;

		// Xác định cột
		int iC0 = getColumn(pS1,"DMVL_STT");
		int iC1 = getColumn(pS1,"DMVL_MAVL");
		int iLM = getColumn(pS1,"DMVL_MHDG");

		// Xác định vị trí chứa comment END
		int xde = FindComment(pS1,iC0,"END");

		CString p0=L"",p1=L"",l2=L"";
		int dem=0;
		inLMVL[0]=0;

		for (int i = 8; i < xde; i++)
		{
			p0 = pR1->GetItem(i,iC0);
			p1 = pR1->GetItem(i,iC1);			

			if(_wtoi(p0)>0 && p1!=L"") 
			{
				dem++;

				l2 = pR1->GetItem(i+1,iLM);
				if(l2.MakeUpper()==keyLMTN) inLMVL[dem]=1;
				else inLMVL[dem]=0;
			}
		}

/////////////////////////////////////////////////////////////////////////////////////////////

		pWsh = name_sheet("shHSNTCV");
		pS1=xl->Sheets->GetItem(pWsh);
		pR1 = pS1->Cells;

		// Xác định cột
		iC0 = getColumn(pS1,"DMBB_STT");
		iC1 = getColumn(pS1,"DMBB_MACV");
		iLM = getColumn(pS1,"DMBB_MH");

		// Xác định vị trí chứa comment END
		xde = FindComment(pS1,iC0,"END");

		dem=0;
		inLMCV[0]=0;
		inKTBT[0]=0;

		for (int i = 8; i < xde; i++)
		{
			p0 = pR1->GetItem(i,iC0);
			p1 = pR1->GetItem(i,iC1);			

			if(_wtoi(p0)>0 && p1!=L"" && p1.Find(L"*")==-1) 
			{
				dem++;

				l2 = pR1->GetItem(i+1,iLM);
				if(l2.MakeUpper()==keyLMTN) inLMCV[dem]=1;
				else inLMCV[dem]=0;

				p1.MakeUpper();
				p0 = p1.Left(4);
				if(p1.Left(3)==L"AG." || p0==L"AF.1" || p0==L"AF.2" 
					|| p0==L"AF.3" || p0==L"AF.4") inKTBT[dem]=1;
				else inKTBT[dem]=0;
			}
		}

/////////////////////////////////////////////////////////////////////////////////////////////

	}
	catch(...){}
}


void Dg_In_ho_so::f_Load_all_sheet()
{
	try
	{
		// Leo edit 13.03.2018
		_WorksheetPtr pSh1;
		int nSheet = xl->ActiveWorkbook->Sheets->Count;
		isxin[0]=0,isxin[1]=0,isxin[2]=0,isxbs=0;

		CString val=L"",sktr=L"";
		for (int i = 1; i <= nSheet; i++)
		{	
			pSh1 = pWb->Worksheets->GetItem(i);			
			val = (LPCTSTR)pSh1->CodeName;
			if(val.GetLength()<=8) continue;
			sktr = val.Left(8);			
			if(sktr==L"shHSNTVL")
			{
				isxin[0]++;
				SXI_VL_cn[isxin[0]] = val;

				if(val.Find(L"_")>0)
				{
					isxbs++;
					SXI_BS_cn[isxbs] = val;
					SXI_BS_name[isxbs] = (LPCTSTR)pSh1->Name;
				}
			}
			else if(sktr==L"shHSNTCV")
			{
				isxin[1]++;
				SXI_CV_cn[isxin[1]] = val;
				
				if(val.Find(L"_")>0)
				{
					isxbs++;
					SXI_BS_cn[isxbs] = val;
					SXI_BS_name[isxbs] = (LPCTSTR)pSh1->Name;
				}
			}
			else if(sktr==L"shHSNTGD")
			{
				isxin[2]++;
				SXI_GD_cn[isxin[2]] = val;
				
				if(val.Find(L"_")>0)
				{
					isxbs++;
					SXI_BS_cn[isxbs] = val;
					SXI_BS_name[isxbs] = (LPCTSTR)pSh1->Name;
				}
			}
		}
	}
	catch(...){}
}


void Dg_In_ho_so::f_Dot_thanh_toanVL()
{
	try
	{
		slTTVL=0;

		_bstr_t pWsh = name_sheet("shHSNTVL");
		_WorksheetPtr psDMVL = xl->Sheets->GetItem(pWsh);
		RangePtr prDMVL = psDMVL->Cells;

		// Xác định cột
		int iSTT = getColumn(psDMVL,"DMVL_STT");				
		int iXX = getColumn(psDMVL,"DMVL_XXU");
		int iDOT = getColumn(psDMVL,"DMVL_DOT");

		int iEnd = FindComment(psDMVL,iSTT,"END");

		int _itrung=0;
		CString szval=L"";
		for (int i = 8; i < iEnd; i++)
		{
			szval = GIOR(i,iSTT,prDMVL,L"");
			if(_wtoi(szval)>0)
			{
				szval = GIOR(i,iDOT,prDMVL,L"");
				szval.Trim();
				if(szval!=L"")
				{
					_itrung=0;

					if(slTTVL>0)
					{
						for (int j = 1; j <= slTTVL; j++)
						{
							if(szval==sDotTT_VL[j])
							{
								_itrung=1;
								break;
							}
						}
					}

					if(_itrung==0)
					{
						slTTVL++;
						sDotTT_VL[slTTVL] = szval;
					}
				}
			}
		}
	}
	catch(...){}
}


void Dg_In_ho_so::f_Dot_thanh_toanCV()
{
	try
	{
		slTTCV=0;

		_bstr_t pWsh = name_sheet("shHSNTCV");
		_WorksheetPtr psDMCV = xl->Sheets->GetItem(pWsh);
		RangePtr prDMCV = psDMCV->Cells;

		// Xác định cột
		int iSTT = getColumn(psDMCV,"DMBB_STT");				
		int iDTLM = getColumn(psDMCV,"DMBB_DTLMAU");
		int iDOT = getColumn(psDMCV,"DMBB_DOT");

		int iEnd = FindComment(psDMCV,iSTT,"END");

		int _itrung=0;
		CString szval=L"";
		for (int i = 8; i < iEnd; i++)
		{
			szval = GIOR(i,iSTT,prDMCV,L"");
			if(_wtoi(szval)>0)
			{
				szval = GIOR(i,iDOT,prDMCV,L"");
				szval.Trim();
				if(szval!=L"")
				{
					_itrung=0;

					if(slTTCV>0)
					{
						for (int j = 1; j <= slTTCV; j++)
						{
							if(szval==sDotTT_CV[j])
							{
								_itrung=1;
								break;
							}
						}
					}

					if(_itrung==0)
					{
						slTTCV++;
						sDotTT_CV[slTTCV] = szval;
					}
				}
			}
		}
	}
	catch(...){}
}


void Dg_In_ho_so::f_Dot_thanh_toanGD()
{
	try
	{
		slTTGD=0;

		_bstr_t pWsh = name_sheet("shHSNTGD");
		_WorksheetPtr psDMGD = xl->Sheets->GetItem(pWsh);
		RangePtr prDMGD = psDMGD->Cells;

		// Xác định cột
		int iSTT = getColumn(psDMGD,"DMGD_STT");				
		int iND = getColumn(psDMGD,"DMGD_ND");
		int iDOT = getColumn(psDMGD,"DMGD_DOT");

		int iEnd = FindComment(psDMGD,iSTT,"END");

		int _itrung=0;
		CString szval=L"";
		for (int i = 8; i < iEnd; i++)
		{
			szval = GIOR(i,iSTT,prDMGD,L"");
			if(_wtoi(szval)>0)
			{
				szval = GIOR(i,iDOT,prDMGD,L"");
				szval.Trim();
				if(szval!=L"")
				{
					_itrung=0;

					if(slTTGD>0)
					{
						for (int j = 1; j <= slTTGD; j++)
						{
							if(szval==sDotTT_GD[j])
							{
								_itrung=1;
								break;
							}
						}
					}

					if(_itrung==0)
					{
						slTTGD++;
						sDotTT_GD[slTTGD] = szval;
					}
				}
			}
		}
	}
	catch(...){}
}


void Dg_In_ho_so::f_SetCheck_VLCV(int ick1, int ick2, int ick3, int ick4, int ick5, int ick6)
{
	iLuuChk[1] = 0;
	iLuuChk[2] = 0;
	iLuuChk[4] = 0;
	iLuuChk[5] = 0;
	iLuuChk[6] = 0;
	iLuuChk[8] = 0;

	chk[1].EnableWindow(ick1);
	chk[2].EnableWindow(ick2);
	chk[5].EnableWindow(ick3);
	chk[6].EnableWindow(ick4);
	chk[7].EnableWindow(ick5);
	chk[8].EnableWindow(ick6);

	chk[1].SetCheck(FALSE);
	chk[2].SetCheck(FALSE);
	chk[5].SetCheck(FALSE);
	chk[6].SetCheck(FALSE);
	chk[7].SetCheck(FALSE);
	chk[8].SetCheck(FALSE);
}


void Dg_In_ho_so::f_Style_ListCV()
{
	list1.InsertColumn(0,L"",LVCFMT_CENTER,0);
	list1.InsertColumn(1,L"STT",LVCFMT_CENTER,_prt0);
	list1.InsertColumn(2,L"Mã hiệu",LVCFMT_LEFT,_prt1);
	list1.InsertColumn(3,L"Nội dung",LVCFMT_LEFT,_prt2);

	list1.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);
}


void Dg_In_ho_so::f_Delete_ListCV()
{
	list1.DeleteAllItems();
	ASSERT(list1.GetItemCount() == 0);
}


void Dg_In_ho_so::f_Add_congviec()
{
	msg = L"";
	int inum = 0;
	for (int i = 0; i<= list1.GetItemCount(); i++)
		if (list1.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
			{
				inum++;

				if(inum==1) msg.Format(L"%s",list1.GetItemText(i,1));
				else if(inum>1) msg.Format(L"%s,%s",msg,list1.GetItemText(i,1));
			}

	CString ktra = L"";
	txt[1].GetWindowTextW(ktra);
	ktra.Trim();
	if(ktra==L"") txt[1].SetWindowText(msg);
	else
	{
		ktra.Format(L"%s,%s",ktra,msg);
		txt[1].SetWindowText(ktra);
	}

	GotoDlgCtrl(GetDlgItem(IHS_T1));
}


CString Dg_In_ho_so::f_Gettext(int n)
{
	CString gt = L"";
	iLen = txt[n].LineLength(txt[n].LineIndex(0));
	lp = gt.GetBuffer(iLen);
	txt[n].GetLine(0, lp,iLen);
	gt.ReleaseBuffer();

	return gt;
}


void Dg_In_ho_so::OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	int nItem = pNMListView->iItem;

	CString inum = list1.GetItemText(nItem,1);

	if(_wtoi(inum)>0)
	{
		//msg = f_Gettext(1);
		txt[1].GetWindowTextW(msg);
		msg.Trim();
		if(msg==L"") txt[1].SetWindowText(inum);
		else
		{
			msg.Format(L"%s,%s",msg,inum);
			txt[1].SetWindowText(msg);
		}
	}
}


void Dg_In_ho_so::OnLvnKeydownL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);
	if(pLVKeyDow->wVKey == VK_SPACE) f_Add_congviec();

	*pResult = 0;
}


int Dg_In_ho_so::f_Xac_dinh_sl_In()
{
	CString ktra = L"";
	txt[1].GetWindowTextW(ktra);
	ktra.Replace(L" ",L"");

	if(ktra!=L"")
	{
		ktra+=L",";
		iKey=0;
		int vtri=0;
		CString val=L"";
		for (int i = 0; i < ktra.GetLength(); i++)
		{
			val = ktra.Mid(i,1);
			if((val==L"," || val==L";") && i>vtri)
			{
				iKey++;
				iCViec[iKey] = ktra.Mid(vtri,i-vtri);
				vtri=i+1;

				if(iKey==999) break;
			}
		}
	}
	else
	{
		// Chú ý những sheet ẩn sẽ không in
		// Trường hợp iKey = 0
		iKey = list1.GetItemCount();
		if(iKey>0) for (int i = 0; i < iKey; i++) iCViec[i+1] = list1.GetItemText(i,1);			
	}
	
	return iKey;
}


void Dg_In_ho_so::f_Kiem_tra_bsung(int istyle,CString inum,int iBanIn)
{
	// istyle = kiểu in/xem trước
	// inum = 1 công việc/ hoặc chuỗi công việc
	// iBanIn = số bản in
	
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
	//				if(erp==1) return;
	//				if(_myPrint.ibreaks==1 && _myPrint.inextpage==1) _itrang = _myPrint.ifirst;
	//				In_chitiet_bienban(istyle,csInbs[i],1,j,iBanIn);
	//			}
	//	}
	//	else
	//	{
	//		// Chỉ có 1 công việc được lựa chọn
	//		// Kiểm tra điều kiện ngắt trang khi thay đổi công việc
	//		for (int i = 1; i <= slbsThuc; i++)
	//		{
	//			if(erp==1) return;
	//			if(_myPrint.ibreaks==1 && _myPrint.inextpage==1) _itrang = _myPrint.ifirst;
	//			In_chitiet_bienban(istyle,csInbs[i],1,_wtoi(inum),iBanIn);
	//		}
	//	}
	//}
	//catch(_com_error & error){_s(L"#QL4.02- Lỗi in hồ sơ 12.");}
}


void Dg_In_ho_so::f_Kiem_tra_inum(CString inum,int icb1, int icb2)
{
	// inum = 1 công việc/ hoặc chuỗi công việc
	// icb1 - giá trị combobox1 - hồ sơ
	// icb2 - giá trị combobox1 - biên bản
	
	try
	{
		CString val=L"";
		int ibd=0,ikt=0;

		// Kiểm tra xem có 1 hay nhiều công việc?
		int iCount = cbb[2].GetCount();
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
			// Ứng với mỗi công việc sẽ in ra các biên bản
			// icb2>0 tương ứng in duy nhất 1 biên bản
			// Ngược lại in tất cả biên bản trong công tác được lựa chọn đó
			// Kiểm tra điều kiện ngắt trang khi thay đổi công việc
			if(erp==1) return;
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
	catch(_com_error & error){_s(L"#QL4.02- Lỗi in hồ sơ 2.");}
}



void Dg_In_ho_so::f_CodeNameActivate(_WorksheetPtr ps1)
{
	try
	{
		_bstr_t _pSht = ps1->CodeName;
		CString sEdit = (LPCTSTR)_pSht;

		if(sEdit.Left(12)==L"shHSNTVLLAYM")
		{
			// Lấy mẫu vật liệu
			_RSpin = getRow(ps1,"MAUVL_BB");
			_CSpin = getColumn(ps1,"MAUVL_BB");

			_bs2=L"PR_LMVL";
			_bs1 =L"PR_LMVL2";
		}
		else if(sEdit.Left(12)==L"shHSNTVLNBVL")
		{
			// Nội bộ vật liệu
			_RSpin = getRow(ps1,"NTNBVL_BB");
			_CSpin = getColumn(ps1,"NTNBVL_BB");

			_bs2=L"PR_NTNBVL";
			_bs1 =L"PR_NTNBVL2";
		}
		else if(sEdit.Left(11)==L"shHSNTVLYVL")
		{
			// Yêu cầu vật liệu
			_RSpin = getRow(ps1,"YCVL_BB");
			_CSpin = getColumn(ps1,"YCVL_BB");

			_bs2=L"PR_YCVL";
			_bs1 =L"PR_YCVL2";
		}
		else if(sEdit.Left(11)==L"shHSNTVLNVL")
		{
			// Nghiệm thu vật liệu
			_RSpin = getRow(ps1,"NTVL_BB");
			_CSpin = getColumn(ps1,"NTVL_BB");

			_bs2=L"PR_NTVL";
			_bs1 =L"PR_NTVL2";
		}
		else if(sEdit.Left(12)==L"shHSNTCVLMTN")
		{
			// Lấy mẫu công việc
			_RSpin = getRow(ps1,"NLM_BB");
			_CSpin = getColumn(ps1,"NLM_BB");

			_bs2=L"PR_LMTN";
			_bs1 =L"PR_LMTN2";
		}
		else if(sEdit.Left(12)==L"shHSNTCVNBCV" 
			&& sEdit.Left(13)!=L"shHSNTCVNBCV1" 
			&& sEdit.Left(13)!=L"shHSNTCVNBCV2")
		{
			// Lưu ý nguy hiểm: codename sheet nội bộ công việc này được đặt khá giống với 2 sheet ở dưới là
			// kiểm tra bê tông & theo dõi bê tông (codename = shHSNTCVNBCV)

			// Nội bộ công việc
			_RSpin = getRow(ps1,"NTNBCVGD_BB");
			_CSpin = getColumn(ps1,"NTNBCVGD_BB");

			_bs2=L"PR_NTNB";
			_bs1 =L"PR_NTNB2";
		}
		else if(sEdit.Left(13)==L"shHSNTCVNBCV1")
		{
			// Kiểm tra bê tông
			_RSpin = getRow(ps1,"NTNB_BB");
			_CSpin = getColumn(ps1,"NTNB_BB");

			_bs2=L"PR_BBKT";
			_bs1 =L"PR_BBKT2";
		}
		else if(sEdit.Left(13)==L"shHSNTCVNBCV2")
		{
			// Theo dõi bê tông
			_RSpin = getRow(ps1,"NTNBGD_BB");
			_CSpin = getColumn(ps1,"NTNBGD_BB");

			_bs2=L"PR_BBTD";
			_bs1 =L"PR_BBTD2";
		}
		else if(sEdit.Left(12)==L"shHSNTCVYNCV")
		{
			// Yêu cầu công việc
			_RSpin = getRow(ps1,"YCNTCV_BB");
			_CSpin = getColumn(ps1,"YCNTCV_BB");

			_bs2=L"PR_YCCV";
			_bs1 =L"PR_YCCV2";
		}
		else if(sEdit.Left(12)==L"shHSNTCVNTCV")
		{
			// Nghiệm thu công việc
			_RSpin = getRow(ps1,"NTCV_BB");
			_CSpin = getColumn(ps1,"NTCV_BB");

			_bs2=L"PR_NTCV";
			_bs1 =L"PR_NTCV2";
		}
		else if(sEdit.Left(14)==L"shHSNTGDNTBGD1")
		{
			// Nội bộ giai đoạn
			_RSpin = getRow(ps1,"NTNBGDG_BB");
			_CSpin = getColumn(ps1,"NTNBGDG_BB");
			
			_bs2=L"PR_NTNBGD";
			_bs1 =L"PR_NTNBGD2";
		}
		else if(sEdit.Left(13)==L"shHSNTGDYCNT1")
		{
			// Yêu cầu giai đoạn
			_RSpin = getRow(ps1,"YCGD_BB");
			_CSpin = getColumn(ps1,"YCGD_BB");

			_bs2=L"PR_YCGD";
			_bs1 =L"PR_YCGD2";
		}
		else if(sEdit.Left(13)==L"shHSNTGDNTGD1")
		{
			// Nghiệm thu giai đoạn
			_RSpin = getRow(ps1,"NTGD_BB");
			_CSpin = getColumn(ps1,"NTGD_BB");

			_bs2=L"PR_NTGD";
			_bs1 =L"PR_NTGD2";
		}
	}
	catch(_com_error & error){_s(L"#QL6.86: Lỗi xác định codename sheet.");}	
}


// Hàm in bổ sung các sheet thêm mới
void Dg_In_ho_so::In_toan_hs_bsung(int istyle,int iBanIn)
{
	/*for (int i = 1; i <= slbsThuc; i++)
	{
		if(erp==1) return;
		for (int j = 1; j <= ctInbs[i]; j++)
		{
			if(erp==1) return;
			if(_myPrint.ibreaks==1 && _myPrint.inextpage==1) _itrang = _myPrint.ifirst;
			In_chitiet_bienban(istyle,csInbs[i],1,i,iBanIn);
		}
	}*/
}


void Dg_In_ho_so::In_chitiet_bienban(_bstr_t sPrint,int ithem,int num_cviec)
{
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
		f_CodeNameActivate(psPrint);

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
		if(call_Worksheet_OnChange == NULL) {
				MessageBox(L"Not found Worksheet_OnChange in OpenCLGXD.dll ", L"Error", 0);
			}
			else {
				call_Worksheet_OnChange(0);
			}
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
		
		// độ dài khổ in thực ( top,bottom = 13mmm) = 297 - 2x13 =  271mm  xấp xỉ  775 point
		int _fsize = (int)((297-(_myPrint.itop+_myPrint.ibottom))/0.35+1);

		int _ipage[100];_ipage[1]=1;
		int dem = 0,_tong1=0,_tich=1;
		
		for (int i = 1; i < xde1; i++)
		{
			dem = psPrint->GetRange(prPrint->GetItem(i,1),prPrint->GetItem(i,1))->RowHeight;
			_tong1+=dem;

			if(_tong1>(_fsize*_tich))
			{
				_tich++;
				_ipage[_tich]=i;
			}
		}

		// _tong1 đơn vị là 'point'
		int page_in = (int)(_tong1*0.35/297)+5;

		// Kiểm tra điều kiện check: tự động căn chỉnh khi in (iLuuChk[3]=1)
		// iKieuIn=1 kiểu pview / = 0 kiểu print
		if(iKieuIn==1)
		{
			if(iLuuChk[3]==1)
			{
				if(_tich>1)
				for (int i = 1; i <_tich; i++)
				{
					if(erp==1) return;
					psPrint->GetRange(prPrint->GetItem(_ipage[i],ibdau),prPrint->GetItem(_ipage[i+1]-1,ikthuc))->PrintPreview();
				}

				int _tong2=0;
				for (int i = _ipage[_tich]; i < xde2; i++)
				{
					dem = psPrint->GetRange(prPrint->GetItem(i,1),prPrint->GetItem(i,1))->RowHeight;
					_tong2+=dem;
				}

				if(_tong2>_fsize)
				{
					if(erp==1) return;
					psPrint->GetRange(prPrint->GetItem(_ipage[_tich],ibdau),prPrint->GetItem(xde1-1,ikthuc))->PrintPreview();
					
					if(erp==1) return;
					psPrint->GetRange(prPrint->GetItem(xde1,ibdau),prPrint->GetItem(xde2,ikthuc))->PrintPreview();
				}
				else
				{
					if(erp==1) return;
					psPrint->GetRange(prPrint->GetItem(_ipage[_tich],ibdau),prPrint->GetItem(xde2,ikthuc))->PrintPreview();
				}
			}
			else
			{
				if(erp==1) return;
				psPrint->GetRange(prPrint->GetItem(1,ibdau),prPrint->GetItem(xde2,ikthuc))->PrintPreview();
			}
			
			int result = MessageBox(L"Bạn có muốn dừng chế độ xem trước?",L"Thông báo", MB_YESNO | MB_ICONQUESTION);
			if (result == 6) erp = 1;
		}
		else
		{
			int _islpage,_ingang,_idoc;

			// In hồ sơ
			if(iLuuChk[3]==1)
			{
				if(_tich>1) for (int i = 1; i <_tich; i++)
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

					psPrint->GetRange(
						prPrint->GetItem(_ipage[i],ibdau),
						prPrint->GetItem(_ipage[i+1]-1,ikthuc))->PrintOut(xlPrintSheetEnd,1,iSLBanIn);
				}

				int _tong2=0;
				for (int i = _ipage[_tich]; i < xde2; i++)
				{
					dem = psPrint->GetRange(prPrint->GetItem(i,1),prPrint->GetItem(i,1))->RowHeight;
					_tong2+=dem;
				}

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
					psPrint->GetRange(
						prPrint->GetItem(_ipage[_tich],ibdau),
						prPrint->GetItem(xde1-1,ikthuc))->PrintOut(xlPrintSheetEnd,1,iSLBanIn);
					
					// Đánh số trang
					if(_myPrint.inumpage==1)
					{
						psPrint->PageSetup->FirstPageNumber = _itrang;
						_itrang++;

						if(_myPrint.isecpage==0) psPrint->PageSetup->RightFooter = (_bstr_t)L"&P";
						else psPrint->PageSetup->CenterFooter = (_bstr_t)L"&P";
					}
					psPrint->GetRange(
						prPrint->GetItem(xde1,ibdau),
						prPrint->GetItem(xde2,ikthuc))->PrintOut(xlPrintSheetEnd,1,iSLBanIn);
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
					psPrint->GetRange(
						prPrint->GetItem(_ipage[_tich],ibdau),
						prPrint->GetItem(xde2,ikthuc))->PrintOut(xlPrintSheetEnd,1,iSLBanIn);
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

				psPrint->GetRange(
					prPrint->GetItem(1,ibdau),
					prPrint->GetItem(xde2,ikthuc))->PrintOut(xlPrintSheetEnd,page_in,iSLBanIn);
			}
		}

		szval.Format(L"In thành công biên bản số %d - %s",num_cviec,(LPCTSTR)shPrint);
		xl->PutStatusBar((_bstr_t)szval);
	}
	catch(_com_error & error){}
}


// Hàm in toàn bộ hồ sơ
void Dg_In_ho_so::In_toan_bo_ho_so()
{
	// istyle = kiểu in/xem trước
	// iBanIn = số bản in
	
	erp = 0;	// khi giá trị lỗi, erp=1, sẽ không in/xem trước được
	for(int i=1;i<=iSLBanIn;i++) // thức hiện vòng lặp in với số bản in = iBanIn 
	{
		// In hồ sơ vật liệu
		for(int j=1;j<=slcv_VL;j++)
		{
			if(_myPrint.ibreaks==1 && _myPrint.inextpage==1) _itrang = _myPrint.ifirst;
			for (int k = 1; k <= isxin[0]; k++)
			{
				if(erp==1) return;
				f_lua_chon_bienban(j,SXI_VL_cn[k],1,0);
			}
		}
	
		// In hồ sơ công việc
		for(int j=1;j<=slcv_CV;j++)
		{
			if(_myPrint.ibreaks==1 && _myPrint.inextpage==1) _itrang = _myPrint.ifirst;
			for (int k = 1; k <= isxin[1]; k++)
			{
				if(erp==1) return;
				f_lua_chon_bienban(j,SXI_CV_cn[k],2,0);
			}
		}

		// In hồ sơ giai đoạn
		for(int j=1;j<=slcv_GD;j++)
		{
			if(_myPrint.ibreaks==1 && _myPrint.inextpage==1) _itrang = _myPrint.ifirst;
			for (int k = 1; k <= isxin[2]; k++)
			{
				if(erp==1) return;
				f_lua_chon_bienban(j,SXI_GD_cn[k],3,0);
			}
		}
	}
}

// Hàm lựa chọn biên bản (hàm mới - 13.03.2018)
// Nếu giá trị icb1=-1 hay icb2=-1 mặc định không sử dụng đến 1 trong 2 biến đó
void Dg_In_ho_so::f_lua_chon_bienban(int num_cviec, CString sCodename, int icb1, int icb2)
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

		//////////// In chi tiết biên bản ////////////////
		In_chitiet_bienban((_bstr_t)sCodename,ithem,num_cviec);
	}
	catch(...){}
}


void Dg_In_ho_so::f_in_danh_muc_hs(int icbb)
{
	// icbb = lựa chọn danh mục cần in

	// Các danh mục
	if(icbb == 1) In_chitiet_danhmuc(L"Sheet6");
	else if(icbb == 2) In_chitiet_danhmuc(L"shHSNTVL");
	else if(icbb == 3) In_chitiet_danhmuc(L"shHSNTCV");
	else if(icbb == 4) In_chitiet_danhmuc(L"shHSNTGD");
}


void Dg_In_ho_so::In_chitiet_danhmuc(CString sPrint)
{
	// sPrint = sheet hiện hành
	
	try
	{
		shPrint = name_sheet((_bstr_t)sPrint);
		psPrint=xl->Sheets->GetItem(shPrint);
		prPrint = psPrint->Cells;

		int iCol1,iCol2;
		// Xác định cột đầu- cuối trang in
		if(sPrint == L"Sheet6")
		{
			// Danh mục hồ sơ
			iCol1 = getColumn(psPrint,"DMHS_STT");
			iCol2 = getColumn(psPrint,"DMHS_TTHS");
		}
		else if(sPrint == L"shHSNTVL")
		{
			// Danh mục vật liệu
			iCol1 = getColumn(psPrint,"DMVL_STT");
			iCol2 = 1+getColumn(psPrint,"DMVL_KBB");
		}
		else if(sPrint == L"shHSNTCV")
		{
			// Danh mục công việc
			iCol1 = getColumn(psPrint,"DMBB_STT");
			iCol2 = getColumn(psPrint,"DMBB_PS");
		}
		else if(sPrint == L"shHSNTGD")
		{
			// Danh mục giai đoạn
			iCol1 = getColumn(psPrint,"DMGD_STT");
			iCol2 = getColumn(psPrint,"DMGD_KBB");
		}

		int xde = FindComment(psPrint,iCol1,"END");

		// Kiểu in nằm nếu là biên bản danh mục
		psPrint->PageSetup->Orientation = xlLandscape;
		
		// Thiết lập chế độ in
		_psp = psPrint->PageSetup;
		f_Setup_Print(psPrint,0);

		if(iKieuIn==1)
		{
			psPrint->GetRange(prPrint->GetItem(1,iCol1),prPrint->GetItem(xde,iCol2))->PrintPreview();

			int result = MessageBox(L"Bạn có muốn kết thúc chế độ xem trước?",L"Thông báo", MB_YESNO | MB_ICONQUESTION);
			if (result == 6) erp = 1;
		}
		else psPrint->GetRange(prPrint->GetItem(1,iCol1),prPrint->GetItem(xde,iCol2))->PrintOut();
	}
	catch(_com_error & error){_s(L"#QL4.14: Lỗi in chi tiết danh mục.");}
}


void Dg_In_ho_so::f_xacdinh_sl_bban()
{
	slcv_VL = f_Num_SLBB((_bstr_t)"shHSNTVL",L"DMVL_STT",8);	// Số lượng biên bản vật liệu
	slcv_CV = f_Num_SLBB((_bstr_t)"shHSNTCV",L"DMBB_STT",8);	// Số lượng biên bản công việc
	slcv_GD = f_Num_SLBB((_bstr_t)"shHSNTGD",L"DMGD_STT",8);	// Số lượng biên bản giai đoạn
}

int Dg_In_ho_so::f_Num_SLBB(_bstr_t pSheet, CString NameColumn,int istart)
{
	// pSheet = sheet cần xác định vị trí chứa END
	// NameColumn = CodeName cột cần xác định END
	// istart = dòng bắt đầu duyệt xác định số lượng biên bản

	try
	{
		int islcv = 0;
		_bstr_t pWsh = name_sheet(pSheet);
		_WorksheetPtr pS1=xl->Sheets->GetItem(pWsh);
		RangePtr pR1 = pS1->Cells;

		// Xác định cột chứa END
		int iColumn = getColumn(pS1,NameColumn);

		// Xác định vị trí chứa comment END
		int xde = FindComment(pS1,iColumn,"END");

		CString p0 = L"";
		for (int i = istart; i < xde; i++)
		{
			p0 = pR1->GetItem(i,iColumn);
			if(_wtoi(p0)>0) islcv++;
		}

		return islcv;
	}
	catch(_com_error & error){_s(L"#QL4.12: Lỗi xác định số lượng công việc.");}
 
	return 0;
}

void Dg_In_ho_so::OnEnKillfocusT2()
{
	CString ktra = L"";
	txt[2].GetWindowTextW(ktra);
	ktra.Trim();
	if(_wtoi(ktra)<=0) txt[2].SetWindowText(L"1");
	else if(_wtoi(ktra)>100) txt[2].SetWindowText(L"100");
	else
	{
		msg.Format(L"%d",_wtoi(ktra));
		txt[2].SetWindowText(msg);
	}
}


void Dg_In_ho_so::OnBnClickedB4()
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


void Dg_In_ho_so::OnBnClickedCk8()
{
	CButton *m_ctlCheck = (CButton*) GetDlgItem(IHS_CK8);
	if(m_ctlCheck->GetCheck() == BST_UNCHECKED) iLuuChk[8]=0;
	else iLuuChk[8]=1;
}


void Dg_In_ho_so::OnBnClickedCk6()
{
	CButton *m_ctlCheck = (CButton*) GetDlgItem(IHS_CK6);
	if(m_ctlCheck->GetCheck() == BST_UNCHECKED) iLuuChk[5]=0;
	else iLuuChk[5]=1;
}


void Dg_In_ho_so::OnBnClickedCk7()
{
	CButton *m_ctlCheck = (CButton*) GetDlgItem(IHS_CK7);
	if(m_ctlCheck->GetCheck() == BST_UNCHECKED) iLuuChk[6]=0;
	else iLuuChk[6]=1;
}


void Dg_In_ho_so::OnBnClickedCk4()
{
	// Hiển thị hình ảnh
	CButton *m_ctlCheck = (CButton*) GetDlgItem(IHS_CK4);
	if(m_ctlCheck->GetCheck() == BST_UNCHECKED) _ihidepic=0;
	else  _ihidepic=1;
}


void Dg_In_ho_so::OnBnClickedB5()
{
	// Hiển thị hộp thoại mở rộng bảng tính
	AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
	Dlg_Cpys _dlg;
	_dlg.DoModal();
}


void Dg_In_ho_so::OnBnClickedB6()
{
	// Hàm check bản quyền
	if(f_Check_ban_quyen()!=1) return;

	iLuuKEY=0;
	luuhsCB2=0;
	luuhsCB1 = cbb[1].GetCurSel();
	if(luuhsCB1>0 && luuhsCB1<4)
	{
		luuhsCB2 = cbb[2].GetCurSel();
		CString szval=L"";
		txt[1].GetWindowTextW(szval);
		szval.Trim();
		if(szval==L"")
		{
			_s(L"Vui lòng lựa chọn công việc cần in");
			return;
		}

		iLuuKEY = f_Xac_dinh_sl_In();
	}

	iCountCB2 = cbb[2].GetCount();

	iLuuType=0;
	CDialog::OnOK();

	// Hiển thị hộp thoại lưu hồ sơ
	AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
	Dlg_luu_Hso _dlg;
	_dlg.DoModal();
}


void Dg_In_ho_so::OnBnClickedB7()
{
	nccBBin = cbb[1].GetCurSel();
	iSapxepIN=0;
	AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
	Dlg_sapxepIn _dlg;
	_dlg.DoModal();
}


void Dg_In_ho_so::OnTimer(UINT_PTR nIDEvent)
{
	iSec++;
	int d = iSec%60;
	if(d==0) sttHuongdan.SetWindowText(
		L"Nếu bạn không muốn in biên bản nào đó, bạn có thể ẩn biên bản đó đi");
	else if(d==10) sttHuongdan.SetWindowText(
		L"Để thay đổi thứ tự in của các biên bản, bạn có thể kéo thả biên bản lên trước hoặc sau");
	else if(d==20) sttHuongdan.SetWindowText(
		L"Lưu ý: Biên bản lấy mẫu công việc chỉ in khi có lấy mẫu (Ký hiệu LM)");
	else if(d==30) sttHuongdan.SetWindowText(
		L"Lưu ý: Biên bản kiểm tra bê tông chỉ in khi mã đó bắt đầu là AG. hoặc AF.1,2,3,4");
	else if(d==40) sttHuongdan.SetWindowText(
		L"Lưu ý: Biên bản theo dõi bê tông chỉ in khi biên bản đó có lấy mẫu");
	else if(d==50) sttHuongdan.SetWindowText(
		L"Khi lưu dạng XPS hoặc PDF, bạn có thể gộp nhiều file lẻ thành 1 file duy nhất bằng Smallpdf hoặc Adobe Acrobat");
	CDialog::OnTimer(nIDEvent);
}