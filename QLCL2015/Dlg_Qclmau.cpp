#include "stdafx.h"
#include "Dlg_Qclmau.h"
#include "Dlg_dsach_ctieu.h"
#include "LaymauNgang.h"

Dlg_Qclmau::Dlg_Qclmau(CWnd* pParent /*=NULL*/)
	: CDialog(Dlg_Qclmau::IDD, pParent)
{
	
}

void Dlg_Qclmau::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, QCLM_L1, listQclm);
	DDX_Control(pDX, QCLM_T1, edTxt);
	DDX_Control(pDX, QCLM_CB1, cbb1);
	DDX_Control(pDX, QCLM_T2, m_txt);
	DDX_Control(pDX, QCLM_CK1, m_chek);	
}

BEGIN_MESSAGE_MAP(Dlg_Qclmau, CDialog)	
	ON_WM_SIZE()
	ON_WM_SIZING()
	ON_WM_CONTEXTMENU()
	ON_WM_SYSCOMMAND()

	ON_COMMAND(ID_TI40002, &Dlg_Qclmau::OnTi40002)
	ON_COMMAND(ID_TI40003, &Dlg_Qclmau::OnTi40003)
	ON_COMMAND(ID_TI40001, &Dlg_Qclmau::OnTi40001)
	ON_BN_CLICKED(QCLM_B2, &Dlg_Qclmau::OnBnClickedB2)
	ON_BN_CLICKED(QCLM_B1, &Dlg_Qclmau::OnBnClickedB1)
	ON_EN_KILLFOCUS(QCLM_T1, &Dlg_Qclmau::OnEnKillfocusT1)
	ON_CBN_SELCHANGE(QCLM_CB1, &Dlg_Qclmau::OnCbnSelchangeCb1)
	ON_BN_CLICKED(QCLM_B3, &Dlg_Qclmau::OnBnClickedB3)
	ON_BN_CLICKED(QCLM_B4, &Dlg_Qclmau::OnBnClickedB4)
	ON_BN_CLICKED(QCLM_B5, &Dlg_Qclmau::OnBnClickedB5)
	ON_NOTIFY(LVN_KEYDOWN, QCLM_L1, &Dlg_Qclmau::OnLvnKeydownL1)
	ON_BN_CLICKED(QCLM_CK1, &Dlg_Qclmau::OnBnClickedCk1)
	ON_NOTIFY(NM_CLICK, QCLM_L1, &Dlg_Qclmau::OnNMClickL1)	
	ON_BN_CLICKED(QCLM_B6, &Dlg_Qclmau::OnBnClickedB6)
	ON_EN_SETFOCUS(QCLM_T2, &Dlg_Qclmau::OnEnSetfocusT2)
	ON_EN_SETFOCUS(QCLM_T1, &Dlg_Qclmau::OnEnSetfocusT1)
	ON_BN_CLICKED(QCLM_B8, &Dlg_Qclmau::OnBnClickedB8)
	ON_BN_CLICKED(QCLM_B7, &Dlg_Qclmau::OnBnClickedB7)
	ON_NOTIFY(NM_DBLCLK, QCLM_L1, &Dlg_Qclmau::OnNMDblclkL1)
END_MESSAGE_MAP()

BEGIN_EASYSIZE_MAP(Dlg_Qclmau)
	EASYSIZE(QCLM_S1,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(QCLM_CB1,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(QCLM_L1,ES_BORDER,ES_BORDER,ES_BORDER,ES_BORDER,0)
	EASYSIZE(QCLM_T2,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)
	EASYSIZE(QCLM_B4,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(QCLM_B5,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(QCLM_B6,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(QCLM_B3,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)
	EASYSIZE(QCLM_B1,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(QCLM_B2,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(QCLM_CK1,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)

	EASYSIZE(QCLM_B7,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(QCLM_B8,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
END_EASYSIZE_MAP

BOOL Dlg_Qclmau::OnInitDialog()
{
	CDialog::OnInitDialog();

	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	iChekUp=0;
	cbb1.SetCueBanner(L"Kích chọn mẫu quy cách lấy mẫu");
	m_txt.SetCueBanner(L"Nhập tên mẫu quy cách mới");
	m_txt.SubclassDlgItem(QCLM_T2,this);
	m_txt.SetBkColor(WHITE);

	// Load path
	if(getPathQCLM()==L"") return FALSE;

	// Load style list
	f_Style_dsach();

	// Load tooltip
	f_Load_tooltip();

	// Load CSDL
	// Kiểm tra dữ liệu đã có comment chưa?
	pSheet = pWb->ActiveSheet;
	sCodeName = (LPCTSTR)pSheet->CodeName;

	xl->PutScreenUpdating(true);
	pRange = pSheet->Cells;
	iRow = pSheet->Application->ActiveCell->Row;
	iCol = pSheet->Application->ActiveCell->Column;
	RangePtr PRS = pRange->GetItem(iRow,iCol);
	_bstr_t bscmt = (_bstr_t)"";
	if(PRS->GetComment()!=NULL) bscmt = PRS->GetComment()->Text();
	CString szval = GIOR(iRow,iCol,pRange,L"");
	szval.Trim();
	if(bscmt==(_bstr_t)L"QCLM" && szval!=L"") f_Load_dulieu(szval);
	else f_Load_dulieu(L"");

	// Set quy cách mặc định
	int num = _wtoi(getVCell(psTS,L"CF_QCLMCV"));
	if(num==1) m_chek.SetCheck(1);

	// Định nghĩa lấy mẫu
	keyLMTN = getVCell(psTS,L"CF_LMTN");
	keyLMTN.Trim();
	if(keyLMTN==L"") keyLMTN=L"LM";

	INIT_EASYSIZE;

	// Set size dialog mfc
	CString s = GIOR(4,40,prTS,L"");
	s.Trim();
	int rcs = _wtoi(s);
	if(rcs<=0) rcs = irecW;

	s = GIOR(4,41,prTS,L"");
	s.Trim();
	int ccs = _wtoi(s);
	if(ccs<=0) ccs = irecH;

	this->SetWindowPos(NULL,0,0,rcs,ccs,SWP_NOMOVE | SWP_NOZORDER);
	return TRUE;
}


BOOL Dlg_Qclmau::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if (pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_RETURN &&	pWndWithFocus == GetDlgItem(QCLM_T2))
	{
		OnBnClickedB3();
		return TRUE;
	}
	else if( pMsg->message == WM_LBUTTONDOWN ||
		pMsg->message == WM_LBUTTONUP ||
		pMsg->message == WM_MOUSEMOVE )
	{
		m_ToolTips.RelayEvent(pMsg);
		return CDialog::PreTranslateMessage(pMsg);
	}
	else if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		if(ClickEsc==0) OnBnClickedB2();
		else GotoDlgCtrl(GetDlgItem(QCLM_L1));
		ClickEsc=0;
		mnCphai=0;
		return TRUE; 
	}

	return FALSE;
}


void Dlg_Qclmau::OnSize(UINT nType, int cx, int cy) 
{
	CDialog::OnSize(nType, cx, cy);
	UPDATE_EASYSIZE;	
}


void Dlg_Qclmau::OnSizing(UINT fwSide, LPRECT pRect) 
{
	CDialog::OnSizing(fwSide, pRect);
	EASYSIZE_MINSIZE(100,100,fwSide,pRect);	// chiều rộng + chiều cao
}


void Dlg_Qclmau::f_Load_tooltip()
{
	CListCtrlEx * lst1 = (CListCtrlEx *)GetDlgItem(QCLM_L1); 
	CButton * btn7 = (CButton *)GetDlgItem(QCLM_B7); 
	CButton * btn8 = (CButton *)GetDlgItem(QCLM_B8); 

	m_ToolTips.Create(this);
	m_ToolTips.SetMaxTipWidth(600);
	m_ToolTips.SetDelayTime(900);

	m_ToolTips.AddTool(lst1, 
		L"Kích đúp để nhập và chỉnh sửa thông tin. "
		L"\nKích chuột phải để hiển thị menu trợ giúp."
		L"\n Nhập ký tự * trước tên chỉ tiêu để tạo thành tiêu đề."
		L"\n Nhập ký tự + trước tên chỉ tiêu để tạo tiểu mục."
		);

	m_ToolTips.AddTool(btn7, L"Sao chép ký tự ±");
	m_ToolTips.AddTool(btn8, L"Sao chép ký tự ÷");
}


void Dlg_Qclmau::f_Style_dsach()
{
	listQclm.InsertColumn(0,L"STT",LVCFMT_CENTER,jW1[0]);
	listQclm.InsertColumn(1,L"Tên chỉ tiêu",LVCFMT_LEFT,jW1[1]);
	listQclm.InsertColumn(2,L"Đơn vị",LVCFMT_CENTER,jW1[2]);
	listQclm.InsertColumn(3,L"Phương pháp thí nghiệm",LVCFMT_LEFT,jW1[3]);
	listQclm.InsertColumn(4,L"Kết quả",LVCFMT_CENTER,jW1[4]);

	listQclm.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);
	listQclm.SetColumnSorting(0, CListCtrlEx::Auto, CListCtrlEx::String);
	listQclm.SetDefaultEditor(NULL, NULL, &edTxt);
	edTxt.SubclassDlgItem(QCLM_T1,this);edTxt.SetBkColor(YELLOW153);edTxt.SetTextColor(BLUE);
}


void Dlg_Qclmau::f_Delete_list()
{
	listQclm.DeleteAllItems();
	ASSERT(listQclm.GetItemCount() == 0);
}


void Dlg_Qclmau::f_Load_dulieu(CString stype)
{
	TRY
	{
		CString szTMau=L"";
		if(stype!=L"")
		{
			int _pos = stype.Find(L":");
			if(_pos==-1) return;
			szTMau= stype.Left(_pos);
			szTMau.Trim();
		}
		
		if(connectDsn(pathQCLM)==-1) return;
		ObjConn.OpenAccessDB(pathQCLM, L"", L"");
		CRecordset* Recset;

		sIDmau=L"";
		if(szTMau!=L"")
		{
			SqlString.Format(L"SELECT * FROM tbTenMau WHERE TenMau = '%s';",szTMau);
			Recset = ObjConn.Execute(SqlString);
			while( !Recset->IsEOF() )
			{
				Recset->GetFieldValue(L"ID",sIDmau);
				break;
			}
			Recset->Close();
			sIDmau.Trim();
		}		

		if(sIDmau==L"")
		{
			// Xác định combobox mặc định
			SqlString = L"SELECT * FROM tbConfig WHERE ID = 'G01';";
			Recset = ObjConn.Execute(SqlString);
			while( !Recset->IsEOF() )
			{
				Recset->GetFieldValue(L"Gtri",sIDmau);
				break;
			}
			Recset->Close();
			sIDmau.Trim();
		}		

		// Xác định combobox dữ liệu
		CString szval=L"";
		int dem=0,vtri=0;
		SqlString.Format(L"SELECT * FROM tbTenMau WHERE ID LIKE 'T%s' ORDER BY TenMau ASC;",L"%");
		Recset = ObjConn.Execute(SqlString);
		while( !Recset->IsEOF() )
		{
			Recset->GetFieldValue(L"ID",szval);
			if(sIDmau==szval) vtri= dem;

			Recset->GetFieldValue(L"TenMau",szval);
			cbb1.AddString(szval);

			dem++;
			Recset->MoveNext();
		}
		Recset->Close();

		if(szval!=L"")
		{
			tongRw=0;
			cbb1.SetCurSel(vtri);
			SqlString.Format(L"SELECT * FROM tbNoidung WHERE (ID  = '%s') ORDER BY STT ASC;",sIDmau);
			Recset = ObjConn.Execute(SqlString);
			while( !Recset->IsEOF() )
			{
				Recset->GetFieldValue(L"Chitieu",svlTen[tongRw]);
				Recset->GetFieldValue(L"Donvi",svlDvi[tongRw]);
				Recset->GetFieldValue(L"PPTNghiem",svlPP[tongRw]);
				
				szval=L"";
				if(tongRw<9) szval.Format(L"0%d",tongRw+1);
				else szval.Format(L"%d",tongRw+1);

				listQclm.InsertItem(tongRw,szval,0);
				listQclm.SetItemText(tongRw,1,svlTen[tongRw]);
				listQclm.SetItemText(tongRw,2,svlDvi[tongRw]);
				listQclm.SetItemText(tongRw,3,svlPP[tongRw]);
				if(svlTen[tongRw].Left(1)==L"*") listQclm.SetRowColors(tongRw,M_VANG,BLACK);

				tongRw++;
				Recset->MoveNext();
			}
			Recset->Close();

			// Tách quy cách lấy mẫu
			if(tongRw>0) f_Tack_quycach();
		}

		ObjConn.CloseDatabase();
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("Lỗi xác định cơ sở dữ liệu.")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_Qclmau::f_Tack_quycach()
{
	try
	{
		iC4=1;
		if(sCodeName!=L"shHSNTVL") iC4 = getColumn(pSheet,L"DMBB_ND");
		else iC4 = getColumn(pSheet,L"DMVL_ND");

		CString czval = pRange->GetItem(iRow,iC4);
		czval.Trim();
		if(czval!=L"")
		{
			CString gan=L"",sGtr[500];
			int slg=0,vtri=0;
			for (int i = 0; i < czval.GetLength(); i++)
			{
				gan = czval.Mid(i,1);
				if((gan==L":" || gan==L";") && i>vtri)
				{
					sGtr[slg] = czval.Mid(vtri,i-vtri);
					sGtr[slg].Replace(L",",L".");
					sGtr[slg].Trim();
					vtri=i+1;
					slg++;
				}
			}

			// Bỏ qua sGtr[0] là tên quy cách lấy mẫu
			if(slg > tongRw) slg = tongRw;
			for (int i = 0; i < slg; i++) listQclm.SetItemText(i,4,sGtr[i+1]);
		}
	}
	catch(...){}
}


void Dlg_Qclmau::f_Get_size()
{
	CRect rectCtrl;
	GetWindowRect(&rectCtrl);	
	irecW = rectCtrl.Width();
	irecH = rectCtrl.Height();

	prTS->PutItem(4,40,irecW);
	prTS->PutItem(4,41,irecH);
}


void Dlg_Qclmau::f_Get_width()
{
	for (int i = 0; i < 5; i++)
		jW1[i] = listQclm.GetColumnWidth(i);
}


void Dlg_Qclmau::OnContextMenu(CWnd* /*pWnd*/, CPoint point)
{
	try
	{
		if(mnCphai==0) return;

		// Load the desired menu
		CMenu mnu1;
		mnu1.LoadMenu(IDR_MENU1);

		// Get a pointer to the button
		CListCtrlEx *pClist;
		pClist = reinterpret_cast<CListCtrlEx *>(GetDlgItem(QCLM_L1));

		// Find the rectangle around the button
		CRect rectSubmitButton;
		pClist->GetWindowRect(&rectSubmitButton);

		// Get a pointer to the first item of the menu
		CMenu *mnuP1 = mnu1.GetSubMenu(0);
		if(iRClick==-1 || iCClick==-1)
		{
			// Ẩn các tính năng nằm ngoài vùng kiểm soát
			mnuP1->EnableMenuItem(ID_TI40003, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
		}

		ASSERT(mnuP1);
	
		// Find out if the user right-clicked the button
		// because we are interested only in the button
		if( rectSubmitButton.PtInRect(point) ) // Since the user right-clicked the button, display the context menu
			mnuP1->TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON, point.x, point.y, this);
	}
	catch(...){}
}


void Dlg_Qclmau::OnTi40002()
{
	try
	{
		int dem = listQclm.GetItemCount();
		int vtr=-1;
		POSITION pos=listQclm.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			for (int i = 0; i < dem; i++)
				if (listQclm.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					vtr=i;
					break;
				}
		}

		if(vtr==-1) vtr = dem;
		
		iChekUp=1;
		listQclm.InsertItem(vtr,L"",0);
		listQclm.SetItemText(vtr,1,L"");
		listQclm.SetRowColors(vtr,WHITE,BLUE);

		f_Sort_stt_qc();
	}
	catch(...){}
}


void Dlg_Qclmau::OnTi40003()
{
	try
	{
		int dem = listQclm.GetItemCount();
		if(dem==0) return;

		int vtr=-1;
		POSITION pos=listQclm.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			for (int i = dem; i >= 0; i--)
				if (listQclm.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED) listQclm.DeleteItem(i);

			iChekUp=1;
			f_Sort_stt_qc();
		}
	}
	catch(...){}
}


void Dlg_Qclmau::OnTi40001()
{
	xl_tukhoa=L"";
	AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
	Dlg_dsach_ctieu open_dlg;
	open_dlg.DoModal();
}


void Dlg_Qclmau::OnBnClickedB2()
{
	mnCphai=0;
	ClickEsc=0;

	// Set quy cách mặc định
	int num = m_chek.GetCheck();
	int rcs = getRow(psTS,L"CF_QCLMCV");
	int ccs = getColumn(psTS,L"CF_QCLMCV");
	prTS->PutItem(rcs,ccs,num);
	prTS->PutItem(rcs,ccs+1,0);

	f_Get_width();
	f_Get_size();

	CDialog::OnCancel();
}


void Dlg_Qclmau::OnBnClickedB1()
{
	try
	{
		ClickEsc=0;

		// Set quy cách mặc định
		int num = m_chek.GetCheck();
		int rcs = getRow(psTS,L"CF_QCLMCV");
		int ccs = getColumn(psTS,L"CF_QCLMCV");
		prTS->PutItem(rcs,ccs,num);
		prTS->PutItem(rcs,ccs+1,0);

		// Hàm check bản quyền có điều kiện đếm
		if(f_Mod_check()!=1) return;

		if(iChekUp==1) f_Update_CSDL();

		int _pos = cbb1.GetCurSel();
		int count = listQclm.GetItemCount();
		if(_pos<0 || count==0) return;

		CString szval=L"";	
		cbb1.GetLBText(_pos,szval);
		szval.Trim();
		szval+=L":";

		CString sgtr=L"";
		for (int i = 0; i < count; i++)
		{
			sgtr = listQclm.GetItemText(i,4);
			sgtr.Replace(L",",L".");
			if(sgtr==L"") sgtr=L" ";
			szval.Format(L"%s%s;",szval,sgtr);
		}

////////////////////////////////////////////////////////

		iC2=-1,iC4=1;
		if(sCodeName!=L"shHSNTVL")
		{
			iC2 = getColumn(pSheet,L"DMBB_MH");
			iC4 = getColumn(pSheet,L"DMBB_ND");
		}
		else iC4 = getColumn(pSheet,L"DMVL_ND");
		
		if(iC2>0) pRange->PutItem(iRow,iC2,(_bstr_t)keyLMTN);
		pRange->PutItem(iRow,iC4,(_bstr_t)szval);

		RangePtr PRS = pRange->GetItem(iRow,iC4);
		PRS->Font->PutItalic(1);
		PRS->EntireRow->Rows->AutoFit();
		if(PRS->GetComment()!=NULL) PRS->ClearComments();
		PRS->AddComment((_bstr_t)L"QCLM");

		f_Get_width();
		f_Get_size();

		CDialog::OnOK();
	}
	catch(...){}
}


void Dlg_Qclmau::OnEnKillfocusT1()
{
	try
	{
		// kiểm tra dữ liệu nhập vào có trùng không?
		CString ktra=L"";
		edTxt.GetWindowTextW(ktra);
		ktra.Replace(L"'",L"");
		ktra.Replace(L"+-",L"±");
		ktra.Replace(L"//",L"÷");

		CString val = listQclm.GetItemText(CLRow,CLCol);
		if(ktra==val) return;
		edTxt.SetWindowText(ktra);

		iChekUp=1;

		// Kiểm tra cột số 1
		if(CLRow==-1 || CLCol!=1) return;

		// Kiểm tra ký tự + hoặc *
		if(ktra.Left(1)!=L"+" && ktra.Left(1)!=L"*")
		{
			listQclm.SetRowColors(CLRow,WHITE,BLACK);
			return;
		}

		// Tô màu
		if(ktra.Left(1)==L"*")
		{
			listQclm.SetRowColors(CLRow,YELLOW,BLACK);
			return;
		}

		listQclm.SetRowColors(CLRow,WHITE,BLACK);
		int N = ktra.GetLength();
		if(N>=2 && ktra.Left(2)!=L"+ ") val = L"+ " + ktra.Right(N-1);
		ktra = L"   " + val;
		edTxt.SetWindowText(ktra);
		listQclm.SetItemText(CLRow,CLCol,ktra);	
		
	}
	catch(...){}
}


void Dlg_Qclmau::OnCbnSelchangeCb1()
{
	TRY
	{
		mnCphai=1;
		ClickEsc=0;
		int _pos = cbb1.GetCurSel();
		if(_pos<0) return;

		iChekUp=0;
		f_Delete_list();

		if(connectDsn(pathQCLM)==-1) return;
		ObjConn.OpenAccessDB(pathQCLM, L"", L"");

		// Xác định combobox mặc định
		CString szval=L"";
		cbb1.GetLBText(_pos,szval);

		// Xác định ID tương ứng
		sIDmau=L"T00";
		SqlString.Format(L"SELECT * FROM tbTenMau WHERE TenMau LIKE '%s';",szval);
		CRecordset* Recset = ObjConn.Execute(SqlString);
		while( !Recset->IsEOF() )
		{
			Recset->GetFieldValue(L"ID",sIDmau);
			break;
		}
		Recset->Close();

		// Xác định combobox dữ liệu
		tongRw=0;
		SqlString.Format(L"SELECT * FROM tbNoidung WHERE (ID  = '%s') ORDER BY STT ASC;",sIDmau);
		Recset = ObjConn.Execute(SqlString);
		while( !Recset->IsEOF() )
		{
			Recset->GetFieldValue(L"Chitieu",svlTen[tongRw]);
			Recset->GetFieldValue(L"Donvi",svlDvi[tongRw]);
			Recset->GetFieldValue(L"PPTNghiem",svlPP[tongRw]);
			
			szval=L"";
			if(tongRw<9) szval.Format(L"0%d",tongRw+1);
			else szval.Format(L"%d",tongRw+1);

			listQclm.InsertItem(tongRw,szval,0);
			listQclm.SetItemText(tongRw,1,svlTen[tongRw]);
			listQclm.SetItemText(tongRw,2,svlDvi[tongRw]);
			listQclm.SetItemText(tongRw,3,svlPP[tongRw]);
			if(svlTen[tongRw].Left(1)==L"*") listQclm.SetRowColors(tongRw,YELLOW,BLACK);

			tongRw++;
			Recset->MoveNext();
		}
		Recset->Close();

		// Update mẫu
		SqlString.Format(L"UPDATE tbConfig SET Gtri='%s' WHERE ID='G01';",sIDmau);
		ObjConn.ExecuteRB(SqlString);

		ObjConn.CloseDatabase();
		GotoDlgCtrl(GetDlgItem(QCLM_L1));
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("Lỗi xác định cơ sở dữ liệu.")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_Qclmau::OnBnClickedB3()
{
	TRY
	{
		mnCphai=1;
		ClickEsc=0;

		CString szval=L"";
		m_txt.GetWindowTextW(szval);
		szval.Trim();
		szval.Replace(L"'",L"");
		if(szval==L"")
		{
			GotoDlgCtrl(GetDlgItem(QCLM_T2));
			return;
		}

		if(connectDsn(pathQCLM)==-1) return;
		ObjConn.OpenAccessDB(pathQCLM, L"", L"");

		// Kiểm tra tên mẫu đã tồn tại trong CSDL chưa?
		CString sdem=L"";
		SqlString.Format(L"SELECT COUNT(*) AS iDem FROM tbTenMau WHERE (TenMau = '%s');",szval);
		CRecordset* Recset = ObjConn.Execute(SqlString);
		Recset->GetFieldValue(L"iDem",sdem);
		Recset->Close();

		if(_wtoi(sdem)>0)
		{
			ObjConn.CloseDatabase();
			_s(L"Dữ liệu đã tồn tại. Vui lòng nhập tên khác!");
			GotoDlgCtrl(GetDlgItem(QCLM_T2));
			m_txt.SetTextColor(RED);
			return;
		}

		// Kiểm tra mã đã tồn tại trong CSDL chưa?
		CString sLuu=L"";
		for (int i = 1; i <= 99; i++)
		{
			if(i<=9) sLuu.Format(L"T0%d",i);
			else sLuu.Format(L"T%d",i);

			SqlString.Format(L"SELECT COUNT(*) AS iDem FROM tbTenMau WHERE (ID = '%s');",sLuu);
			Recset = ObjConn.Execute(SqlString);
			Recset->GetFieldValue(L"iDem",sdem);
			Recset->Close();

			if(_wtoi(sdem)==0) break;
		}

		// Lưu nội dung vào CSDL
		sIDmau = sLuu;
		SqlString.Format(L"INSERT INTO tbTenMau (ID,TenMau) VALUES ('%s','%s');",sLuu,szval);
		ObjConn.ExecuteRB(SqlString);
		m_txt.SetTextColor(BLACK);

		// Thêm vào combobox
		cbb1.AddString(szval);
		cbb1.SetCurSel(cbb1.GetCount()-1);

		// Xóa list
		f_Delete_list();
		tongRw=0;
		iChekUp=0;

		ObjConn.CloseDatabase();
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("Lỗi xác định cơ sở dữ liệu.")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_Qclmau::OnBnClickedB4()
{
	mnCphai=1;
	ClickEsc=0;
	if(f_Update_CSDL()==1) _s(L"Cập nhật dữ liệu thành công!");
}


int Dlg_Qclmau::f_Update_CSDL()
{
	TRY
	{
		iChekUp=0;
		if(f_Check_update()==0) return 0;

		if(connectDsn(pathQCLM)==-1) return 0;
		ObjConn.OpenAccessDB(pathQCLM, L"", L"");

		// Xóa dữ liệu
		SqlString.Format(L"DELETE FROM tbNoidung WHERE (ID='%s');",sIDmau);
		ObjConn.ExecuteRB(SqlString);

		// Thêm mới
		int count = listQclm.GetItemCount();
		if(count>0)
		{
			int dem=1;
			CRecordset* Recset;
			CString sdem=L"";
			CString vl[5]={L"",L"",L"",L"",L""};
			for (int i = 0; i < count; i++)
			{
				if(i<9) vl[0].Format(L"0%d",i+1);
				else vl[0].Format(L"%d",i+1);

				for (int j = 1; j < 5; j++) vl[j] = listQclm.GetItemText(i,j);

				SqlString.Format(
					L"INSERT INTO tbNoidung (ID,STT,Chitieu,Donvi,PPTNghiem) VALUES ('%s','%s','%s','%s','%s');",
					sIDmau,vl[0],vl[1],vl[2],vl[3]);
				ObjConn.ExecuteRB(SqlString);

				// Lưu thêm tên chỉ tiêu & đơn vị vào bảng tbDanhsach
				// Kiểm tra dữ liệu có tồn tại không?
				SqlString.Format(L"SELECT COUNT(*) AS iDem FROM tbDanhsach WHERE (DSTen LIKE '%s' AND Donvi LIKE '%s');",vl[1],vl[2]);
				Recset = ObjConn.Execute(SqlString);
				Recset->GetFieldValue(L"iDem",sdem);
				Recset->Close();

				if(_wtoi(sdem)==0)
				{
					// Xác định ID mới (Dxxx)
					for (int j = dem; j <= 999; j++)
					{
						if(j<9) vl[0].Format(L"D00%d",j+1);
						else if(j>=9 && j<99) vl[0].Format(L"D0%d",j+1);
						else vl[0].Format(L"D%d",j+1);

						SqlString.Format(L"SELECT COUNT(*) AS iDem FROM tbDanhsach WHERE (ID = '%s');",vl[0]);
						Recset = ObjConn.Execute(SqlString);
						Recset->GetFieldValue(L"iDem",sdem);
						Recset->Close();

						if(_wtoi(sdem)==0)
						{
							dem=j+1;
							break;
						}
					}

					SqlString.Format(
						L"INSERT INTO tbDanhsach (ID,DSTen,Donvi) VALUES ('%s','%s','%s');",vl[0],vl[1],vl[2]);
					ObjConn.ExecuteRB(SqlString);
				}
			}
		}	

		ObjConn.CloseDatabase();
		return 1;
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("Lỗi xác định cơ sở dữ liệu.")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_Qclmau::OnBnClickedB5()
{
	TRY
	{
		mnCphai=1;
		ClickEsc=0;
		int sl = cbb1.GetCount();
		if(sl==0) return;

		int vtr = cbb1.GetCurSel();
		if(vtr<0) return;

		CString vl1=L"",szval=L"";
		cbb1.GetLBText(vtr,vl1);

		szval.Format(L"Bạn có chắc chắn xóa mẫu '%s'?",vl1);
		if(_yn(szval,0)!=6) return;

		iChekUp=0;

		// Dữ liệu hiển thị trên list
		f_Delete_list();

		int dem=0;
		CString sluu[100];
		for (int i = 0; i < sl; i++)
		{
			if(i!=vtr)
			{
				cbb1.GetLBText(i,sluu[dem]);
				dem++;
			}
		}

		cbb1.ResetContent();
		ASSERT(cbb1.GetCount() == 0);
		if(dem>0) for (int i = 0; i < dem; i++) cbb1.AddString(sluu[i]);

		if(connectDsn(pathQCLM)==-1) return;
		ObjConn.OpenAccessDB(pathQCLM, L"", L"");

		// Xóa dữ liệu
		SqlString.Format(L"DELETE FROM tbNoidung WHERE (ID='%s');",sIDmau);
		ObjConn.ExecuteRB(SqlString);

		SqlString.Format(L"DELETE FROM tbTenMau WHERE (ID='%s');",sIDmau);
		ObjConn.ExecuteRB(SqlString);
		ObjConn.CloseDatabase();
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("Lỗi xác định cơ sở dữ liệu.")+e->m_strError);
	}
	END_CATCH;
}


int Dlg_Qclmau::f_Check_update()
{
	try
	{
		int count = listQclm.GetItemCount();
		if(count==0 && tongRw==0) return 0;
		if(count!=tongRw) return 1;		

		CString szval=L"";
		for (int i = 0; i < count; i++)
		{
			szval = listQclm.GetItemText(i,1);
			if(szval!=svlTen[i]) return 1;

			szval = listQclm.GetItemText(i,2);
			if(szval!=svlDvi[i]) return 1;

			szval = listQclm.GetItemText(i,3);
			if(szval!=svlPP[i]) return 1;
		}

		return 0;
	}
	catch(...){}
}


void Dlg_Qclmau::OnLvnKeydownL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);
	if(pLVKeyDow->wVKey == VK_DELETE) OnTi40003();
	if(pLVKeyDow->wVKey == VK_F4) OnTi40001();

	*pResult = 0;
}


void Dlg_Qclmau::OnBnClickedCk1()
{
	mnCphai=1;
	ClickEsc=0;
}


void Dlg_Qclmau::OnNMClickL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	mnCphai=1;
	ClickEsc=0;
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	iRClick = pNMListView->iItem;
	iCClick = pNMListView->iSubItem;
	
	*pResult = 0;
}


void Dlg_Qclmau::OnSysCommand(UINT nID, LPARAM lParam)
{
	if(nID == SC_CLOSE) OnBnClickedB2();
	else CDialog::OnSysCommand(nID, lParam);
}

void Dlg_Qclmau::OnBnClickedB6()
{
	mnCphai=1;
	ClickEsc=0;
	CDialog::OnOK();
	AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
	LaymauNgang _dlg;
	_dlg.DoModal();	
}


void Dlg_Qclmau::OnEnSetfocusT2()
{
	mnCphai=1;
	ClickEsc=0;
}


void Dlg_Qclmau::OnEnSetfocusT1()
{
	ClickEsc=1;
}


void Dlg_Qclmau::OnBnClickedB7()
{
	CopyDulieu(L"±");
}


void Dlg_Qclmau::OnBnClickedB8()
{
	CopyDulieu(L"÷");
}


void Dlg_Qclmau::OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	mnCphai=1;
	ClickEsc=0;
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	iRClick = pNMListView->iItem;
	iCClick = pNMListView->iSubItem;

	if(iRClick==-1 || iCClick==-1)
	{
		int vtr = listQclm.GetItemCount();
		CString szval=L"";
		if(vtr<9) szval.Format(L"0%d",vtr+1);
		else szval.Format(L"%d",vtr+1);

		iChekUp=1;
		listQclm.InsertItem(vtr,szval,0);
		listQclm.SetItemText(vtr,1,L"[giá trị]");
		listQclm.SetRowColors(vtr,WHITE,BLUE);
	}
	
	*pResult = 0;
}
