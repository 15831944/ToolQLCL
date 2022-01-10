#include "stdafx.h"
#include "Dlg_Manager_wb.h"

Dlg_Manager_wb::Dlg_Manager_wb(CWnd* pParent /*=NULL*/)
	: CDialog(Dlg_Manager_wb::IDD, pParent)
{
	
}

void Dlg_Manager_wb::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, QLWB_L1, qlwb_list);
	DDX_Control(pDX, QLWB_T1, qlwb_edit);
	DDX_Control(pDX, QLWB_K1, m_check);
	DDX_Control(pDX, QLWB_B1, m_btn[1]);
	DDX_Control(pDX, QLWB_B2, m_btn[2]);
	DDX_Control(pDX, QLWB_B3, m_btn[3]);
	DDX_Control(pDX, QLWB_B4, m_btn[4]);
	DDX_Control(pDX, QLWB_B5, m_btn[5]);
	DDX_Control(pDX, QLWB_B6, m_btn[6]);
	DDX_Control(pDX, QLWB_B7, m_btn[7]);
}

BEGIN_MESSAGE_MAP(Dlg_Manager_wb, CDialog)
	ON_WM_SIZE()
	ON_WM_SIZING()

	ON_BN_CLICKED(QLWB_B1, &Dlg_Manager_wb::OnBnClickedB1)
	ON_BN_CLICKED(QLWB_B2, &Dlg_Manager_wb::OnBnClickedB2)
	ON_BN_CLICKED(QLWB_K1, &Dlg_Manager_wb::OnBnClickedK1)
	ON_BN_CLICKED(QLWB_B3, &Dlg_Manager_wb::OnBnClickedB3)
	ON_NOTIFY(NM_CLICK, QLWB_L1, &Dlg_Manager_wb::OnNMClickL1)
	ON_EN_KILLFOCUS(QLWB_T1, &Dlg_Manager_wb::OnEnKillfocusT1)
	ON_BN_CLICKED(QLWB_B4, &Dlg_Manager_wb::OnBnClickedB4)
	ON_BN_CLICKED(QLWB_B5, &Dlg_Manager_wb::OnBnClickedB5)
	ON_BN_CLICKED(QLWB_B6, &Dlg_Manager_wb::OnBnClickedB6)
	ON_BN_CLICKED(QLWB_B7, &Dlg_Manager_wb::OnBnClickedB7)
	ON_NOTIFY(LVN_KEYDOWN, QLWB_L1, &Dlg_Manager_wb::OnLvnKeydownL1)
	ON_NOTIFY(NM_DBLCLK, QLWB_L1, &Dlg_Manager_wb::OnNMDblclkL1)
	ON_NOTIFY(NM_RCLICK, QLWB_L1, &Dlg_Manager_wb::OnNMRClickL1)
END_MESSAGE_MAP()

BEGIN_EASYSIZE_MAP(Dlg_Manager_wb)
	EASYSIZE(QLWB_G1,ES_BORDER,ES_BORDER,ES_BORDER,ES_BORDER,0)
	EASYSIZE(QLWB_L1,ES_BORDER,ES_BORDER,ES_BORDER,ES_BORDER,0)
	EASYSIZE(QLWB_B1,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(QLWB_B2,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(QLWB_B3,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(QLWB_B4,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(QLWB_B5,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(QLWB_B6,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(QLWB_B7,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
END_EASYSIZE_MAP

BOOL Dlg_Manager_wb::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	iCode=0;
	f_Style_List();
	f_Load_ToolTip();

	f_GetAllSheet();
	f_Load_AllSheet();

	INIT_EASYSIZE;

	return TRUE; 
}

// Bắt sự kiện click Enter
BOOL Dlg_Manager_wb::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if (pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(QLWB_L1))
	{
		f_FindSheet();
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
		catch(...){}*/

		CDialog::OnCancel();
		return TRUE; 
	}
	else if( pMsg->message == WM_LBUTTONDOWN ||
		pMsg->message == WM_LBUTTONUP ||
		pMsg->message == WM_MOUSEMOVE )
	{
		m_toolTips.RelayEvent(pMsg);
		return CDialog::PreTranslateMessage(pMsg);
	}

	return FALSE;
}

void Dlg_Manager_wb::OnSize(UINT nType, int cx, int cy) 
{
	CDialog::OnSize(nType, cx, cy);
	UPDATE_EASYSIZE;
}

void Dlg_Manager_wb::OnSizing(UINT fwSide, LPRECT pRect) 
{
	CDialog::OnSizing(fwSide, pRect);
	EASYSIZE_MINSIZE(50,50,fwSide,pRect);	// chiều rộng + chiều cao
}


void Dlg_Manager_wb::f_Load_ToolTip()
{
	try
	{
		CListCtrlEx *ls1 = (CListCtrlEx *)GetDlgItem(QLWB_L1);
		CButton *btn3 = (CButton *)GetDlgItem(QLWB_B3);
		CButton *btn4 = (CButton *)GetDlgItem(QLWB_B4);
		CButton *btn5 = (CButton *)GetDlgItem(QLWB_B5);
		CButton *btn6 = (CButton *)GetDlgItem(QLWB_B6);

		m_toolTips.Create(this);
		m_toolTips.SetMaxTipWidth(600);
		m_toolTips.SetDelayTime(900);

		m_toolTips.AddTool(ls1, 
			L"- Tích chọn bảng tính cần hiển thị."
			L"\n- Kích đúp để thay đổi tên bảng tính hoặc Codename."
			L"\n- Sử dụng phím (lên+xuống) để sắp xếp lại thứ bảng tính."
			L"\n- Chọn một hoặc nhiều bảng tính để đổi màu.");

		m_toolTips.AddTool(btn3, L"Kích đúp hoặc kích phải vào cột màu để thay đổi.");
		m_toolTips.AddTool(btn4, L"Kích chọn vị trí cần di chuyển lên (phím tắt UP).");
		m_toolTips.AddTool(btn5, L"Kích chọn vị trí cần di chuyển xuống (phím tắt DOWN).");
		m_toolTips.AddTool(btn6, L"Kích chọn hai vị trí cần hoán đổi (phím tắt TAB).");
	}
	catch(...){}
}


void Dlg_Manager_wb::f_Style_List()
{
	// Định dạng kiểu List Control
	qlwb_list.InsertColumn(0,L"",LVCFMT_CENTER,22);							// Hiện/Ẩn
	qlwb_list.InsertColumn(1,L"Trạng thái",LVCFMT_LEFT,70);					// Trạng thái hiện/ẩn
	qlwb_list.InsertColumn(2,L"STT",LVCFMT_CENTER,50);						// Thứ tự bảng tính
	qlwb_list.InsertColumn(3,L"Tên bảng tính",LVCFMT_LEFT,300);				// Tên bảng tính
	qlwb_list.InsertColumn(4,L"Codename",LVCFMT_LEFT,0);					// Code name
	qlwb_list.InsertColumn(5,L"Màu sắc",LVCFMT_LEFT,100);					// Màu hiển thị
	qlwb_list.InsertColumn(6,L"",LVCFMT_CENTER,0);							// STT gốc

	qlwb_list.SetColumnReadOnly(0);
	qlwb_list.SetColumnReadOnly(1);
	qlwb_list.SetColumnReadOnly(2);
	qlwb_list.SetColumnReadOnly(5);
	qlwb_list.SetColumnReadOnly(6);
	qlwb_list.SetColumnEditor(3, &qlwb_edit);
	//qlwb_list.SetColumnEditor(4, &qlwb_edit);	(Tạm thời khóa tính năng thay đổi codename)

	qlwb_list.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_CHECKBOXES | LVS_EX_GRIDLINES);
}


void Dlg_Manager_wb::OnBnClickedB1()
{
	try
	{
		xl->PutScreenUpdating(VARIANT_FALSE);

///////// Đổi tên Name & Codename
		CString val=L"";
		_WorksheetPtr pSh1;
		int n=0,kt=0;
		int iUnhide=0;
		int iChange=0;

		// Các trạng thái cần kiểm tra
		for (int i = 0; i < nSheet; i++)
		{
			// Kiểm tra các vị trí có thay đổi hay không?
			n = _wtoi(qlwb_list.GetItemText(i,6));
			if(n!=i+1) iChange=1;

			// Kiểm tra có phải tất cả các sheet cùng ẩn không?
			if(qlwb_list.GetCheck(i)==1) iUnhide=1;
		}

		// Cần hiển thị ít nhất 1 sheet
		if(iUnhide==0) qlwb_list.SetCheck(0,1);

////////// Thay đổi name + codename + color + hide
		for (int i = 0; i < nSheet; i++)
		{
			n = _wtoi(qlwb_list.GetItemText(i,6));
			pSh1 = pWb->Worksheets->GetItem(n);

			// Name
			val = qlwb_list.GetItemText(i,3);
			if(val!=(LPCTSTR)MyWb[n].sName) pSh1->PutName((_bstr_t)val);
			
			// Codename
			val = qlwb_list.GetItemText(i,4);
			if(val!=(LPCTSTR)MyWb[n].sCode && val!=L"") 
				pWb->VBProject->VBComponents->Item(n)->PutName((_bstr_t)val); 
			
			// Color
			pSh1->Tab->PutColor(MyWb[n].cAfter);
			
			// Hide/Unhide
			if(qlwb_list.GetCheck(i)==0) pSh1->PutVisible(XlSheetVisibility::xlSheetHidden);
			else pSh1->PutVisible(XlSheetVisibility::xlSheetVisible);
		}

/////// Đổi các vị trí nếu có
		if(iChange==1)
		{
			_variant_t varOption((long) DISP_E_PARAMNOTFOUND,VT_ERROR);
			VARIANT  varBefore;
			varBefore.vt = VT_DISPATCH;
			for (int i = 0; i < nSheet-1; i++)
			{
				n = _wtoi(qlwb_list.GetItemText(i,6));
				varBefore.pdispVal = pWb->Worksheets->GetItem(i+1);
				pSh1 = xl->Worksheets->GetItem(MyWb[n].sName);

				// Chèn vào trước bảng tính
				pSh1->Move(varBefore,varOption);
			}

			// Chèn vào sau bảng tính
			n = _wtoi(qlwb_list.GetItemText(nSheet-1,6));
			varBefore.pdispVal = pWb->Worksheets->GetItem(nSheet);
			pSh1 = xl->Worksheets->GetItem(MyWb[n].sName);
			pSh1->Move(varOption,varBefore);
		}
		CoUninitialize();

		OnOK();
	}
	catch(...){}
}

void Dlg_Manager_wb::OnBnClickedB2()
{
	m_check.SetCheck(0);
	f_Delete_List();
	f_Load_AllSheet();
}


void Dlg_Manager_wb::f_Delete_List()
{
	if(qlwb_list.GetItemCount()>0)
	{
		qlwb_list.DeleteAllItems();
		ASSERT(qlwb_list.GetItemCount() == 0);
	}
}


void Dlg_Manager_wb::f_GetAllSheet()
{
	try
	{
		_WorksheetPtr pSh1;
		nSheet = xl->ActiveWorkbook->Sheets->Count;

		CString val=L"";
		for (int i = 1; i <= nSheet; i++)
		{
			pSh1 = pWb->Worksheets->GetItem(i);
			
			MyWb[i].iNum = i;
			MyWb[i].iShow = pSh1->GetVisible();
			MyWb[i].sName = pSh1->Name;
			MyWb[i].sCode = pSh1->CodeName;
			MyWb[i].sColor = pSh1->Tab->GetColor();
			val = (LPCTSTR)(_bstr_t)MyWb[i].sColor;
			if(val==L"0") MyWb[i].sColor = WHITE;
		}

		CDialog::ActivateTopParent();
	}
	catch(...)
	{
		MessageBox(
			L"Bạn chưa khởi động Excel. Vui lòng kiểm tra lại.",
			L"Thông báo",MB_OK | MB_ICONINFORMATION);
		OnCancel();
	}
}


void Dlg_Manager_wb::f_Load_AllSheet()
{
	try
	{
		CString val=L"";
		for (int i = 0; i < nSheet; i++)
		{
			// Check
			qlwb_list.InsertItem(i,L"",0);

			// Show/Hide (value= -1,0)
			if(MyWb[i+1].iShow==-1)
			{
				qlwb_list.SetCheck(i,1);
				qlwb_list.SetItemText(i,1,L"Hiển thị");
			}

			// STT
			val.Format(L"%d",i+1);
			qlwb_list.SetItemText(i,2,val);

			// Name
			qlwb_list.SetItemText(i,3,(LPCTSTR)MyWb[i+1].sName);

			// CodeName
			qlwb_list.SetItemText(i,4,(LPCTSTR)MyWb[i+1].sCode);

			// Color
			qlwb_list.SetCellColors(i,5,MyWb[i+1].sColor,BLACK);

			// Lưu trữ màu sau
			MyWb[i+1].cAfter = MyWb[i+1].sColor;

			// Id
			val.Format(L"%d",MyWb[i+1].iNum);
			qlwb_list.SetItemText(i,6,val);

		}
	}
	catch(...){}
}


void Dlg_Manager_wb::f_Ktra_sheet()
{
	if(nSheet==0)
	{
		MessageBox(L"Không tồn tại bảng tính. Vui lòng kiểm tra lại.",L"Thông báo",MB_OK | MB_ICONINFORMATION);
		OnCancel();
	}
}

void Dlg_Manager_wb::OnBnClickedK1()
{
	CButton *m_ctlCheck = (CButton*) GetDlgItem(QLWB_K1);
	if(m_ctlCheck->GetCheck() == BST_CHECKED) f_SetCheck_List(1);
	else f_SetCheck_List(0);
}


void Dlg_Manager_wb::f_SetCheck_List(BOOLEAN f)
{
	for(int i=0;i<nSheet;i++)
	{
		qlwb_list.SetCheck(i, f);
		if(f==1) qlwb_list.SetItemText(i,1,L"Hiển thị");
		else qlwb_list.SetItemText(i,1,L"");
	}
}


void Dlg_Manager_wb::OnBnClickedB3()
{
	// Kiểm tra vị trí lựa chọn
	int nKtra=0;
	for (int i = 0; i < nSheet; i++)
	{
		if (qlwb_list.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
		{
			nKtra=1;
			break;
		}
	}

	if(nKtra==0)
	{
		MessageBox(L"Bạn chưa lựa chọn bảng tính cần đổ màu.",L"Thông báo",MB_OK | MB_ICONINFORMATION);
		return;
	}

	// Hiển thị hộp thoại Color Picker
	CColorDialog dlg; 
	if (dlg.DoModal() == IDOK) 
	{ 
		COLORREF m_color = dlg.GetColor(); 
		int n=0;
		for (int i = 0; i < nSheet; i++)
			if (qlwb_list.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED) 
			{
				qlwb_list.SetCellColors(i,5,m_color,BLACK);
				n = _wtoi(qlwb_list.GetItemText(i,6));
				MyWb[n].cAfter = m_color;
			}
	}
}


void Dlg_Manager_wb::OnNMClickL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	try
	{
		NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
		int nItem = pNMListView->iItem;
		int nSubItem = pNMListView->iSubItem;

		if(nSubItem==0 && nItem>=0)
		{
			CString val=L"";
			if(qlwb_list.GetCheck(nItem)==0) val = L"Hiển thị";
			qlwb_list.SetItemText(nItem,1,val);
		}
	}
	catch(...){}
}


void Dlg_Manager_wb::OnEnKillfocusT1()
{
	try
	{
		if(CLRow==-1 || CLCol==-1) return;

		CString ktra=L"",val=L"";
		qlwb_edit.GetWindowTextW(ktra);
		val = qlwb_list.GetItemText(CLRow,CLCol);

		// kiểm tra dữ liệu nhập vào có trùng không?
		if(ktra==val) return;

		if(CLCol==3)
		{
			// Kiểm tra trường tên bảng tính
			int nLen = ktra.GetLength();
			if(nLen==0)
			{
				MessageBox(L"Tên bảng tính không được để trống (NULL).",L"Thông báo",MB_OK | MB_ICONINFORMATION);
				qlwb_edit.SetWindowText(val);
				qlwb_list.SetItemText(CLRow,CLCol,val);
				return;
			}

			// Kiểm tra giới hạn chuỗi
			int iGHan=20,iChek=0;
			CString sKtu = L"";
			for (int i = 0; i < nLen; i++)
			{
				sKtu = ktra.Mid(i,1);
				if(_wtoi(sKtu)==0 && sKtu!=L"0")
				{
					// Tồn tại 1 ký tự là chuỗi không phải số
					iChek=1;
					break;
				}
			}

			// Nếu tất cả là số, giới hạn =20. Nếu tồn tại ít nhất 1 ký tự là chuỗi, giới hạn =31
			if(iChek) iGHan=31;

			if(nLen>iGHan)
			{
				MessageBox(L"Tên bảng tính quá dài. Vui lòng đặt tên lại.",L"Thông báo",MB_OK | MB_ICONINFORMATION);
				qlwb_edit.SetWindowText(val);
				qlwb_list.SetItemText(CLRow,CLCol,val);
				return;
			}

			// Kiểm tra các ký tự đặc biệt
			for (int i = 0; i < nLen; i++)
			{
				sKtu = ktra.Mid(i,1);
				for (int j = 0; j < 6; j++)
				{
					if(sKtu==MySymbol[j])
					{
						MessageBox(
							L"Tên bảng tính không chứa các ký tự đặc biệt (*,?,/,...).",
							L"Thông báo",MB_OK | MB_ICONINFORMATION);
						qlwb_edit.SetWindowText(val);
						qlwb_list.SetItemText(CLRow,CLCol,val);
						return;
					}
				}
			}

			// Kiểm tra sự trùng lặp
			CString s1 = ktra,s2 = L"";
			s1.MakeLower();
			for (int i = 0; i < nSheet; i++)
			{
				s2 = qlwb_list.GetItemText(i,3);
				s2.MakeLower();

				// So sánh không phân biệt chữ hoa,thường
				if(s1==s2)
				{
					MessageBox(
						L"Tên bảng tính đã tồn tại. Vui lòng đặt tên lại.",
						L"Thông báo",MB_OK | MB_ICONINFORMATION);
					qlwb_edit.SetWindowText(val);
					qlwb_list.SetItemText(CLRow,CLCol,val);
					return;
				}
			}
		}
		else if(CLCol==4)
		{
			// Tạm thời khóa tính năng đổi codename
			return;

			if(iCode==0)
			{
				iCode=1;
				MessageBox(
					L"Bạn chỉ đổi Codename khi thực sự cần thiết.",
					L"Cảnh báo",MB_OK | MB_ICONWARNING);
				qlwb_edit.SetWindowText(val);
				qlwb_list.SetItemText(CLRow,CLCol,val);
				return;

			}

			// Kiểm tra trường Codename tính rỗng(NULL)
			int nLen = ktra.GetLength();
			if(nLen==0)
			{
				MessageBox(
					L"Tên Codename không được để trống (NULL).",
					L"Thông báo",MB_OK | MB_ICONINFORMATION);
				qlwb_edit.SetWindowText(val);
				qlwb_list.SetItemText(CLRow,CLCol,val);
				return;
			}

			// Kiểm tra ký tự đầu tiên
			CString sKtu = ktra.Left(1);
			if(_wtoi(ktra)>0 || ktra==L"0")
			{
				MessageBox(
					L"Ký tự đầu tiên phải là chữ cái.",
					L"Thông báo",MB_OK | MB_ICONINFORMATION);
				qlwb_edit.SetWindowText(val);
				qlwb_list.SetItemText(CLRow,CLCol,val);
				return;
			}

			// Kiểm tra giới hạn ký tự
			if(nLen>31)
			{
				MessageBox(L"Tên Codename quá dài.",L"Thông báo",MB_OK | MB_ICONINFORMATION);
				qlwb_edit.SetWindowText(val);
				qlwb_list.SetItemText(CLRow,CLCol,val);
				return;
			}

			// Kiểm tra ký tự SPACE
			int _pos = ktra.Find(L" ");
			if(_pos>=0)
			{
				MessageBox(L"Tên Codename phải viết liền.",L"Thông báo",MB_OK | MB_ICONINFORMATION);
				qlwb_edit.SetWindowText(val);
				qlwb_list.SetItemText(CLRow,CLCol,val);
				return;
			}

			// Kiểm tra các ký tự đặc biệt
			int iABC=0;
			CString s1 = ktra;
			s1.MakeLower();
			for (int i = 0; i < nLen; i++)
			{
				iABC=0;
				sKtu = s1.Mid(i,1);
				for (int j = 0; j < 36; j++)
				{
					if(sKtu==MySymABC[j])
					{
						iABC=1;
						break;
					}
				}

				if(iABC==0)
				{
					MessageBox(
						L"Tên Codename chỉ chứa chữ cái và số, không chứa ký tự đặc biệt.",
						L"Thông báo",MB_OK | MB_ICONINFORMATION);
					qlwb_edit.SetWindowText(val);
					qlwb_list.SetItemText(CLRow,CLCol,val);
					return;
				}
			}

			// Kiểm tra sự trùng lặp
			CString s2 = L"";
			for (int i = 0; i < nSheet; i++)
			{
				s2 = qlwb_list.GetItemText(i,4);
				s2.MakeLower();

				// So sánh không phân biệt chữ hoa,thường
				if(s1==s2)
				{
					MessageBox(
						L"Tên Codename đã tồn tại. Vui lòng đặt tên lại.",
						L"Thông báo",MB_OK | MB_ICONINFORMATION);
					qlwb_edit.SetWindowText(val);
					qlwb_list.SetItemText(CLRow,CLCol,val);
					return;
				}
			}
		}
	}
	catch(...){}
}


void Dlg_Manager_wb::OnBnClickedB4()
{
	try
	{
		CString val=L"";
		POSITION pos=qlwb_list.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			int nId=1,nItem=0;
			for (int i = 0; i < nSheet; i++)
				if (qlwb_list.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					nItem=i;
					break;
				}
			
			if(nItem>0)
			{
				// Check
				int iChk0 = qlwb_list.GetCheck(nItem);
				int iChk1 = qlwb_list.GetCheck(nItem-1);

				qlwb_list.SetCheck(nItem,iChk1);
				qlwb_list.SetCheck(nItem-1,iChk0);

				// Show/Hide
				val = qlwb_list.GetItemText(nItem,1);
				qlwb_list.SetItemText(nItem,1,qlwb_list.GetItemText(nItem-1,1));
				qlwb_list.SetItemText(nItem-1,1,val);

				// Name
				val = qlwb_list.GetItemText(nItem,3);
				qlwb_list.SetItemText(nItem,3,qlwb_list.GetItemText(nItem-1,3));
				qlwb_list.SetItemText(nItem-1,3,val);

				// Codename
				val = qlwb_list.GetItemText(nItem,4);
				qlwb_list.SetItemText(nItem,4,qlwb_list.GetItemText(nItem-1,4));
				qlwb_list.SetItemText(nItem-1,4,val);

				// Color
				nId = _wtoi(qlwb_list.GetItemText(nItem,6));
				COLORREF cLr0 = MyWb[nId].cAfter;

				nId = _wtoi(qlwb_list.GetItemText(nItem-1,6));
				COLORREF cLr1 = MyWb[nId].cAfter;
				
				qlwb_list.SetCellColors(nItem,5,cLr1,BLACK);
				qlwb_list.SetCellColors(nItem-1,5,cLr0,BLACK);

				// ID
				val = qlwb_list.GetItemText(nItem,6);
				qlwb_list.SetItemText(nItem,6,qlwb_list.GetItemText(nItem-1,6));
				qlwb_list.SetItemText(nItem-1,6,val);

				// Vị trí lựa chọn mới
				qlwb_list.SetItemState(nItem,0,LVIS_SELECTED);
				qlwb_list.SetItemState(nItem-1,LVIS_SELECTED,LVIS_SELECTED);
			}
		}
	}
	catch(...){}
}


void Dlg_Manager_wb::OnBnClickedB5()
{
	try
	{
		CString val=L"";
		POSITION pos=qlwb_list.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			int nId=1,nItem=0;
			for (int i = 0; i < nSheet; i++)
				if (qlwb_list.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					nItem=i;
					break;
				}

			if(nItem<nSheet-1)
			{
				// Check
				int iChk0 = qlwb_list.GetCheck(nItem);
				int iChk1 = qlwb_list.GetCheck(nItem+1);

				qlwb_list.SetCheck(nItem,iChk1);
				qlwb_list.SetCheck(nItem+1,iChk0);

				// Show/Hide
				val = qlwb_list.GetItemText(nItem,1);
				qlwb_list.SetItemText(nItem,1,qlwb_list.GetItemText(nItem+1,1));
				qlwb_list.SetItemText(nItem+1,1,val);

				// Name
				val = qlwb_list.GetItemText(nItem,3);
				qlwb_list.SetItemText(nItem,3,qlwb_list.GetItemText(nItem+1,3));
				qlwb_list.SetItemText(nItem+1,3,val);

				// Codename
				val = qlwb_list.GetItemText(nItem,4);
				qlwb_list.SetItemText(nItem,4,qlwb_list.GetItemText(nItem+1,4));
				qlwb_list.SetItemText(nItem+1,4,val);

				// Color
				nId = _wtoi(qlwb_list.GetItemText(nItem,6));
				COLORREF cLr0 = MyWb[nId].cAfter;

				nId = _wtoi(qlwb_list.GetItemText(nItem+1,6));
				COLORREF cLr1 = MyWb[nId].cAfter;
				
				qlwb_list.SetCellColors(nItem,5,cLr1,BLACK);
				qlwb_list.SetCellColors(nItem+1,5,cLr0,BLACK);

				// ID
				val = qlwb_list.GetItemText(nItem,6);
				qlwb_list.SetItemText(nItem,6,qlwb_list.GetItemText(nItem+1,6));
				qlwb_list.SetItemText(nItem+1,6,val);

				// Vị trí lựa chọn mới
				qlwb_list.SetItemState(nItem,0,LVIS_SELECTED);
				qlwb_list.SetItemState(nItem+1,LVIS_SELECTED,LVIS_SELECTED);
			}
		}
	}
	catch(...){}
}


void Dlg_Manager_wb::OnBnClickedB6()
{
	try
	{
		CString val=L"";
		POSITION pos=qlwb_list.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			int nId=1,vtri1,vtri2,k=0;
			for (int i = 0; i < nSheet; i++)
				if (qlwb_list.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					k++;
					if(k==1) vtri1=i;
					else if(k==2)
					{
						vtri2=i;
						break;
					}
				}

			if(k==2)
			{
				// Check
				int iChk0 = qlwb_list.GetCheck(vtri1);
				int iChk1 = qlwb_list.GetCheck(vtri2);

				qlwb_list.SetCheck(vtri1,iChk1);
				qlwb_list.SetCheck(vtri2,iChk0);

				// Show/Hide
				val = qlwb_list.GetItemText(vtri1,1);
				qlwb_list.SetItemText(vtri1,1,qlwb_list.GetItemText(vtri2,1));
				qlwb_list.SetItemText(vtri2,1,val);

				// Name
				val = qlwb_list.GetItemText(vtri1,3);
				qlwb_list.SetItemText(vtri1,3,qlwb_list.GetItemText(vtri2,3));
				qlwb_list.SetItemText(vtri2,3,val);

				// Codename
				val = qlwb_list.GetItemText(vtri1,4);
				qlwb_list.SetItemText(vtri1,4,qlwb_list.GetItemText(vtri2,4));
				qlwb_list.SetItemText(vtri2,4,val);

				// Color
				nId = _wtoi(qlwb_list.GetItemText(vtri1,6));
				COLORREF cLr0 = MyWb[nId].cAfter;

				nId = _wtoi(qlwb_list.GetItemText(vtri2,6));
				COLORREF cLr1 = MyWb[nId].cAfter;
				
				qlwb_list.SetCellColors(vtri1,5,cLr1,BLACK);
				qlwb_list.SetCellColors(vtri2,5,cLr0,BLACK);

				// ID
				val = qlwb_list.GetItemText(vtri1,6);
				qlwb_list.SetItemText(vtri1,6,qlwb_list.GetItemText(vtri2,6));
				qlwb_list.SetItemText(vtri2,6,val);
			}
		}
	}
	catch(...){}
}


void Dlg_Manager_wb::OnBnClickedB7()
{
	OnCancel();
}

void Dlg_Manager_wb::OnLvnKeydownL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);
	if(pLVKeyDow->wVKey == VK_TAB) OnBnClickedB6();
	else if(pLVKeyDow->wVKey == VK_UP) OnBnClickedB4();
	else if(pLVKeyDow->wVKey == VK_DOWN) OnBnClickedB5();
	*pResult = 0;
}


void Dlg_Manager_wb::OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	try
	{
		NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
		int nItem = pNMListView->iItem;
		int nSubItem = pNMListView->iSubItem;
		if(nItem<0) return;
		if(nSubItem==1)
		{
			CString val = qlwb_list.GetItemText(nItem,1);
			if(val==L"")
			{
				qlwb_list.SetCheck(nItem,1);
				qlwb_list.SetItemText(nItem,1,L"Hiển thị");
			}
			else
			{
				qlwb_list.SetCheck(nItem,0);
				qlwb_list.SetItemText(nItem,1,L"");
			}
		}
		else if(nSubItem==5)
		{
			CColorDialog dlg; 
			if (dlg.DoModal() == IDOK) 
			{ 
				COLORREF m_color = dlg.GetColor(); 
				int n=0;
				for (int i = 0; i < nSheet; i++)
					if (qlwb_list.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED) 
					{
						qlwb_list.SetCellColors(i,5,m_color,BLACK);
						n = _wtoi(qlwb_list.GetItemText(i,6));
						MyWb[n].cAfter = m_color;
					}
			}
		}
	}
	catch(...){}
}


void Dlg_Manager_wb::OnNMRClickL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	try
	{
		NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
		int nItem = pNMListView->iItem;
		int nSubItem = pNMListView->iSubItem;
		if(nItem<0) return;
		if(nSubItem==5)
		{
			CColorDialog dlg; 
			if (dlg.DoModal() == IDOK) 
			{ 
				COLORREF m_color = dlg.GetColor(); 
				int n=0;
				for (int i = 0; i < nSheet; i++)
					if (qlwb_list.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED) 
					{
						qlwb_list.SetCellColors(i,5,m_color,BLACK);
						n = _wtoi(qlwb_list.GetItemText(i,6));
						MyWb[n].cAfter = m_color;
					}
			}
		}
	}
	catch(...){}
}


void Dlg_Manager_wb::f_FindSheet()
{
	try
	{
		POSITION pos=qlwb_list.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			for (int i = 0; i < qlwb_list.GetItemCount();i++ )
			if (qlwb_list.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
			{
				xl->PutScreenUpdating(VARIANT_FALSE);

				int n = _wtoi(qlwb_list.GetItemText(i,6));
				pSheet = pWb->Worksheets->GetItem(n);

				int iVisible = pSheet->GetVisible();
				if (iVisible!=-1) pSheet->PutVisible(XlSheetVisibility::xlSheetVisible);
				pSheet->Activate();

				break;
			}
		}
	}
	catch(...){}
}
