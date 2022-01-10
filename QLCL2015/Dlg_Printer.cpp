#include "Dlg_Printer.h"

Dlg_Printer::Dlg_Printer(CWnd* pParent /*=NULL*/)
	: CDialog(Dlg_Printer::IDD, pParent)
{
	
}

void Dlg_Printer::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, PRT_L1, m_List);
	DDX_Control(pDX, PRT_CK1, m_Check);
}

BEGIN_MESSAGE_MAP(Dlg_Printer, CDialog)
	ON_WM_SIZE()
	ON_WM_SIZING()
	ON_BN_CLICKED(PRT_CK1, &Dlg_Printer::OnBnClickedCk1)
	ON_BN_CLICKED(PRT_B1, &Dlg_Printer::OnBnClickedB1)
	ON_BN_CLICKED(PRT_B3, &Dlg_Printer::OnBnClickedB3)
	ON_BN_CLICKED(PRT_B2, &Dlg_Printer::OnBnClickedB2)
	ON_NOTIFY(NM_DBLCLK, PRT_L1, &Dlg_Printer::OnNMDblclkL1)
END_MESSAGE_MAP()

BEGIN_EASYSIZE_MAP(Dlg_Printer)
	EASYSIZE(PRT_G1,ES_BORDER,ES_BORDER,ES_BORDER,ES_BORDER,0)
	EASYSIZE(PRT_L1,ES_BORDER,ES_BORDER,ES_BORDER,ES_BORDER,0)
	EASYSIZE(PRT_CK1,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,0)
	EASYSIZE(PRT_B1,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)
	EASYSIZE(PRT_B2,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(PRT_B3,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
END_EASYSIZE_MAP

BOOL Dlg_Printer::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Load Icon
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	f_StyleList();
	f_GetAllSheet();
	m_Check.SetCheck(1);

	INIT_EASYSIZE;

	return TRUE; 
}

// Bắt sự kiện click Enter
BOOL Dlg_Printer::PreTranslateMessage(MSG* pMsg)
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
		catch(...){}*/

		CDialog::OnCancel();
		return TRUE; 
	}

	return FALSE;
}

void Dlg_Printer::OnSize(UINT nType, int cx, int cy) 
{
	CDialog::OnSize(nType, cx, cy);
	UPDATE_EASYSIZE;
}

void Dlg_Printer::OnSizing(UINT fwSide, LPRECT pRect) 
{
	CDialog::OnSizing(fwSide, pRect);
	EASYSIZE_MINSIZE(320,240,fwSide,pRect);
}


void Dlg_Printer::f_GetAllSheet()
{
	try
	{
		CString sName,sCode;
		int dem=0,nSheet = xl->ActiveWorkbook->Sheets->Count;
		for (int i = 1; i <= nSheet; i++)
		{	
			pSh1 = pWb->Worksheets->GetItem(i);
			sName = (LPCTSTR)pSh1->Name;
			sCode = (LPCTSTR)pSh1->CodeName;

			if(sCode!=L"shConfig" && sCode!=L"shName" 
				&& sCode.Left(8)!=L"shHSNTVL" && sCode.Left(8)!=L"shHSNTCV" && sCode.Left(8)!=L"shHSNTGD")
			{
				// Load dữ liệu vào list
				m_List.InsertItem(dem,L"",0);
				m_List.SetCheck(dem,1);
				m_List.SetItemText(dem,1,sName);
				m_List.SetItemText(dem,2,sCode);

				dem++;
			}
		}
	}
	catch(...){}
}


void Dlg_Printer::f_StyleList()
{
	m_List.InsertColumn(0,L"",LVCFMT_CENTER,22);
	m_List.InsertColumn(1,L"Tên bảng tính",LVCFMT_LEFT,250);
	m_List.InsertColumn(2,L"Codename",LVCFMT_LEFT,0);
	m_List.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_CHECKBOXES);
}


void Dlg_Printer::OnBnClickedCk1()
{
	CButton *m_ctlCheck = (CButton*) GetDlgItem(PRT_CK1);
	if(m_ctlCheck->GetCheck() == BST_CHECKED) f_SetCheck_List(1);
	else f_SetCheck_List(0);
}

void Dlg_Printer::f_SetCheck_List(BOOLEAN f)
{
	for(int i=0;i<m_List.GetItemCount();i++) m_List.SetCheck(i, f);
}


void Dlg_Printer::OnBnClickedB1()
{
	try
	{
		xl->PutStatusBar((_bstr_t)L"Đang tiến hành chọn máy in. Vui lòng đợi trong giây lát..");
		xl->Application->Dialogs->GetItem(xlDialogPrinterSetup)->Show();
		CDialog::ActivateTopParent();
		xl->PutStatusBar((_bstr_t)L"Ready");
	}
	catch(...){}
}


void Dlg_Printer::OnBnClickedB3()
{
	OnCancel();
}


void Dlg_Printer::OnBnClickedB2()
{
	try
	{
		// Hàm check bản quyền
		if(f_Check_ban_quyen()!=1) return;

		CString val = L"Bạn có chắc chắn in các bảng tính đã chọn? \nNhấn YES để in ngay, NO để xem trước khi in.";
		int iVisible,nCount = m_List.GetItemCount();
		int n = _yn(val,1);
		if(n==6 || n==7)
		{
			OnOK();
			xl->PutScreenUpdating(VARIANT_FALSE);
			for (int i = 0; i < nCount; i++)
			{
				if(m_List.GetCheck(i)==1)
				{
					val = m_List.GetItemText(i,1);
					pSh1 = pWb->Sheets->GetItem((_bstr_t)val);
					iVisible = pSh1->GetVisible();
					if (iVisible!=-1) pSh1->PutVisible(XlSheetVisibility::xlSheetVisible);

					if(n==6) pSh1->PrintOut();
					else pSh1->PrintPreview();
				}
			}
		}
	}
	catch(...){}
}

void Dlg_Printer::OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	int nItem = pNMListView->iItem;
	int nSubItem = pNMListView->iSubItem;

	if(nSubItem>0)
	{
		int nCheck = m_List.GetCheck(nItem);
		if(nCheck==1) m_List.SetCheck(nItem,0);
		else m_List.SetCheck(nItem,1);		
	}
}
