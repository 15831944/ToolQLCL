#include "Dlg_ThuvienTL.h"

// Class tra cứu phụ lục vữa
Dlg_ThuvienTL::Dlg_ThuvienTL(CWnd* pParent /*=NULL*/)
	: CDialog(Dlg_ThuvienTL::IDD, pParent)
{
	
}

void Dlg_ThuvienTL::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, TVTL_T1, txtKey);
	DDX_Control(pDX, TVTL_L1, list1);
	DDX_Control(pDX, TVTL_G2, sttic);
}

BEGIN_MESSAGE_MAP(Dlg_ThuvienTL, CDialog)
	

	ON_WM_SIZE()
	ON_WM_SIZING()
	ON_BN_CLICKED(TVTL_B4, &Dlg_ThuvienTL::OnBnClickedB4)
	ON_BN_CLICKED(TVTL_B2, &Dlg_ThuvienTL::OnBnClickedB2)
	ON_BN_CLICKED(TVTL_B3, &Dlg_ThuvienTL::OnBnClickedB3)
	ON_BN_CLICKED(TVTL_B1, &Dlg_ThuvienTL::OnBnClickedB1)
	ON_NOTIFY(NM_DBLCLK, TVTL_L1, &Dlg_ThuvienTL::OnNMDblclkL1)
END_MESSAGE_MAP()

BEGIN_EASYSIZE_MAP(Dlg_ThuvienTL)
	// EASYSIZE(control,left,top,right,bottom,options)
	// options = 1 -> không thay đổi, =0 có thay đổi
	// ES_BORDER = ghim (thay đổi theo)
	// ES_KEEPSIZE = giữ nguyên ban đầu

	EASYSIZE(TVTL_G1,ES_BORDER,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)	
	EASYSIZE(TVTL_T1,ES_BORDER,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)
	EASYSIZE(TVTL_B1,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)

	EASYSIZE(TVTL_L1,ES_BORDER,ES_BORDER,ES_BORDER,ES_BORDER,0)
	EASYSIZE(TVTL_G2,ES_BORDER,ES_BORDER,ES_BORDER,ES_BORDER,0)

	EASYSIZE(TVTL_B2,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)
	EASYSIZE(TVTL_B3,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)
	EASYSIZE(TVTL_B4,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
END_EASYSIZE_MAP

BOOL Dlg_ThuvienTL::OnInitDialog()
{
CDialog::OnInitDialog();
	// Load Icon dự toán
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	// Từ khóa hiển thị hướng dẫn
	txtKey.SetCueBanner(_T("Nhập từ khóa tìm kiếm ..."), TRUE);

	szFolder = L"C:\\QLCL GXD\\Dulieu\\Tailieu";
	CString pDlg = _T("[Thư viện tài liệu] ") + szFolder;
	this->SetWindowTextW(pDlg);

	f_Style_tieuchuan();
	f_Load_list_tieuchuan();

	INIT_EASYSIZE;

	return TRUE; 
}

// Bắt sự kiện click Enter
BOOL Dlg_ThuvienTL::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if(pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(TVTL_T1))
	{
		OnBnClickedB1();
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		CDialog:OnCancel();
		return TRUE; 
	}

	return FALSE;
}

void Dlg_ThuvienTL::OnSize(UINT nType, int cx, int cy) 
{
	CDialog::OnSize(nType, cx, cy);
	UPDATE_EASYSIZE;
}

void Dlg_ThuvienTL::OnSizing(UINT fwSide, LPRECT pRect) 
{
	CDialog::OnSizing(fwSide, pRect);
	EASYSIZE_MINSIZE(560,250,fwSide,pRect);	// chiều rộng + chiều cao
}


void Dlg_ThuvienTL::f_Style_tieuchuan()
{
	// Định dạng kiểu List Control
	list1.InsertColumn(0,L"",LVCFMT_CENTER,0);
	list1.InsertColumn(1,L"Tên tài liệu",LVCFMT_LEFT,300);
	list1.InsertColumn(2,L"Xem chi tiết",LVCFMT_CENTER,100);
	list1.InsertColumn(3,L"Kiểu",LVCFMT_CENTER,0);

	list1.SetColumnColors(2,RGB(255,255,255),RGB(0,51,204));
	list1.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);
}


void Dlg_ThuvienTL::f_delete_list()
{
	list1.DeleteAllItems();
	ASSERT(list1.GetItemCount() == 0);
}


void Dlg_ThuvienTL::f_Load_list_tieuchuan()
{
	try
	{
		// * = lựa chọn tất cả các file
		f_Add_Csdl(szFolder,L"*");

		slTL=0;
		CString sVal=L"",sKtra=L"";
		int nCount = _finder.GetFileCount();
		if(nCount>0)
		{
			for(int i = 0; i<nCount; i++)
			{
				sVal = _finder.GetFilePath(i).GetFileName();
				sKtra = sVal.Right(4);

				// Kiểm tra các tài liệu thỏa mãn
				if(sKtra==L".doc" || sKtra==L"docx" || sKtra==L"docm" 
					|| sKtra==L".xls" || sKtra==L"xlsx" || sKtra==L"xlsm" || sKtra==L".csv" 
					|| sKtra==L".pdf" || sKtra==L".xps" || sKtra==L".txt")
				{
					// Lưu vào dữ liệu để tìm kiếm
					sList[slTL+1] = sVal;

					// Load lên list
					list1.InsertItem(slTL,L"",0);
					list1.SetItemText(slTL,1,sVal);
					list1.SetItemText(slTL,2,L"Xem chi tiết");
					list1.SetItemText(slTL,3,L"");

					slTL++;
				}
			}
		}

		sVal.Format(L"Tìm thấy (%d) tài liệu",slTL);
		sttic.SetWindowText(sVal);
	}
	catch(...){_s(L"Lỗi xác định thư viện tài liệu.");}
}


void Dlg_ThuvienTL::f_Add_Csdl(CString m_sBaseFolder, CString m_sFileMask)
{
	try
	{
		CFileFinder::CFindOpts	opts;

		//UpdateData();

		// Set CFindOpts object
		opts.sBaseFolder = m_sBaseFolder;
		opts.sFileMask.Format(L"*%s*", m_sFileMask);
		opts.FindNormalFiles();

		_finder.RemoveAll();
		_finder.Find(opts);
	}
	catch(...){_s(L"Lỗi load dữ liệu (all)");}
}



void Dlg_ThuvienTL::OnBnClickedB4()
{
	OnCancel();
}


void Dlg_ThuvienTL::OnBnClickedB2()
{
	try
	{
		BROWSEINFO bi = {0};
		bi.lpszTitle = _T("Lựa chọn đường dẫn chứa thư viện tài liệu...");
		LPITEMIDLIST pidl = SHBrowseForFolder (&bi);
		if (pidl != 0)
		{
			// get the name of the folder
			TCHAR path[MAX_PATH];
			SHGetPathFromIDList (pidl, path);
			
			// free memory used
			IMalloc * imalloc = 0;
			if (SUCCEEDED(SHGetMalloc (&imalloc)))
			{
				imalloc->Free (pidl);
				imalloc->Release ();
				szFolder.Format(L"%s",path);

				CString pDlg = _T("[Thư viện tài liệu] ") + szFolder;
				this->SetWindowTextW(pDlg);

				f_delete_list();
				f_Load_list_tieuchuan();
			}
		}
	}
	catch(...){_s(L"Lỗi chọn đường dẫn.");}	
}


void Dlg_ThuvienTL::OnBnClickedB3()
{
	// Hiển thị thư mục đường dẫn
	ShellExecute(NULL, L"open",szFolder, NULL, NULL, SW_SHOWMAXIMIZED);
}


void Dlg_ThuvienTL::OnBnClickedB1()
{
	// TODO: Add your control notification handler code here
}


void Dlg_ThuvienTL::OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	int nItem = pNMListView->iItem;
	int nSubItem = pNMListView->iSubItem;

	if(nItem>=0 && nSubItem==2)
	{
		CString sVal = list1.GetItemText(nItem,1);
		CString sPath = L"";
		if(sPath.Right(1)==L"\\") sPath = szFolder + sVal;
		else sPath.Format(L"%s\\%s",szFolder,sVal);
		
		ShellExecute(NULL, L"open", sPath, NULL, NULL, SW_SHOWNORMAL);
	}

	*pResult = 0;
}
