#include "Dlg_importpic.h"
#include "cls_main.h"

Dlg_importpic::Dlg_importpic(CWnd* pParent /*=NULL*/)
	: CDialog(Dlg_importpic::IDD, pParent)
{
	
}

void Dlg_importpic::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IMPPIC_T1, txt1);
	DDX_Control(pDX, IMPPIC_B1, btn[1]);
	DDX_Control(pDX, IMPPIC_B2, btn[2]);
	DDX_Control(pDX, IMPPIC_B3, btn[3]);
	DDX_Control(pDX, IMPPIC_B4, btn[4]);
	DDX_Control(pDX, IMPPIC_L1, list1);
	DDX_Control(pDX, IMPPIC_P1, m_picCtrl);
	DDX_Control(pDX, IMPPIC_G1, stt1);
	DDX_Control(pDX, IMPPIC_G2, stt2);
}

BEGIN_MESSAGE_MAP(Dlg_importpic, CDialog)
	ON_BN_CLICKED(IMPPIC_B1, &Dlg_importpic::OnBnClickedB1)
	ON_BN_CLICKED(IMPPIC_B2, &Dlg_importpic::OnBnClickedB2)
	ON_BN_CLICKED(IMPPIC_B3, &Dlg_importpic::OnBnClickedB3)
	ON_NOTIFY(NM_CLICK, IMPPIC_L1, &Dlg_importpic::OnNMClickL1)
	ON_BN_CLICKED(IMPPIC_B4, &Dlg_importpic::OnBnClickedB4)
	ON_WM_SIZE()
	ON_WM_SIZING()

	
END_MESSAGE_MAP()

BEGIN_EASYSIZE_MAP(Dlg_importpic)
	// EASYSIZE(control,left,top,right,bottom,options)
	// options = 1 -> không thay đổi, =0 có thay đổi
	// ES_BORDER = ghim (thay đổi theo)
	// ES_KEEPSIZE = giữ nguyên ban đầu

	// Trái
	EASYSIZE(IMPPIC_T1,ES_BORDER,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)	
	EASYSIZE(IMPPIC_G1,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_BORDER,0)	
	EASYSIZE(IMPPIC_L1,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_BORDER,0)	
	EASYSIZE(IMPPIC_G2,ES_BORDER,ES_BORDER,ES_BORDER,ES_BORDER,0)	
	EASYSIZE(IMPPIC_P1,ES_BORDER,ES_BORDER,ES_BORDER,ES_BORDER,0)	
	EASYSIZE(IMPPIC_B1,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(IMPPIC_B2,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(IMPPIC_B3,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)
	EASYSIZE(IMPPIC_B4,ES_KEEPSIZE,ES_BORDER,ES_BORDER,ES_KEEPSIZE,0)

END_EASYSIZE_MAP

BOOL Dlg_importpic::OnInitDialog()
{
CDialog::OnInitDialog();
	// Load Icon dự toán
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	_fullpath = L"";
	list1.InsertColumn(0,L"Tên tệp tin",LVCFMT_LEFT,175);
	list1.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP | LVS_EX_CHECKBOXES);

	f_Load_List();
	f_Load_ToolTip();

	INIT_EASYSIZE;

	return TRUE; 
}

// Bắt sự kiện click Enter
BOOL Dlg_importpic::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
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

void Dlg_importpic::OnSize(UINT nType, int cx, int cy) 
{
	CDialog::OnSize(nType, cx, cy);
	UPDATE_EASYSIZE;
}

void Dlg_importpic::OnSizing(UINT fwSide, LPRECT pRect) 
{
	CDialog::OnSizing(fwSide, pRect);
	EASYSIZE_MINSIZE(338,360,fwSide,pRect);
}


void Dlg_importpic::f_Load_ToolTip()
{
	CEdit * pbt0 = (CEdit *)GetDlgItem(IMPPIC_T1);

	m_ToolTips.Create(this);
	m_ToolTips.SetMaxTipWidth(600);
	m_ToolTips.SetDelayTime(900);
		
	m_ToolTips.AddTool(pbt0, L"Đường dẫn chứa tập tin ảnh cần sử dụng.");
}


void Dlg_importpic::OnBnClickedB1()
{
	_bienpic=L"";
	CString szval=L"";
	for (int i = 0; i < list1.GetItemCount(); i++)
	{
		if((int)list1.GetCheck(i)==1)
		{
			szval = list1.GetItemText(i,0);
			_fullpath.Format(L"%s\\%s|",sFolder,szval);
			_bienpic+=_fullpath;
		}
	}

	if(_bienpic!=L"") CDialog::OnOK();
	else _s(L"Vui lòng tích chọn tệp tin ảnh sử dụng.");
}


void Dlg_importpic::f_Delete_List()
{
	list1.DeleteAllItems();
	ASSERT(list1.GetItemCount() == 0);
}


void Dlg_importpic::f_Load_List()
{
	try
	{
		if (sFolder == L"")
		{
			CString spathDefault = _fGPathFolder(L"ToolQLCL.dll");
			if (spathDefault.Right(1) != L"\\") spathDefault += L"\\";
			sFolder.Format(L"%s%s%s", spathDefault, _pathDlieu, _pathImage);
		}

		f_Add_Csdl(sFolder,L".jpg");
		int q = _finder.GetFileCount();

		if (q == 0)
		{
			sFolder = getDesktopDir();
			f_Add_Csdl(sFolder, L".jpg");
			q = _finder.GetFileCount();
		}
		
		msg.Format(L"Danh sách tệp tin ảnh (%d)",q);
		stt1.SetWindowText(msg);
		txt1.SetWindowText(sFolder);

		CString val1;
		if(q>0)
		{
			for (int i = 0; i < q; i++)
			{
				val1 = _finder.GetFilePath(i).GetFileName();
				list1.InsertItem(i,val1,0);
			}

			_fullpath.Format(L"%s\\%s",sFolder,list1.GetItemText(0,0));
			m_picCtrl.Load(_fullpath);
			msg.Format(L"Chi tiết tập tin (%s)",list1.GetItemText(0,0));
			stt2.SetWindowText(msg);
		}
	}
	catch(_com_error & error){_s(L"Lỗi xác định danh sách tập tin ảnh.");}
}


void Dlg_importpic::OnBnClickedB2()
{
	OnCancel();
}

void Dlg_importpic::OnBnClickedB3()
{
	try
	{
		CString spt = sFolder;
		LPCTSTR lpszTitle = L"Chọn đường dẫn thư mục";
		UINT uFlags = BIF_RETURNONLYFSDIRS | BIF_USENEWUI;
		CFolderDialog dlgRoot(lpszTitle, spt, this, uFlags);
		if (dlgRoot.DoModal() == IDOK)
		{
			sFolder = dlgRoot.GetFolderPath();
			if (sFolder.Right(1) != L"\\") sFolder += L"\\";
			
			txt1.SetWindowText(sFolder);
			f_Add_Csdl(sFolder, L".jpg");
			int q = _finder.GetFileCount();
			if (q > 0)
			{
				f_Delete_List();
				CString szval=_T("");
				for (int i = 0; i < q; i++)
				{
					szval = _finder.GetFilePath(i).GetFileName();
					list1.InsertItem(i, szval, 0);
				}

				szval = _finder.GetFilePath(0).GetFileName();
				msg.Format(L"Chi tiết tệp tin (%s)", szval);
				stt2.SetWindowText(msg);

				_fullpath.Format(L"%s\\%s", sFolder, szval);
				m_picCtrl.Load(_fullpath);
			}

			msg.Format(L"Danh sách tệp tin ảnh (%d)", q);
			stt1.SetWindowText(msg);
		}
	}
	catch(_com_error & error){_s(L"Lỗi đường dẫn.");}
}


void Dlg_importpic::f_Add_Csdl(CString m_sBaseFolder, CString m_sFileMask)
{
	try
	{
		CFileFinder::CFindOpts	opts;

		// Set CFindOpts object
		opts.sBaseFolder = m_sBaseFolder;
		opts.sFileMask.Format(L"*%s*", m_sFileMask);
		opts.FindNormalFiles();

		_finder.RemoveAll();
		_finder.Find(opts);

	}
	catch(_com_error & error){_s(L"Lỗi xác định dữ liệu jpg.");}
}


void Dlg_importpic::OnNMClickL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	int nClick = pNMListView->iItem;

	CString p0 = list1.GetItemText(nClick,0);
	if(p0!=L"")
	{
		msg.Format(L"Chi tiết tập tin (%s)",p0);
		stt2.SetWindowText(msg);

		_fullpath.Format(L"%s\\%s",sFolder,p0);
		m_picCtrl.Load(_fullpath);
	}
}


void Dlg_importpic::OnBnClickedB4()
{
	// Hiển thị thư mục đường dẫn
	ShellExecute(NULL, L"open",sFolder, NULL, NULL, SW_SHOWMAXIMIZED);
}
