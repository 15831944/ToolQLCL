// Dlg_Qly_image_nky message handlers
#include "Dlg_Qly_image_nky.h"

Dlg_Qly_image_nky::Dlg_Qly_image_nky(CWnd* pParent /*=NULL*/)
	: CDialog(Dlg_Qly_image_nky::IDD, pParent)
{
	
}

void Dlg_Qly_image_nky::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, GLIM_T1, txt1);
	DDX_Control(pDX, GLIM_L1, listIMG);
	DDX_Control(pDX, GLIM_S1, m_picCtrl);
	DDX_Control(pDX, GLIM_G2, stt2);
}

BEGIN_MESSAGE_MAP(Dlg_Qly_image_nky, CDialog)
	ON_WM_SIZE()
	ON_WM_SIZING()	
	ON_BN_CLICKED(GLIM_B8, &Dlg_Qly_image_nky::OnBnClickedB8)
	ON_BN_CLICKED(GLIM_B7, &Dlg_Qly_image_nky::OnBnClickedB7)
	ON_NOTIFY(NM_CLICK, GLIM_L1, &Dlg_Qly_image_nky::OnNMClickL1)
	ON_EN_KILLFOCUS(GLIM_T1, &Dlg_Qly_image_nky::OnEnKillfocusT1)
	ON_BN_CLICKED(GLIM_B2, &Dlg_Qly_image_nky::OnBnClickedB2)
	ON_BN_CLICKED(GLIM_B1, &Dlg_Qly_image_nky::OnBnClickedB1)
	ON_BN_CLICKED(GLIM_B3, &Dlg_Qly_image_nky::OnBnClickedB3)
	ON_BN_CLICKED(GLIM_B4, &Dlg_Qly_image_nky::OnBnClickedB4)
END_MESSAGE_MAP()

BEGIN_EASYSIZE_MAP(Dlg_Qly_image_nky)
	// EASYSIZE(control,left,top,right,bottom,options)
	// options = 1 -> không thay đổi, =0 có thay đổi
	// ES_BORDER = ghim (thay đổi theo)
	// ES_KEEPSIZE = giữ nguyên ban đầu

	// Trái
	EASYSIZE(GLIM_G1,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_BORDER,0)
	EASYSIZE(GLIM_L1,ES_BORDER,ES_BORDER,ES_KEEPSIZE,ES_BORDER,0)
	EASYSIZE(GLIM_B1,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)
	EASYSIZE(GLIM_B2,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)
	EASYSIZE(GLIM_B3,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)
	EASYSIZE(GLIM_B4,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)

	EASYSIZE(GLIM_G2,ES_BORDER,ES_BORDER,ES_BORDER,ES_BORDER,0)
	EASYSIZE(GLIM_S1,ES_BORDER,ES_BORDER,ES_BORDER,ES_BORDER,0)

	EASYSIZE(GLIM_G3,ES_BORDER,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(GLIM_T1,ES_BORDER,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)

	EASYSIZE(GLIM_B5,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(GLIM_B6,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(GLIM_B7,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(GLIM_B8,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
END_EASYSIZE_MAP

BOOL Dlg_Qly_image_nky::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	nClick=-1;
	f_Style_ListIMG();
	f_Load_ToolTip();
	f_Load_data();

	// Mặc định load hình ảnh đầu tiên
	int N = listIMG.GetItemCount();
	if(N>0)
	{
		_fullpath = listIMG.GetItemText(0,2);
		m_picCtrl.Load(_fullpath);

		CString val=L"";
		msg.Format(L"Hình ảnh (%s)",_fullpath);
		stt2.SetWindowText(msg);

		nClick=0;
		val = listIMG.GetItemText(0,3);
		txt1.SetWindowText(val);
	}

	INIT_EASYSIZE;

	return TRUE; 
}

// Bắt sự kiện click Enter
BOOL Dlg_Qly_image_nky::PreTranslateMessage(MSG* pMsg)
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

void Dlg_Qly_image_nky::OnSize(UINT nType, int cx, int cy) 
{
	CDialog::OnSize(nType, cx, cy);
	UPDATE_EASYSIZE;
}

void Dlg_Qly_image_nky::OnSizing(UINT fwSide, LPRECT pRect) 
{
	CDialog::OnSizing(fwSide, pRect);
	EASYSIZE_MINSIZE(338,360,fwSide,pRect);
}


void Dlg_Qly_image_nky::f_Load_ToolTip()
{
	CEdit * pbt0 = (CEdit *)GetDlgItem(GLIM_T1);

	m_ToolTips.Create(this);
	m_ToolTips.SetMaxTipWidth(600);
	m_ToolTips.SetDelayTime(900);

	m_ToolTips.AddTool(pbt0, L"Nhập nội dung mô tả chi tiết về hình ảnh.");
}


void Dlg_Qly_image_nky::f_Style_ListIMG()
{
	listIMG.InsertColumn(0,L"",LVCFMT_CENTER,0);
	listIMG.InsertColumn(1,L"Tệp tệp tin",LVCFMT_LEFT,250);
	listIMG.InsertColumn(2,L"Đường dẫn",LVCFMT_LEFT,0);
	listIMG.InsertColumn(3,L"Nội dung",LVCFMT_LEFT,0);
	listIMG.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);
}


void Dlg_Qly_image_nky::f_Load_data()
{
	TRY
	{
		_WorksheetPtr pSh1 = pWb->ActiveSheet;
		RangePtr pRg1 = pSh1->Cells;
		int ir = getRow(pSh1,"HK_NGAY");
		int ic = getColumn(pSh1,"HK_NGAY");
		sNgay = pRg1->GetItem(ir,ic);

		// Đường dẫn nhật ký
		if(getPathNHKY()==L"") return;
		if(connectDsn(pathNK)==-1) return;
		ObjConn.OpenAccessDB(pathNK,L"",L"");
		CRecordset* Recset;

		CString val0=L"",val1=L"",val2=L"";
		SqlString.Format(L"SELECT * FROM tbAnh WHERE ngayghi = '%s' ORDER BY maanh ASC;",sNgay);
		Recset = ObjConn.Execute(SqlString);
		int dem = 0;
		while( !Recset->IsEOF() )
		{
			Recset->GetFieldValue(L"hinhanh",val1);			
			val1.Trim();
			if(val1!=L"")
			{
				// Tách tên ảnh
				for (int i = val1.GetLength()-1; i > 0; i--)
				{
					val0 = val1.Mid(i,1);
					if(val0==L"\\")
					{
						val0 = val1.Right(val1.GetLength()-i-1);
						break;
					}
				}

				Recset->GetFieldValue(L"noidung",val2);
				listIMG.InsertItem(dem,L"",0);
				listIMG.SetItemText(dem,1,val0);
				listIMG.SetItemText(dem,2,val1);
				listIMG.SetItemText(dem,3,val2);

				dem++;
			}

			Recset->MoveNext();
		}
		Recset->Close();
		ObjConn.CloseDatabase();
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("Lỗi tải dữ liệu hình ảnh")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_Qly_image_nky::OnBnClickedB8()
{
	OnCancel();
}


void Dlg_Qly_image_nky::OnBnClickedB7()
{
	TRY
	{
		int N = listIMG.GetItemCount();
		if(N<=0) return;	

		// Kiểm tra sNgay đã tồn tại trong CSDL chưa?
		ObjConn.OpenAccessDB(pathNK, L"", L"");
		CRecordset* Recset;
		CString sTVan=L"",sDem=L"";

		sTVan.Format(L"SELECT COUNT(*) AS iDem FROM tbAnh WHERE ngayghi='%s';",sNgay);
		Recset = ObjConn.Execute(sTVan);
		Recset->GetFieldValue(L"iDem",sDem);
		Recset->Close();

		if(_wtoi(sDem)>0)
		{
			// Đã tồn tại --> Xóa để thêm mới
			sTVan.Format(L"DELETE FROM tbAnh WHERE ngayghi='%s';",sNgay);
			ObjConn.ExecuteRB(sTVan);
		}

		//////// THÊM MỚI DỮ LIỆU HÌNH ẢNH
		CString _spath=L"",_snoidung=L"";
		for (int i = 0; i <N; i++)
		{
			// Lưu dữ liệu
			_spath = listIMG.GetItemText(i,2);
			_snoidung = listIMG.GetItemText(i,3);

			if(i<9) sDem.Format(L"A0%d",i+1);
			else sDem.Format(L"A%d",i+1);			

			sTVan.Format(
				L"INSERT INTO tbAnh (ngayghi,maanh,hinhanh,noidung) VALUES ('%s','%s','%s','%s');"
				,sNgay,sDem,_spath,_snoidung);
			ObjConn.ExecuteRB(sTVan);
		}

		ObjConn.CloseDatabase();
		OnOK();
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("Lỗi cập nhật hình ảnh")+e->m_strError);
	}
	END_CATCH;
}


void Dlg_Qly_image_nky::OnNMClickL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	nClick = pNMListView->iItem;
	if(nClick==-1) return;

	CString p2 = listIMG.GetItemText(nClick,2);
	if(p2!=L"")
	{
		m_picCtrl.Load(p2);
		msg.Format(L"Hình ảnh (%s)",p2);
		stt2.SetWindowText(msg);	

		p2 = listIMG.GetItemText(nClick,3);
		txt1.SetWindowText(p2);
	}

	*pResult = 0;
}


void Dlg_Qly_image_nky::OnEnKillfocusT1()
{
	try
	{
		if(nClick==-1) return;

		CString p3 = L"";
		txt1.GetWindowTextW(p3);
		p3.Trim();

		listIMG.SetItemText(nClick,3,p3);
	}
	catch(...){}
}


void Dlg_Qly_image_nky::OnBnClickedB2()
{
	try
	{
		int N = listIMG.GetItemCount();
		if(N<=0) return;

		POSITION pos=listIMG.GetFirstSelectedItemPosition();
		if(pos!=NULL)
			for (int i = N-1; i >=0; i--)
				if (listIMG.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED) listIMG.DeleteItem(i);
	}
	catch(...){}
}


void Dlg_Qly_image_nky::OnBnClickedB1()
{
	try
	{
		UpdateData();

		CMultiFileOpenDialog dlgOpenFile(TRUE);
		dlgOpenFile.GetOFN().Flags |= OFN_ALLOWMULTISELECT;
		dlgOpenFile.SetUseExtBuffer(TRUE);		

		if(IDOK == dlgOpenFile.DoModal())
		{
			CString val=L"";
			int N = listIMG.GetItemCount();
			int dem=0;

			// get selected files
			POSITION pos = dlgOpenFile.GetStartPosition();
			while(NULL != pos)
			{
				CString strFilePath = dlgOpenFile.GetNextPathName(pos);
				val = strFilePath.Right(4);
				val.MakeLower();
				if(val==L".jpg")
				{
					for (int i = strFilePath.GetLength()-1; i > 0; i--)
					{
						val = strFilePath.Mid(i,1);
						if(val==L"\\")
						{
							val = strFilePath.Right(strFilePath.GetLength()-i-1);
							break;
						}
					}

					listIMG.InsertItem(N+dem,L"",0);
					listIMG.SetItemText(N+dem,1,val);
					listIMG.SetItemText(N+dem,2,strFilePath);
					dem++;
				}				
			}

			if(N<=0 && dem>0)
			{
				_fullpath = listIMG.GetItemText(0,2);
				m_picCtrl.Load(_fullpath);

				val=L"";
				msg.Format(L"Hình ảnh (%s)",_fullpath);
				stt2.SetWindowText(msg);

				nClick=0;
				val = listIMG.GetItemText(0,3);
				txt1.SetWindowText(val);
			}
		}
	}
	catch(...){}
}


void Dlg_Qly_image_nky::OnBnClickedB3()
{
	try
	{
		int N = listIMG.GetItemCount();
		if(N<=0) return;

		CString val=L"";
		POSITION pos=listIMG.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			int nId=1,nItem=0;
			for (int i = 0; i < N; i++)
				if (listIMG.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					nItem=i;
					break;
				}
			
			if(nItem>0)
			{
				for (int i = 1; i <= 3; i++)
				{
					val = listIMG.GetItemText(nItem,i);
					listIMG.SetItemText(nItem,i,listIMG.GetItemText(nItem-1,i));
					listIMG.SetItemText(nItem-1,i,val);
				}

				// Vị trí lựa chọn mới
				listIMG.SetItemState(nItem,0,LVIS_SELECTED);
				listIMG.SetItemState(nItem-1,LVIS_SELECTED,LVIS_SELECTED);
			}
		}
	}
	catch(...){}
}


void Dlg_Qly_image_nky::OnBnClickedB4()
{
	try
	{
		int N = listIMG.GetItemCount();
		if(N<=0) return;

		CString val=L"";
		POSITION pos=listIMG.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			int nId=1,nItem=0;
			for (int i = 0; i < N; i++)
				if (listIMG.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					nItem=i;
					break;
				}
			
			if(nItem<N-1)
			{
				for (int i = 1; i <= 3; i++)
				{
					val = listIMG.GetItemText(nItem,i);
					listIMG.SetItemText(nItem,i,listIMG.GetItemText(nItem+1,i));
					listIMG.SetItemText(nItem+1,i,val);
				}

				// Vị trí lựa chọn mới
				listIMG.SetItemState(nItem,0,LVIS_SELECTED);
				listIMG.SetItemState(nItem+1,LVIS_SELECTED,LVIS_SELECTED);
			}
		}
	}
	catch(...){}
}
