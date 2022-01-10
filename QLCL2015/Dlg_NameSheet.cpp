#include "Dlg_NameSheet.h"

Dlg_NameSheet::Dlg_NameSheet(CWnd* pParent /*=NULL*/)
	: CDialog(Dlg_NameSheet::IDD, pParent)
{
	
}

void Dlg_NameSheet::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, NMSH_L1, m_listName);
	DDX_Control(pDX, NMSH_S1, m_huongdan);
	DDX_Control(pDX, NMSH_E1, m_editName);
	DDX_Control(pDX, NMSH_B2, m_btnKhoiphuc);	
}

BEGIN_MESSAGE_MAP(Dlg_NameSheet, CDialog)
	ON_WM_SIZE()
	ON_WM_SIZING()
	ON_BN_CLICKED(NMSH_B2, &Dlg_NameSheet::OnBnClickedB2)
	ON_BN_CLICKED(NMSH_B1, &Dlg_NameSheet::OnBnClickedB1)
	ON_NOTIFY(NM_CLICK, NMSH_L1, &Dlg_NameSheet::OnNMClickL1)
	ON_EN_SETFOCUS(NMSH_E1, &Dlg_NameSheet::OnEnSetfocusE1)
	ON_EN_KILLFOCUS(NMSH_E1, &Dlg_NameSheet::OnEnKillfocusE1)
	ON_WM_SYSCOMMAND()
END_MESSAGE_MAP()

BEGIN_EASYSIZE_MAP(Dlg_NameSheet)
	EASYSIZE(NMSH_L1,ES_BORDER,ES_BORDER,ES_BORDER,ES_BORDER,0)
	EASYSIZE(NMSH_S1,ES_BORDER,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,0)	
	EASYSIZE(NMSH_B1,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(NMSH_B2,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
END_EASYSIZE_MAP

BOOL Dlg_NameSheet::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	sBefore=L"",sAfter=L"";
	m_huongdan.SubclassDlgItem(NMSH_S1,this);
	m_huongdan.SetTextColor(BLUE);
	
	smsg1=L"Lỗi Name";
	smsg2=L"Name đã xóa";
	smsg3=L"Nhập lại";

	CString szval=L"";
	szval.Format(
		L"Nếu cột Trạng thái thông báo '%s' hoặc '%s', bạn có thể đặt lại Name bằng cách "
		L"\nnhập lại vị trí ô đặt Name vào ô #REF hoặc '%s' tương ứng tại cột Vị trí và "
		L"kích vào nút 'Khôi phục Name'",smsg1,smsg2,smsg3);
	m_huongdan.SetWindowText(szval);

	pSheet = pWb->ActiveSheet;
	pRange = pSheet->Cells;
	sNSheet = (LPCTSTR)pSheet->Name;
	
	szval.Format(L"Danh sách Name tại sheet '%s'",sNSheet);
	this->SetWindowTextW(szval);

	_StyleList();
	_GetAllName();
	_GetAllNameSheet();

	INIT_EASYSIZE;

	return TRUE; 
}


BOOL Dlg_NameSheet::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		if(ClickEsc==0) OnBnClickedB1();
		else GotoDlgCtrl(GetDlgItem(NMSH_L1));
		ClickEsc=0;

		return TRUE; 
	}

	return FALSE;
}


void Dlg_NameSheet::OnSize(UINT nType, int cx, int cy) 
{
	CDialog::OnSize(nType, cx, cy);
	UPDATE_EASYSIZE;
}


void Dlg_NameSheet::OnSizing(UINT fwSide, LPRECT pRect) 
{
	CDialog::OnSizing(fwSide, pRect);
	EASYSIZE_MINSIZE(560,250,fwSide,pRect);	// chiều rộng + chiều cao
}


void Dlg_NameSheet::_StyleList()
{
	m_listName.InsertColumn(0,L"",LVCFMT_LEFT,0);
	m_listName.InsertColumn(1,L"Tên",LVCFMT_LEFT,150);
	m_listName.InsertColumn(2,L"Trạng thái",LVCFMT_CENTER,100);
	m_listName.InsertColumn(3,L"Vị trí",LVCFMT_CENTER,100);
	m_listName.InsertColumn(4,L"Mặc định",LVCFMT_CENTER,100);
	m_listName.InsertColumn(5,L"Mô tả",LVCFMT_LEFT,400);
	m_listName.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);
	m_listName.SetColumnEditor(5, &m_editName);
}


void Dlg_NameSheet::_GetAllName()
{
	try
	{
		int count = pWb->Names->Count;
		if(count==0) return;

		int _pos=-1;
		CString szval=L"",sref=L"";
		NamePtr _nameP;
		totalName=0;
		for (int i = 0; i < count; i++)
		{
			_nameP = pWb->Names->Item(i+1);		
			szval = (LPCTSTR)_nameP->GetName();
			szval.Trim();
			if(szval!=L"")
			{
				sref = _nameP->GetRefersTo();
				_pos = sref.Find(sNSheet);
				if(_pos>=0)
				{
					_pos = szval.Find(L"!");
					if(_pos>0) sName[totalName] = szval.Right(szval.GetLength()-_pos-1);
					else sName[totalName] = szval;

					sref.Replace(sNSheet,L"");
					sref.Replace(L"=",L"");
					sref.Replace(L"'",L"");
					sref.Replace(L"!",L"");
					sref.Replace(L"$",L"");
					sref.Replace(L" ",L"");
					sVtri[totalName] = sref;

					totalName++;
				}				
			}
		}
	}
	catch(...){}
}


void Dlg_NameSheet::_GetAllNameSheet()
{
	try
	{
		sPathCF.Format(L"%s%s%s",_fGPathFolder(L"ToolQLCL.dll"),_pathDlieu,_dFileConfig);
		if(connectDsn(sPathCF)==-1) return;
		ObjConn.OpenAccessDB(sPathCF, L"", L"");

		SqlString.Format(L"SELECT * FROM tblDSName WHERE Bangtinh LIKE '%s';",sNSheet);
		CRecordset* Recset = ObjConn.Execute(SqlString);

		int dem=0,icheck=0;
		CString szval=L"";
		while( !Recset->IsEOF() )
		{
			Recset->GetFieldValue(L"TenName",szval);		

			m_listName.InsertItem(dem,L"",0);
			m_listName.SetItemText(dem,1,szval);

			int ck = _CheckName(szval);
			if(ck==-1)
			{
				icheck=1;
				m_listName.SetItemText(dem,2,smsg2);
				m_listName.SetItemText(dem,3,smsg3);
				m_listName.SetRowColors(dem,WHITE,RED);
				m_listName.SetCellEditor(dem,3, &m_editName);
			}
			else
			{
				m_listName.SetItemText(dem,3,sVtri[ck]);
				if((int)sVtri[ck].Find(L"#")==-1)
				{
					m_listName.SetItemText(dem,2,L"Tốt");
					m_listName.SetRowColors(dem,WHITE,RGB(0,153,0));
				}
				else
				{
					icheck=1;
					m_listName.SetItemText(dem,2,smsg1);
					m_listName.SetRowColors(dem,WHITE,RED);
					m_listName.SetCellEditor(dem,3, &m_editName);
				}
			}

			Recset->GetFieldValue(L"Vitri",szval);
			m_listName.SetItemText(dem,4,szval);

			Recset->GetFieldValue(L"Mota",szval);
			m_listName.SetItemText(dem,5,szval);

			dem++;
			Recset->MoveNext();
		}
		Recset->Close();
		ObjConn.Close();

		if(icheck==1) m_btnKhoiphuc.EnableWindow(1);
	}
	catch(...){}
}


int Dlg_NameSheet::_CheckName(CString snm)
{
	if(totalName==0) return -1;
	for (int i = 0; i < totalName; i++) if(snm==sName[i]) return i;
	return -1;
}


void Dlg_NameSheet::OnBnClickedB2()
{
	try
	{
		CDialog::OnOK();
		CString szval=L"",sgt[5]={L"",L"",L"",L"",L""};
		int count = m_listName.GetItemCount();
		for (int i = count-1; i >= 0; i--)
		{
			sgt[2] = m_listName.GetItemText(i,2);	// trạng thái lỗi
			if(sgt[2]==smsg1 || sgt[2]==smsg2)
			{
				sgt[3] = m_listName.GetItemText(i,3);	// vị trí nhập
				sgt[3].Trim();
				if((int)sgt[3].Find(L"#")==-1 && sgt[3]!=smsg3 && sgt[3]!=L"")
				{
					sgt[3].MakeUpper();
					sgt[1] = m_listName.GetItemText(i,1);	// tên name

					for (int j = 0; j < sgt[3].GetLength(); j++)
					{
						sgt[0] = sgt[3].Mid(j,1);
						if(_wtoi(sgt[0])>0 || sgt[0]==L"0")
						{
							sgt[0].Format(L"$%s$%s",sgt[3].Left(j),sgt[3].Right(sgt[3].GetLength()-j));
							sgt[3] = sgt[0];
							break;
						}
					}

					szval.Format(L"='%s'!%s",sNSheet,sgt[3]);
					if(sgt[2]==smsg2) pWb->Names->Item((_bstr_t)sgt[1])->Delete();
					pSheet->Names->Add((_bstr_t)sgt[1],(_bstr_t)szval,true);
				}
			}
		}
	}
	catch(...){}
}


void Dlg_NameSheet::OnBnClickedB1()
{
	CDialog::OnCancel();
}


void Dlg_NameSheet::OnNMClickL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	ClickEsc=0;
	*pResult = 0;
}


void Dlg_NameSheet::OnEnSetfocusE1()
{
	ClickEsc=1;
}


void Dlg_NameSheet::OnEnKillfocusE1()
{
	try
	{
		sBefore = m_listName.GetItemText(CLRow,CLCol);
		sBefore.Replace(L"'",L"");

		m_editName.GetWindowTextW(sAfter);
		sAfter.Replace(L"'",L"");
		if(sAfter==sBefore) return;
		m_listName.SetItemText(CLRow,CLCol,sAfter);

		if(CLCol==5)
		{
			ObjConn.OpenAccessDB(sPathCF,L"",L"");
			CString sgt1 = m_listName.GetItemText(CLRow,1);
			SqlString.Format(L"UPDATE tblDSName SET Mota='%s' WHERE TenName LIKE '%s';",sAfter,sgt1);
			ObjConn.ExecuteRB(SqlString);
			ObjConn.CloseDatabase();
		}
	}
	catch(...){}
}


void Dlg_NameSheet::OnSysCommand(UINT nID, LPARAM lParam)
{
	if(nID == SC_CLOSE) OnBnClickedB1();
	else CDialog::OnSysCommand(nID, lParam);
}
