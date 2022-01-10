#include "stdafx.h"
#include "LaymauNgang.h"
#include "Dlg_ChangeLM3.h"
#include "Dlg_QCLMBetong.h"
#include "Dlg_Create_Tieude.h"
#include "frm_QuycachOld.h"

LaymauNgang::LaymauNgang(CWnd* pParent /*=NULL*/)
	: CDialog(LaymauNgang::IDD, pParent)
{

}

void LaymauNgang::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, LM3_L1, litLMN3);
	DDX_Control(pDX, LM3_E1, edLMN);
	DDX_Control(pDX, LM3_T1, m_txt1);
	DDX_Control(pDX, LM3_T2, m_txt2);
	DDX_Control(pDX, LM3_SP1, m_spin1);
	DDX_Control(pDX, LM3_B1, m_btn1);
	DDX_Control(pDX, LM3_CK1, m_chk1);
	DDX_Control(pDX, LM3_CK2, m_chk2);
	DDX_Control(pDX, LM3_CB1, m_cbb1);
	DDX_Control(pDX, LM3_B2, m_btn2);
	DDX_Control(pDX, LM3_B4, m_btn4);
	DDX_Control(pDX, LM3_T3, m_txt3);	
	DDX_Control(pDX, LM3_CK5, m_chk_dong);
	DDX_Control(pDX, LM3_CHKNT, m_chk_nthu);
}

BEGIN_MESSAGE_MAP(LaymauNgang, CDialog)	
	ON_WM_SIZE()
	ON_WM_SIZING()
	ON_WM_SYSCOMMAND()
	ON_WM_CONTEXTMENU()
	ON_BN_CLICKED(LM3_B4, &LaymauNgang::OnBnClickedB4)
	ON_BN_CLICKED(LM3_CK1, &LaymauNgang::OnBnClickedCk1)
	ON_BN_CLICKED(LM3_B1, &LaymauNgang::OnBnClickedB1)	
	ON_BN_CLICKED(LM3_B2, &LaymauNgang::OnBnClickedB2)
	ON_BN_CLICKED(LM3_B3, &LaymauNgang::OnBnClickedB3)
	ON_COMMAND(MNML_TEN, &LaymauNgang::OnMNML_Ten)
	ON_NOTIFY(NM_CLICK, LM3_L1, &LaymauNgang::OnNMClickL1)
	ON_NOTIFY(NM_RCLICK, LM3_L1, &LaymauNgang::OnNMRClickL1)
	ON_COMMAND(MNML_DCHUYEN, &LaymauNgang::OnThemdong)
	ON_CBN_SELCHANGE(LM3_CB1, &LaymauNgang::OnCbnSelchangeCb1)
	ON_BN_CLICKED(LM3_CK2, &LaymauNgang::OnBnClickedCk2)	
	ON_COMMAND(MNML_CHEN, &LaymauNgang::OnThemcot)
	ON_COMMAND(MNML_XOACOT, &LaymauNgang::OnXoacot)
	ON_COMMAND(MNML_XDONG, &LaymauNgang::OnXdong)
	ON_BN_CLICKED(LM3_B5, &LaymauNgang::OnBnClickedB5)
	ON_EN_SETFOCUS(LM3_E1, &LaymauNgang::OnEnSetfocusE1)
	ON_EN_SETFOCUS(LM3_T2, &LaymauNgang::OnEnSetfocusT2)
	ON_EN_SETFOCUS(LM3_T1, &LaymauNgang::OnEnSetfocusT1)
	ON_NOTIFY(UDN_DELTAPOS, LM3_SP1, &LaymauNgang::OnDeltaposSp1)
	ON_EN_KILLFOCUS(LM3_E1, &LaymauNgang::OnEnKillfocusE1)
	ON_EN_SETFOCUS(LM3_T3, &LaymauNgang::OnEnSetfocusT3)
	ON_EN_KILLFOCUS(LM3_T3, &LaymauNgang::OnEnKillfocusT3)
	ON_NOTIFY(NM_DBLCLK, LM3_L1, &LaymauNgang::OnNMDblclkL1)
	ON_COMMAND(MNML_MAU, &LaymauNgang::OnMau)
	ON_NOTIFY(LVN_KEYDOWN, LM3_L1, &LaymauNgang::OnLvnKeydownL1)
	ON_EN_KILLFOCUS(LM3_T1, &LaymauNgang::OnEnKillfocusT1)
	ON_BN_CLICKED(LM3_B8, &LaymauNgang::OnBnClickedB8)
	ON_EN_CHANGE(LM3_T3, &LaymauNgang::OnEnChangeT3)
END_MESSAGE_MAP()

BEGIN_EASYSIZE_MAP(LaymauNgang)
	EASYSIZE(LM3_S1, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(LM3_CB1, ES_BORDER, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(LM3_B5, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)

	EASYSIZE(LM3_CK1, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(LM3_T2, ES_BORDER, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)

	EASYSIZE(LM3_S3, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(LM3_T1, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(LM3_SP1, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(LM3_CK5, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(LM3_B1, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)

	EASYSIZE(LM3_LINE, ES_BORDER, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)

	EASYSIZE(LM3_S4, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(LM3_T3, ES_BORDER, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(LM3_CK2, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(LM3_B4, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)

	EASYSIZE(LM3_L1, ES_BORDER, ES_BORDER, ES_BORDER, ES_BORDER, 0)

	EASYSIZE(LM3_B8, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, 0)
	EASYSIZE(LM3_CHKNT, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, 0)

	EASYSIZE(LM3_B2, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(LM3_B3, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)	
END_EASYSIZE_MAP


BOOL LaymauNgang::OnInitDialog()
{
	CDialog::OnInitDialog();

	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);	

	f_Load_tooltip();

	m_cbb1.AdjustDroppedWidth();
	m_cbb1.SetMode(CComboBoxExt::MODE_AUTOCOMPLETE);
	m_txt2.SetCueBanner(L"Nhập tên mẫu mới...");

	edLMN.SubclassDlgItem(LM3_E1,this);
	edLMN.SetTextColor(BLACK);
	edLMN.SetBkColor(RGB(255,255,153));

	nItem=-1,nSubItem=-1;
	m_spin1.SetDecimalPlaces (0);
	m_spin1.SetTrimTrailingZeros (FALSE);
	m_spin1.SetRangeAndDelta (1,8,1);
	m_spin1.SetPos(5);

	// Load path
	if(getPathQCLM()==L"") return FALSE;
	
	int num = _wtoi(getVCell(psTS,L"CF_QCLMCV"));
	if(num!=1) num=0;
	m_chk2.SetCheck(num);
	m_txt3.SetCueBanner(L"Ví dụ: 4,6,10,15,... hoặc D,F,K,Q,Z,...");

	// Load CSDL
	f_Load_dulieu();
	_SetEnableNewMau(0);
	
	// Định nghĩa lấy mẫu
	keyLMTN = getVCell(psTS,L"CF_LMTN");
	keyLMTN.Trim();
	if(keyLMTN==L"") keyLMTN=L"LM";
	if((int)m_cbb1.GetCurSel()>=0) litLMN3.EnableWindow(1);

	INIT_EASYSIZE;

	// Set size dialog mfc
	if(irecW2>0) this->SetWindowPos(NULL,0,0,irecW2,irecH2,SWP_NOMOVE | SWP_NOZORDER);
	return TRUE;
}


BOOL LaymauNgang::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		if(ClickEsc==0) OnBnClickedB3();
		else GotoDlgCtrl(GetDlgItem(LM3_L1));
		ClickEsc=0;
		mnCphai=0;

		return TRUE; 
	}
	else if(pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(LM3_T2))
	{
		GotoDlgCtrl(GetDlgItem(LM3_T1));
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_TAB && pWndWithFocus == GetDlgItem(LM3_T1))
	{
		GotoDlgCtrl(GetDlgItem(LM3_T2));
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN && pWndWithFocus == GetDlgItem(LM3_T1))
	{
		OnBnClickedB1();
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


void LaymauNgang::OnSize(UINT nType, int cx, int cy) 
{
	CDialog::OnSize(nType, cx, cy);
	UPDATE_EASYSIZE;	
}


void LaymauNgang::OnSizing(UINT fwSide, LPRECT pRect) 
{
	CDialog::OnSizing(fwSide, pRect);
	EASYSIZE_MINSIZE(100,100,fwSide,pRect);	// chiều rộng + chiều cao
}


void LaymauNgang::f_Get_size()
{
	CRect rectCtrl;
	GetWindowRect(&rectCtrl);
	irecW2 = rectCtrl.Width();
	irecH2 = rectCtrl.Height();	
}


void LaymauNgang::f_Load_dulieu()
{
	TRY
	{
		if(connectDsn(pathQCLM)==-1) return;
		ObjConn.OpenAccessDB(pathQCLM, L"", L"");
		CRecordset* Recset;
		CString szval=L"",sGchu=L"";

		xl->PutScreenUpdating(true);
		pSheet = pWb->ActiveSheet;
		sCodeName = (LPCTSTR)pSheet->CodeName;
		pRange = pSheet->Cells;
		
		CButton* m_btn8 = (CButton*)GetDlgItem(LM3_B8);

		iC2=-1,iC4=1;
		iRow = pSheet->Application->ActiveCell->Row;
		if(sCodeName==L"shHSNTVL")
		{
			iC2 = getColumn(pSheet,L"DMVL_MHDG");
			iC4 = getColumn(pSheet,L"DMVL_ND");
			m_btn8->SetWindowText(L"Sử dụng mẫu vật liệu");
		}
		else
		{
			iC2 = getColumn(pSheet,L"DMBB_MH");
			iC4 = getColumn(pSheet,L"DMBB_ND");
			m_btn8->SetWindowText(L"Sử dụng mẫu bêtông");
		}

		CString sChitiet=L"";
		CString sadd=pRange->GetItem(iRow,iC4);
		sadd.Trim();

		// Xác định comment
		CString scmt=L"";
		PRS = pRange->GetItem(iRow,iC4);
		if(PRS->GetComment()!=NULL)
		{
			_bstr_t bscmt = PRS->GetComment()->Text();
			scmt = (LPCTSTR)(bscmt);
			scmt.Trim();
		}

		if(sadd!=L"" && scmt==L"LMN3")
		{
			// Tách tên mẫu
			int _pos = sadd.Find(L"|");
			if(_pos>0)
			{
				szval = sadd.Left(_pos);
				szval.Trim();
				sChitiet = sadd.Right(sadd.GetLength()-_pos-1);

				SqlString.Format(L"SELECT * FROM tbTenMau WHERE TenMau LIKE '%s';",szval);
				Recset = ObjConn.Execute(SqlString);
				while( !Recset->IsEOF() )
				{
					Recset->GetFieldValue(L"ID",sIDmau);
					sIDmau.Trim();
					break;
				}
				Recset->Close();
			}			
		}

/////////////////////////////////////////////////////////////

		// Xác định combobox dữ liệu			
		int dem=0,vtri=-1;
		SqlString = L"SELECT * FROM tbTenMau ORDER BY TenMau ASC;";
		Recset = ObjConn.Execute(SqlString);
		while( !Recset->IsEOF() )
		{
			Recset->GetFieldValue(L"ID",szval);
			if(sIDmau==szval)
			{
				vtri= dem;
				Recset->GetFieldValue(L"GChu",sGchu);
				sGchu.Trim();
			}

			Recset->GetFieldValue(L"TenMau",szval);
			m_cbb1.AddString(szval);

			dem++;
			Recset->MoveNext();
		}
		Recset->Close();
		ObjConn.CloseDatabase();
		m_cbb1.SetCueBanner(L"Kích chọn mẫu");

		iKeyX=0;
		if(vtri>=0)
		{
			m_cbb1.SetCurSel(vtri);

			if(sadd==L"")
			{
				OnCbnSelchangeCb1();
				return;
			}

			sDuoi=L"null";
			int pos = sGchu.Find(L"@");
			if(pos>=0)
			{
				sDuoi = sGchu.Right(sGchu.GetLength()-pos-1);
				sDuoi.Replace(L" ",L"");
				sDuoi.Replace(L";",L",");
				sGchu = sGchu.Left(pos);
			}
			
			if(sGchu!=L"")
			{
				sGchu+=L"|";
				
				vtri=0;
				CString val=L"";
				for (int i = 0; i < sGchu.GetLength(); i++)
				{
					val = sGchu.Mid(i,1);
					if(val==L"|" && i>vtri)
					{
						iKeyX++;
						pKeyX[iKeyX] = sGchu.Mid(vtri,i-vtri);
						vtri=i+1;

						if(iKeyX>=9) break;
					}
				}
			}
			else
			{
				iKeyX=5;
				pKeyX[1] = L"TT";
				pKeyX[2] = L"Tên chỉ tiêu";
				pKeyX[3] = L"Đơn vị";
				pKeyX[4] = L"Phương pháp TN";
				pKeyX[5] = L"Kết quả";
			}

			if(sDuoi!=L"null") m_txt3.SetWindowText(sDuoi);
			else m_txt3.SetWindowText(L"");

			// Load
			litLMN3.InsertColumn(0,L"Row-height (RH)",LVCFMT_CENTER,0);
			szval = pKeyX[1]; szval.MakeUpper();
			if(szval==L"TT") litLMN3.InsertColumn(1,L"TT",LVCFMT_CENTER,50);
			else if(szval==L"STT") litLMN3.InsertColumn(1,L"STT",LVCFMT_CENTER,60);
			else litLMN3.InsertColumn(1, pKeyX[1],LVCFMT_LEFT,60);

			if(iKeyX>=2)
			{
				for (int i = 2; i <= iKeyX; i++) litLMN3.InsertColumn(i, pKeyX[i],LVCFMT_LEFT,100);
			}			

			//litLMN3.SetColumnReadOnly(0);
			litLMN3.SetDefaultEditor(NULL, NULL, &edLMN);
			litLMN3.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);

			// Load dữ liệu từ excel hoặc từ CSDL
			int iKExl=0;
			CString sGtExl[1000];

			// Bổ sung 01.08.2018
			CString sDocao=L"";
			CString sFind = _rowheight;
			pos = sChitiet.Find(sFind);
			if(pos>0)
			{
				sDocao = sChitiet.Right(sChitiet.GetLength()-pos-sFind.GetLength());
				sChitiet = sChitiet.Left(pos);
			}

			if(sChitiet!=L"")
			{
				sChitiet+=L";";

				int vt=0;
				CString sel=L"";				
				for (int i = 0; i < sChitiet.GetLength(); i++)
				{
					sel = sChitiet.Mid(i,1);
					if(sel==L";" && i>vt)
					{
						iKExl++;
						sGtExl[iKExl] = sChitiet.Mid(vt,i-vt);
						vt=i+1;
					}
				}
			}

			if(iKExl<iKeyX)
			{
				for (int i = iKExl+1; i <= iKeyX; i++) sGtExl[i]=L"";
				iKExl=iKeyX;
			}						
			int idong = (int)(iKExl/iKeyX);

			dem=0;
			for (int i = 0; i < idong; i++)
			{
				litLMN3.InsertItem(i,L"",0);
				for (int j = 1; j <= iKeyX; j++)
				{
					dem++;
					if(dem<=iKExl) litLMN3.SetItemText(i,j,sGtExl[dem]);
				}
			}

			// Bổ sung 01.08.2018
			if(sDocao!=L"")
			{
				for (int i = 0; i < litLMN3.GetItemCount(); i++)
				{
					pos = sDocao.Find(L";");
					if(pos>=0)
					{
						litLMN3.SetItemText(i,0,sDocao.Left(pos));
						if(pos==sDocao.GetLength()-1) return;
						sDocao = sDocao.Right(sDocao.GetLength()-pos-1);
					}
				}
			}
		}		
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("Lỗi xác định cơ sở dữ liệu.")+e->m_strError);
	}
	END_CATCH;
}


// Cập nhật dữ liệu
void LaymauNgang::f_Luu_mau_ngang(CString sn)
{
	try
	{
		int ncount = litLMN3.GetColumnCount();
		if(ncount<=1) return;

		int nc = m_cbb1.GetCurSel();
		if(nc==-1)
		{
			_s(L"Bạn chưa chọn mẫu lưu. Vui lòng kiểm tra lại.");
			return;
		}

		CString sTen=L"";
		m_cbb1.GetLBText(nc,sTen);
		sTen.Trim();

		CString szval=L"";
		szval.Format(L"Mẫu '%s' đã tồn tại. Bạn có muốn cập nhật?",sTen);
		if(_yn(szval,0)!=6) return;

		ObjConn.OpenAccessDB(pathQCLM, L"", L"");

		CString skt = _NameColumnHeading(0,1,L"");
		if(ncount>2)
		for (int i = 2; i < ncount; i++)
		{
			skt+=L"|";
			skt+=_NameColumnHeading(0,i,L"");
		}

		skt+=L"@";
		if(sn!=L"") skt+=sn;
		else skt+=L"null";

		// Leo bổ sung lưu thành phần chi tiết
		CString sChitiet=L"";
		int tlCol = litLMN3.GetColumnCount();
		int tlRow = litLMN3.GetItemCount();

		if(tlRow>0)
		for (int i = 0; i < tlRow; i++)
			for (int j = 1; j < tlCol; j++)
			{
				szval = litLMN3.GetItemText(i,j);
				if(szval==L"") szval=L" ";
				sChitiet+=szval;
				sChitiet+=L";";
			}

		// Bổ sung độ cao dòng 01.08.2018
		int rw=0;
		CString sDocao=L"";
		for (int i = 0; i < tlRow; i++)
		{
			szval = litLMN3.GetItemText(i,0);
			if(szval!=L"")
			{
				rw=1;
				break;
			}
		}

		if(rw==1)
		{
			for (int i = 0; i < tlRow; i++)
			{
				szval = litLMN3.GetItemText(i,0);
				rw = _wtoi(szval);
				if(rw<=0 || rw>409) szval=L"0";

				sDocao+=szval;
				sDocao+=L";";
			}
		}		

		if(sDocao!=L"")
		{
			sChitiet+=_rowheight;
			sChitiet+=sDocao;
		}

		SqlString.Format(
			L"UPDATE tbTenMau SET GChu='%s',Chitietmau='%s' WHERE ID='%s';",skt,sChitiet,sIDmau);
		ObjConn.ExecuteRB(SqlString);
		ObjConn.CloseDatabase();
	}
	catch(...){}
}


void LaymauNgang::OnBnClickedB4()
{
	TRY
	{
		ClickEsc=0;

		int ncount = litLMN3.GetColumnCount();
		if(ncount<=1) return;
		
		// Kiểm tra cột mẫu có chuẩn không
		CString snum=L"",szval=L"";
		m_txt3.GetWindowTextW(snum);
		snum.Trim();
		snum.Replace(L" ",L"");
		snum.Replace(L";",L",");
		snum.MakeLower();
		int iLen = snum.GetLength();
		if(snum!=L"")
		{
			int dem=0;				
			for (int i = 0; i < iLen; i++)
			{
				szval = snum.Mid(i,1);
				if(szval==L",") dem++;
			}

			if(dem!=ncount-2 || iLen<= 2*dem)
			{
				szval.Format(L"Vui lòng nhập đúng số cột sẽ đổ dữ liệu vào (%d cột)",ncount-1);
				_s(szval);
				GotoDlgCtrl(GetDlgItem(LM3_T3));

				return;
			}
		}

		// Kiểm tra có phải là cập nhật lại dữ liệu không?
		if((int)m_chk1.GetCheck()==0)
		{
			f_Luu_mau_ngang(snum);
			return;
		}

		CString stenmau=L"";
		m_txt2.GetWindowTextW(stenmau);
		stenmau.Replace(L"'",L"");
		stenmau.Trim();
		if(stenmau==L"")
		{
			_s(L"Vui lòng nhập tên mẫu");
			GotoDlgCtrl(GetDlgItem(LM3_T2));
			return;
		}

		// Kiểm tra tên mẫu có trùng không?
		CString skt=L"";
		int nc = m_cbb1.GetCount();
		if(nc>0)
		{
			for (int i = 0; i < nc; i++)
			{
				m_cbb1.GetLBText(i,skt);
				if(skt==stenmau)
				{
					_s(L"Tên mẫu đã trùng, vui lòng kiểm tra lại!");
					GotoDlgCtrl(GetDlgItem(LM3_T2));
					return;
				}
			}
		}

		m_chk1.SetCheck(0);
		_SetEnableNewMau(0);

		m_btn2.EnableWindow(1);
		m_cbb1.EnableWindow(1);	
		m_cbb1.AddString(stenmau);

		nc = m_cbb1.GetCount();
		for (int i = 0; i < nc; i++)
		{
			m_cbb1.GetLBText(i,skt);
			if(skt==stenmau)
			{
				m_cbb1.SetCurSel(i);
				break;
			}
		}

		// Lưu vào CSDL
		ObjConn.OpenAccessDB(pathQCLM, L"", L"");
		CRecordset* Recset;

		CString sID=L"N01",sdem=L"";
		for (int i = 1; i <= 99; i++)
		{
			if(i>9) sID.Format(L"N%d",i);
			else sID.Format(L"N0%d",i);

			SqlString.Format(L"SELECT COUNT(*) AS iDem FROM tbTenMau WHERE ID='%s';",sID);
			Recset = ObjConn.Execute(SqlString);
			Recset->GetFieldValue(L"iDem",sdem);
			Recset->Close();

			if(_wtoi(sdem)==0) break;
		}

		skt = _NameColumnHeading(0,1,L"");
		if(ncount>2)
		for (int i = 2; i < ncount; i++)
		{
			skt+=L"|";
			skt+=_NameColumnHeading(0,i,L"");
		}

		skt+=L"@";
		if(snum!=L"") skt+=snum;
		else skt+=L"null";

		// Leo bổ sung lưu thành phần chi tiết
		CString sChitiet=L"";
		int tlCol = litLMN3.GetColumnCount();
		int tlRow = litLMN3.GetItemCount();

		if(tlRow>0)
		for (int i = 0; i < tlRow; i++)
			for (int j = 1; j < tlCol; j++)
			{
				szval = litLMN3.GetItemText(i,j);
				if(szval==L"") szval=L" ";
				sChitiet+=szval;
				sChitiet+=L";";
			}

		// Bổ sung độ cao dòng 01.08.2018
		int rw=0;
		CString sDocao=L"";
		for (int i = 0; i < tlRow; i++)
		{
			szval = litLMN3.GetItemText(i,0);
			if(szval!=L"")
			{
				rw=1;
				break;
			}
		}

		if(rw==1)
		{
			for (int i = 0; i < tlRow; i++)
			{
				szval = litLMN3.GetItemText(i,0);
				rw = _wtoi(szval);
				if(rw<=0 || rw>409) szval=L"0";
				
				sDocao+=szval;
				sDocao+=L";";
			}
		}		
		
		if(sDocao!=L"")
		{
			sChitiet+=_rowheight;
			sChitiet+=sDocao;
		}

		SqlString.Format(L"INSERT INTO tbTenMau (ID,TenMau,GChu,Chitietmau) "
			L"VALUES ('%s','%s','%s','%s');",sID,stenmau,skt,sChitiet);
		ObjConn.ExecuteRB(SqlString);
		ObjConn.CloseDatabase();
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("Lỗi lưu dữ liệu.")+e->m_strError);
	}
	END_CATCH;
}


void LaymauNgang::OnBnClickedCk1()
{
	ClickEsc=0;
	mnCphai=1;
	int nc = m_chk1.GetCheck();
	_SetEnableNewMau(nc);

	if(nc==1)
	{
		m_btn2.EnableWindow(0);
		m_cbb1.EnableWindow(0);
		GotoDlgCtrl(GetDlgItem(LM3_T2));
	}
	else
	{
		m_btn2.EnableWindow(1);
		m_cbb1.EnableWindow(1);

		//...
	}
}


void LaymauNgang::_SetEnableNewMau(int itype)
{
	m_txt2.EnableWindow(itype);
	m_spin1.EnableWindow(itype);
	m_txt1.EnableWindow(itype);
	m_btn1.EnableWindow(itype);
	m_chk_dong.EnableWindow(itype);
}


void LaymauNgang::_StyleListLM3(CString sTenmau)
{
	TRY
	{
		if(connectDsn(pathQCLM)==-1) return;
		ObjConn.OpenAccessDB(pathQCLM, L"", L"");

		// Xác định combobox dữ liệu
		CString szval=L"",sGchu=L"",sChitiet=L"";
		SqlString.Format(L"SELECT * FROM tbTenMau WHERE TenMau LIKE '%s';",sTenmau);
		CRecordset* Recset = ObjConn.Execute(SqlString);
		while( !Recset->IsEOF() )
		{
			Recset->GetFieldValue(L"ID",sIDmau);
			Recset->GetFieldValue(L"GChu",sGchu);
			Recset->GetFieldValue(L"Chitietmau",sChitiet);
			break;
		}
		Recset->Close();
		ObjConn.CloseDatabase();

		sDuoi=L"null";
		int pos = sGchu.Find(L"@");
		if(pos>=0)
		{
			sDuoi = sGchu.Right(sGchu.GetLength()-pos-1);
			sDuoi.Replace(L" ",L"");
			sDuoi.Replace(L";",L",");

			sGchu = sGchu.Left(pos);
		}

		iKeyX=0;
		if(sGchu!=L"")
		{
			sGchu+=L"|";
			int vtri=0;
			CString val=L"";
			for (int i = 0; i < sGchu.GetLength(); i++)
			{
				val = sGchu.Mid(i,1);
				if(val==L"|" && i>vtri)
				{
					iKeyX++;
					pKeyX[iKeyX] = sGchu.Mid(vtri,i-vtri);
					vtri=i+1;

					if(iKeyX==9) break;
				}
			}
		}
		else
		{
			iKeyX=5;
			pKeyX[1] = L"TT";
			pKeyX[2] = L"Tên chỉ tiêu";
			pKeyX[3] = L"Đơn vị";
			pKeyX[4] = L"Phương pháp thí nghiệm";
			pKeyX[5] = L"Kết quả";
		}

		if(sDuoi!=L"null") m_txt3.SetWindowText(sDuoi);
		else m_txt3.SetWindowText(L"");

		// Load
		litLMN3.InsertColumn(0,L"Row-height (RH)",LVCFMT_CENTER,0);
		szval = pKeyX[1];
		szval.MakeUpper();
		if(szval==L"TT") litLMN3.InsertColumn(1,L"TT",LVCFMT_CENTER,50);
		else if(szval==L"STT") litLMN3.InsertColumn(1,L"STT",LVCFMT_CENTER,60);
		else litLMN3.InsertColumn(1,pKeyX[1],LVCFMT_LEFT,80);

		if(iKeyX>=2)
		for (int i = 2; i <= iKeyX; i++)
			litLMN3.InsertColumn(i,pKeyX[i],LVCFMT_LEFT,100);

		//litLMN3.SetColumnReadOnly(0);
		litLMN3.SetDefaultEditor(NULL, NULL, &edLMN);
		litLMN3.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);

		// Leo đổ chi tiết dữ liệu
		if(sChitiet==L"") return;

		// Bổ sung 01.08.2018
		CString sDocao=L"";
		CString sFind = _rowheight;
		pos = sChitiet.Find(sFind);
		if(pos>0)
		{
			sDocao = sChitiet.Right(sChitiet.GetLength()-pos-sFind.GetLength());
			sChitiet = sChitiet.Left(pos);
		}

		int tlCol = litLMN3.GetColumnCount();
		for (int i = 0; i < 100; i++)
		{
			litLMN3.InsertItem(i,L"",0);
			for (int j = 1; j < tlCol; j++)
			{
				pos = sChitiet.Find(L";");
				if(pos>=0)
				{
					litLMN3.SetItemText(i,j,sChitiet.Left(pos));
					if(pos==sChitiet.GetLength()-1)
					{
						// Bổ sung 01.08.2018
						if(sDocao!=L"")
						{
							for (int k = 0; k <= i; k++)
							{
								pos = sDocao.Find(L";");
								if(pos>=0)
								{
									litLMN3.SetItemText(k,0,sDocao.Left(pos));
									if(pos==sDocao.GetLength()-1) return;
									sDocao = sDocao.Right(sDocao.GetLength()-pos-1);
								}
							}
						}							

						return;
					}

					sChitiet = sChitiet.Right(sChitiet.GetLength()-pos-1);
				}
			}
		}		
	}
	CATCH(CDBException, e)
	{
		AfxMessageBox(_T("Lỗi xác định cơ sở dữ liệu.")+e->m_strError);
	}
	END_CATCH;
}

void LaymauNgang::OnBnClickedB1()
{
	// Xác định số cột sử dụng
	CString szval = L"";
	m_txt1.GetWindowTextW(szval);
	szval.Trim();
	int sl = _wtoi(szval);
	if (sl <= 1) sl = 1;
	else if (sl > 8) sl = 8;
	sl++;

	int sldong = 0;
	CString szTitle = L"";
	if ((int)m_chk_dong.GetCheck() == 1)
	{
		AFX_MANAGE_STATE(AfxGetStaticModuleState());
		Dlg_Create_Tieude open_dlg;
		open_dlg.sl = sl;
		open_dlg.DoModal();

		sldong = open_dlg.sldong;
		szTitle = open_dlg.szTitle;
	}

	if (sldong == -1) return;

	ClickEsc = 0;
	litLMN3.EnableWindow(1);
	_ClearAllList();

	litLMN3.InsertColumn(0, L"Row-height (RH)", LVCFMT_LEFT, 0);

	if (szTitle != L"")
	{
		_TackTukhoaChuan(szTitle, L"|", L"|", 0, 0);

		if (sl >= 1)
		{
			for (int i = 1; i < sl; i++)
			{
				litLMN3.InsertColumn(i, pKey[i], LVCFMT_LEFT, 100);
			}
		}
	}
	else
	{
		litLMN3.InsertColumn(1, L"TT", LVCFMT_CENTER, 80);

		if (sl >= 2)
		{
			for (int i = 2; i < sl; i++)
			{
				szval.Format(L"Cột %d", i);
				litLMN3.InsertColumn(i, szval, LVCFMT_LEFT, 100);
			}
		}		
	}

	litLMN3.InsertItem(0, L"", 0);
	litLMN3.SetItemText(0, 1, L"1");

	if (sl >= 2)
	{
		for (int i = 2; i < sl; i++)
		{
			litLMN3.SetItemText(0, i, L"[giá trị]");
		}
	}

	//litLMN3.SetColumnReadOnly(0);
	litLMN3.SetDefaultEditor(NULL, NULL, &edLMN);
	litLMN3.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);

	szval=L"";

	switch (sl)
	{
		case 2:
			szval = L"C";
			break;

		case 3:
			szval = L"C,D";
			break;

		case 4:
			szval = L"C,D,T";
			break;

		case 5:
			szval = L"C,D,T,Y";
			break;

		case 6:
			szval = L"C,D,O,T,Y";
			break;

		case 7:
			szval = L"C,D,M,R,V,Z";
			break;

		case 8:
			szval = L"C,D,L,O,R,V,Z";
			break;

		case 9:
			szval = L"C,D,L,O,R,S,V,AA";
			break;

		default:
			break;
	}

	m_txt3.SetWindowText(szval);
}


void LaymauNgang::_ClearAllList()
{
	if(litLMN3.GetColumnCount()>0)
	{
		litLMN3.DeleteAllColumns();
		ASSERT(litLMN3.GetColumnCount() == 0);
	}
	
	if(litLMN3.GetItemCount()>0)
	{
		litLMN3.DeleteAllItems();
		ASSERT(litLMN3.GetItemCount() == 0);
	}	
}


void LaymauNgang::OnContextMenu(CWnd* /*pWnd*/, CPoint point)
{
	try
	{
		if(mnCphai==0) return;

		CMenu _menu;
		CMenu *_pMenu;
		CListCtrlEx *pClist;
		pClist = reinterpret_cast<CListCtrlEx *>(GetDlgItem(LM3_L1));
		CRect rectSubmitButton;
		pClist->GetWindowRect(&rectSubmitButton);

		_menu.LoadMenu(IDR_MENU3);
		_pMenu = _menu.GetSubMenu(0);

		if(nItem==-1 || nSubItem==-1)
		{
			_pMenu->EnableMenuItem(MNML_TEN, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			_pMenu->EnableMenuItem(MNML_CHEN, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			_pMenu->EnableMenuItem(MNML_XOACOT, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			_pMenu->EnableMenuItem(MNML_XDONG, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
		}

		if((int)m_chk1.GetCheck()==1) _pMenu->EnableMenuItem(MNML_MAU, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);

		ASSERT(_pMenu);
			if( rectSubmitButton.PtInRect(point) )
				_pMenu->TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON, point.x, point.y, this);
	}
	catch(...){}
}


void LaymauNgang::OnBnClickedB2()
{
	try
	{
		ClickEsc=0;
		int ncount = litLMN3.GetItemCount();
		int sl = litLMN3.GetColumnCount();
		if(ncount==0 || sl<=1)
		{
			_s(L"Vui lòng chọn mẫu sử dụng.");
			return;
		}

		// Kiểm tra cột mẫu có chuẩn không
		CString snum=L"",szval=L"";
		m_txt3.GetWindowTextW(snum);
		snum.Trim();
		snum.Replace(L" ",L"");
		snum.Replace(L";",L",");
		int iLen = snum.GetLength();
		if(snum!=L"")
		{
			int dem=0;				
			for (int i = 0; i < iLen; i++)
			{
				szval = snum.Mid(i,1);
				if(szval==L",") dem++;
			}

			if(dem!=sl-2 || iLen<= 2*dem)
			{
				szval.Format(L"Vui lòng nhập đúng số cột sẽ đổ dữ liệu vào (%d cột)",sl-1);
				_s(szval);
				GotoDlgCtrl(GetDlgItem(LM3_T3));
				return;
			}
		}

		CString sgan=L"";
		m_cbb1.GetLBText(m_cbb1.GetCurSel(),sgan);

		szval=L"";
		sgan+=L" |";
		szval+=sgan;

		// Đổ dữ liệu
		for (int i = 0; i < ncount; i++)
		{
			for (int j = 1; j < sl; j++)
			{
				sgan = litLMN3.GetItemText(i,j);
				if(sgan==L"") sgan=L" ";
				szval+=sgan;
				szval+=L";";
			}
		}

		// Bổ sung độ cao dòng 01.08.2018
		int rw=0;
		CString sDocao=L"";
		for (int i = 0; i < ncount; i++)
		{
			sgan = litLMN3.GetItemText(i,0);
			if(sgan!=L"")
			{
				rw=1;
				break;
			}
		}

		if(rw==1)
		{
			for (int i = 0; i < ncount; i++)
			{
				sgan = litLMN3.GetItemText(i,0);
				rw = _wtoi(sgan);
				if(rw<=0 || rw>409) sgan=L"0";
				
				sDocao+=sgan;
				sDocao+=L";";
			}
		}		
		
		if(sDocao!=L"")
		{
			szval+=_rowheight;
			szval+=sDocao;
		}

		mnCphai = 0;
		f_Get_size();
		CDialog::OnOK();

		xl->PutScreenUpdating(VARIANT_FALSE);		

		// Set quy cách mặc định
		if ((int)m_chk2.GetCheck() == 1) putVal(psTS, L"CF_QCLMCV", L"1");
		else putVal(psTS, L"CF_QCLMCV", L"0");

		pSheet = pWb->ActiveSheet;
		CString sSheet = (LPCTSTR)pSheet->CodeName;
		if (sSheet.Left(8) == L"shHSNTVL")
		{
			if (sSheet.GetLength() > 8)
			{
				pSheet = pWb->Worksheets->GetItem(name_sheet((_bstr_t)"shHSNTVL"));
				pRange = pSheet->Cells;
				pSheet->Activate();
			}
		}
		else
		{
			if (sSheet != L"shHSNTCV")
			{
				pSheet = pWb->Worksheets->GetItem(name_sheet((_bstr_t)"shHSNTCV"));
				pRange = pSheet->Cells;
				pSheet->Activate();
			}
		}

		CString szLM = keyLMTN;
		if ((int)m_chk_nthu.GetCheck() == 1) szLM = L"NT";
		pRange->PutItem(iRow,iC2,(_bstr_t)szLM);
		pRange->PutItem(iRow,iC4,(_bstr_t)szval);

		PRS = pRange->GetItem(iRow,iC4);
		PRS->Interior->PutColor(xlNone);
		PRS->Font->PutColor(RGB(0,0,0));
		PRS->Font->PutItalic(1);
		PRS->Font->PutBold(0);
		PRS->EntireRow->Rows->AutoFit();
		if((int)PRS->GetRowHeight()>80) PRS->PutWrapText(FALSE);

		if(PRS->GetComment()!=NULL) PRS->ClearComments();
		PRS->AddComment((_bstr_t)L"LMN3");

		xl->PutScreenUpdating(VARIANT_TRUE);		
	}
	catch(...){}	
}


void LaymauNgang::OnBnClickedB3()
{
	mnCphai=0;
	f_Get_size();

	// Set quy cách mặc định
	if((int)m_chk2.GetCheck()==1) putVal(psTS,L"CF_QCLMCV",L"1");
	else putVal(psTS,L"CF_QCLMCV",L"0");

	CDialog::OnCancel();
}


void LaymauNgang::OnMNML_Ten()
{
	iKqua=nSubItem;	

	AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
	Dlg_ChangeLM3 _dlg;
	_dlg.DoModal();
	
	xl_tukhoa=L"";
	iKqua=0;
}


void LaymauNgang::OnNMClickL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	ClickEsc=0;
	mnCphai=1;
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	nItem = pNMListView->iItem;
	nSubItem = pNMListView->iSubItem;
	*pResult = 0;	
}


void LaymauNgang::OnNMRClickL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	ClickEsc=0;
	mnCphai=1;
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	nItem = pNMListView->iItem;
	nSubItem = pNMListView->iSubItem;
	*pResult = 0;
}


void LaymauNgang::OnThemdong()
{
	int nc = litLMN3.GetItemCount();
	
	CString s=L"";
	s.Format(L"%d",nc+1);
	litLMN3.InsertItem(nc,L"",0);
	litLMN3.SetItemText(nc,1,s);

	int sl = litLMN3.GetColumnCount();
	if(sl>=2)
	for (int i = 2; i < sl; i++) litLMN3.SetItemText(nc,i,L"[giá trị]");
}

void LaymauNgang::OnCbnSelchangeCb1()
{
	ClickEsc=0;
	mnCphai=1;
	CString szval=L"";
	int ns = m_cbb1.GetCurSel();	
	if(ns>=0)
	{
		m_cbb1.GetLBText(ns,szval);
		litLMN3.EnableWindow(1);
	}

	_ClearAllList();
	_StyleListLM3(szval);
}


void LaymauNgang::OnBnClickedCk2()
{
	ClickEsc=0;
	mnCphai=1;
}


void LaymauNgang::OnSysCommand(UINT nID, LPARAM lParam)
{
	if(nID == SC_CLOSE) OnBnClickedB3();
	else CDialog::OnSysCommand(nID, lParam);
}


void LaymauNgang::_AutoColumnName()
{
	int ncount = litLMN3.GetColumnCount();
	if(ncount==0 || ncount>9) return;

	CString szval=L"";
	for (int i = 1; i < ncount; i++)
	{
		szval = _NameColumnHeading(0,i,L"");
		if(szval.Left(8)==L"[column-")
		{
			szval.Format(L"[column-%d]",i);
			_NameColumnHeading(1,i,szval);
		}
	}
}

void LaymauNgang::OnThemcot()
{
	int ncount = litLMN3.GetColumnCount();
	if(ncount>=9)
	{
		_s(L"Bạn chỉ có thể tạo tối đa 8 cột.");
		return;
	}
	else if(ncount==0)
	{
		litLMN3.InsertColumn(0,L"Row-height (RH)",LVCFMT_CENTER,0);
		litLMN3.InsertColumn(1,L"TT",LVCFMT_CENTER,60);
		return;
	}

	int vt=nSubItem;
	if(vt==-1) vt = ncount;
	litLMN3.InsertColumn(vt,L"[column-?]",LVCFMT_LEFT,100);

	// Đánh lại tên cột
	_AutoColumnName();
}


void LaymauNgang::OnXoacot()
{
	if(nSubItem<=0) return;

	CString szval=L"";
	szval.Format(L"Bạn có chắc chắn xóa cột '%s'?",_NameColumnHeading(0,nSubItem,L""));
	if(_yn(szval,0)!=6) return;
	litLMN3.DeleteColumn(nSubItem);

	// Đánh lại tên cột
	_AutoColumnName();
	m_btn4.EnableWindow(1);
}


void LaymauNgang::OnXdong()
{
	int ncount = litLMN3.GetItemCount();
	if(ncount==0) return;

	POSITION pos=litLMN3.GetFirstSelectedItemPosition();
	if(pos==NULL) return;

	for (int i = ncount-1; i >= 0; i--)
		if (litLMN3.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED) litLMN3.DeleteItem(i);
}


void LaymauNgang::OnBnClickedB5()
{
	ClickEsc=0;
	int nc = m_cbb1.GetCurSel();
	if(nc==-1) return;

	CString sTen=L"";
	m_cbb1.GetLBText(nc,sTen);
	sTen.Trim();

	CString szval=L"";
	szval.Format(L"Bạn có chắc chắn xóa mẫu '%s'?",sTen);
	if(_yn(szval,0)!=6) return;

	ObjConn.OpenAccessDB(pathQCLM, L"", L"");
	SqlString.Format(L"DELETE FROM tbTenMau WHERE TenMau = '%s';",sTen);
	ObjConn.ExecuteRB(SqlString);
	ObjConn.CloseDatabase();
	m_cbb1.DeleteString(nc);
	m_cbb1.SetCurSel(-1);	

	_ClearAllList();
}


void LaymauNgang::OnBnClickedCk3()
{
	ClickEsc=0;
	mnCphai=1;
}


void LaymauNgang::OnBnClickedCk4()
{
	ClickEsc=0;
	mnCphai=1;
	m_txt3.EnableWindow(1);
}


void LaymauNgang::OnEnSetfocusE1()
{
	ClickEsc=1;
	mnCphai=1;
}


void LaymauNgang::OnEnSetfocusT2()
{
	ClickEsc=0;
	mnCphai=1;
}


void LaymauNgang::OnEnSetfocusT1()
{
	ClickEsc=0;
	mnCphai=1;
}


void LaymauNgang::OnDeltaposSp1(NMHDR *pNMHDR, LRESULT *pResult)
{
	ClickEsc=0;
	mnCphai=1;
	*pResult = 0;
}


void LaymauNgang::OnEnKillfocusE1()
{
	if(CLRow==-1 || CLCol==-1) return;

	// kiểm tra dữ liệu nhập vào có trùng không?
	CString ktra=L"";
	edLMN.GetWindowTextW(ktra);
	ktra.Replace(L"'",L"");
	ktra.Replace(L"+-",L"±");
	ktra.Replace(L"//",L"÷");

	CString val = listQclm.GetItemText(CLRow,CLCol);
	if(ktra==val) return;
	edLMN.SetWindowText(ktra);
}


void LaymauNgang::f_Load_tooltip()
{
	CListCtrlEx *lst1 = (CListCtrlEx *)GetDlgItem(LM3_L1); 
	m_ToolTips.Create(this);
	m_ToolTips.SetMaxTipWidth(600);
	m_ToolTips.SetDelayTime(900);

	m_ToolTips.AddTool(lst1, L"Kích đúp vào dưới dữ liệu cuối để thêm mới. \nKích chuột phải để hiển thị menu trợ giúp.");
}

void LaymauNgang::OnEnSetfocusT3()
{
	ClickEsc=0;
	mnCphai=1;
}

void LaymauNgang::OnEnKillfocusT3()
{
	
}

void LaymauNgang::OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	ClickEsc=0;
	mnCphai=1;
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	nItem = pNMListView->iItem;
	nSubItem = pNMListView->iSubItem;

	if(nItem==-1 || nSubItem==-1) OnThemdong();

	*pResult = 0;
}

void LaymauNgang::OnMau()
{
	try
	{
		iKqua = -1;
		m_cbb1.GetLBText(m_cbb1.GetCurSel(), xl_tukhoa);

		AFX_MANAGE_STATE(AfxGetStaticModuleState());
		Dlg_ChangeLM3 _dlg;
		_dlg.DoModal();

		if (iKqua == -2)
		{
			int nc = m_cbb1.GetCurSel();
			if (nc == -1) return;

			CString szval = L"";
			vector<CString> vcbb;
			iKqua = m_cbb1.GetCount();
			for (int i = 0; i < iKqua; i++)
			{
				m_cbb1.GetLBText(i, szval);
				if (i == nc) szval = xl_tukhoa;
				vcbb.push_back(szval);
			}

			m_cbb1.ResetContent();
			ASSERT(m_cbb1.GetCount() == 0);

			int nsize = (int)vcbb.size();
			for (int i = 0; i < nsize; i++) m_cbb1.AddString(vcbb[i]);

			for (int i = 0; i < nsize; i++)
			{
				m_cbb1.GetLBText(i, szval);
				if (xl_tukhoa == szval)
				{
					m_cbb1.SetCurSel(i);
					break;
				}
			}
		}

		xl_tukhoa = L"";
		iKqua = 0;
	}
	catch (...) {}
}

void LaymauNgang::OnLvnKeydownL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);
	if(pLVKeyDow->wVKey == VK_DELETE) OnXdong();

	*pResult = 0;
}

void LaymauNgang::OnEnKillfocusT1()
{
	CString szval=L"";
	m_txt1.GetWindowTextW(szval);
	int num = _wtoi(szval);
	if(num<1) num=1;
	else if(num>8) num=8;
	else return;
	szval.Format(L"%d",num);
	m_txt1.SetWindowText(szval);
}


void LaymauNgang::OnBnClickedB8()
{
	OnBnClickedB3();

	if(sCodeName==L"shHSNTVL")
	{
		AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
		frm_QuycachOld open_dlg;
		open_dlg.DoModal();
	}
	else
	{
		AFX_MANAGE_STATE(AfxGetStaticModuleState( ));
		Dlg_QCLMBetong open_dlg;
		open_dlg.DoModal();
	}	
}


void LaymauNgang::OnEnChangeT3()
{
	// Xác định lại độ rộng các cột
	CString szval = L"";
	m_txt3.GetWindowTextW(szval);
	szval.Trim();

	if (szval != L"") ChangeColumnWidth(szval);
}

void LaymauNgang::ChangeColumnWidth(CString szColumn)
{
	try
	{
		if (szColumn.Right(1) == L"," || szColumn.Right(1) == L";") return;

		if (_wtoi(szColumn.Left(1)) == 0)
		{
			int iNum[10];
			int ncount = litLMN3.GetColumnCount();

			_TackTukhoaChuan(szColumn, L",", L";", 0, 0);
			if (iKey > ncount - 1) iKey = ncount - 1;
			
			for (int i = 1; i <= iKey; i++)
			{
				iNum[i] = 0;
				for (int j = 1; j < 27; j++)
				{
					if (pKey[i] == sColumnA1[j])
					{
						iNum[i] = j;
						break;
					}
				}

				if (iNum[i] == 0)
				{
					if (pKey[i] == L"AA") iNum[i] = 27;
					else if (pKey[i] == L"AB") iNum[i] = 28;
				}
			}

			iNum[iKey + 1] = 29;

			int num = 0;
			for (int i = 1; i <= iKey; i++)
			{
				num = iNum[i + 1] - iNum[i];
				if (num > 0) litLMN3.SetColumnWidth(i, 20 * num);
				else litLMN3.SetColumnWidth(i, 5);				
			}
		}
	}
	catch (...){}
}
