#include "SeachCG.h"

// Class Seach chuyên gia
SeachCG::SeachCG() : CDialog(SeachCG::IDD)
{
}

void SeachCG::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, LISTCHG1, listTCG);
}

BEGIN_MESSAGE_MAP(SeachCG, CDialog)
	ON_BN_CLICKED(BTNCHG1, &SeachCG::OnBnClickedBtnchg1)
	ON_BN_CLICKED(BTNCHG2, &SeachCG::OnBnClickedBtnchg2)
	ON_NOTIFY(LVN_KEYDOWN, LISTCHG1, &SeachCG::OnLvnKeydownListchg1)
	ON_NOTIFY(NM_DBLCLK, LISTCHG1, &SeachCG::OnNMDblclkListchg1)
END_MESSAGE_MAP()

BOOL SeachCG::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	if(readCG()!=1)
	{
		MessageBox(L"Chưa có dữ liệu chuyên gia trong bảng tiền lương",L"Thông báo",MB_OK | MB_ICONWARNING);
		EndDialog(0);
		return FALSE;
	}

	// Load list
	listTCG.InsertColumn(0,L"Mã CG",LVCFMT_LEFT,60);
	listTCG.InsertColumn(1,L"Họ tên chuyên gia",LVCFMT_LEFT,250);
	listTCG.InsertColumn(2,L"Chức danh",LVCFMT_LEFT,100);
	listTCG.InsertColumn(3,L"Chi phí ngày công",LVCFMT_RIGHT,140);
	listTCG.InsertColumn(4,L"",LVCFMT_LEFT,0);
	listTCG.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);	

	CString szval=L"";
	if(num>0)
	for(int it=0; it<num;it++)
	{
		listTCG.InsertItem(it,CGArr[it].macg,0);
		listTCG.SetItemText(it,1,CGArr[it].hoten);
		listTCG.SetItemText(it,2,CGArr[it].chucd);

		szval.Format(L"%d",(long)fRoundDouble(_ttof(CGArr[it].cpnc),0));
		listTCG.SetItemText(it,3,FormatTien(szval,L",",L"."));

		szval.Format(L"%d",CGArr[it].index);
		listTCG.SetItemText(it,4,szval);
	}

	UpdateData(false);

	return TRUE; 
}


// Bắt sự kiện click Enter
BOOL SeachCG::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if (pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(LISTCHG1))
	{
		OnBnClickedBtnchg1();
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		OnCancel();
		return TRUE;  
	}

	return FALSE;
}


void SeachCG::fillcell(int irn)
{
	try
	{
		xl->PutScreenUpdating(VARIANT_FALSE);

		_bstr_t nameBSTR =  name_sheet("shCPCGia");
		_WorksheetPtr pSheet=xl->Sheets->GetItem(nameBSTR);
		RangePtr pRange = pSheet->Cells;
	
		//Bang tien luong (15)
		_bstr_t sh_TL = name_sheet("shCPTLuong");
		_WorksheetPtr ps_TL=xl->Sheets->GetItem(sh_TL);
		RangePtr pr_TL = ps_TL->Cells;

		// Kiểm tra nếu số lượng >=2 --> chèn dòng
		// Hoặc vị trí nhập cận END
		int iRow = pSheet->Application->ActiveCell->Row;
		int iEND = FindCEnd(pSheet,2,"END",iRow+2);
		if(iEND==0) return;

		RangePtr PRS;
		if(iRow+2==iEND || irn>=2)
		{
			PRS = pSheet->GetRange(pRange->GetItem(iRow+1,1),pRange->GetItem(iRow+irn,1));
			PRS->EntireRow->Insert(xlShiftDown);
		}

		// Thêm dữ liệu
		CString szval=L"";
		for (int i = 0; i < irn; i++)
		{
			pRange->PutItem(iRow+i,3,(_bstr_t)pCGArr[i].macg);
			pRange->PutItem(iRow+i,4,(_bstr_t)pCGArr[i].hoten);

			szval.Format(L"='%s'!%s",(LPCTSTR)ps_TL->Name,GAR(pCGArr[i].index,11,pr_TL,0));
			pRange->PutItem(iRow+i,7,(_bstr_t)szval);

			szval.Format(L"=%s*%s",GAR(iRow+i,6,pRange,0),GAR(iRow+i,7,pRange,0));
			pRange->PutItem(iRow+i,8,(_bstr_t)szval);
		}

		// Đánh lại stt
		int iBDAU=6;
		for (int i = iRow; i >=3; i--)
		{
			szval = pRange->GetItem(i,2);
			if(szval==L"STT")
			{
				iBDAU=i+1;
				break;
			}
		}

		iEND = FindCEnd(pSheet,2,"END",iRow+2);
		int stt=1;
		for (int i = iBDAU; i < iEND-1;i++ )
		{
			szval = pRange->GetItem(i,3);
			if(szval!=L"")
			{
				pRange->PutItem(i,2,stt);
				stt++;
			}
		}

		PRS = pSheet->GetRange(pRange->GetItem(iBDAU,2),pRange->GetItem(iEND-2,8));
		PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->Weight=xlThin;

		PRS->Borders->GetItem(XlBordersIndex::xlInsideHorizontal)->PutWeight(xlThin);
		PRS->Borders->GetItem(XlBordersIndex::xlInsideHorizontal)->PutLineStyle(xlDot);

		PRS->Font->PutBold(0);

		// Tính lại tổng tiền
		szval.Format(L"=SUM(%s:%s)",GAR(iBDAU,8,pRange,0),GAR(iEND-2,8,pRange,0));
		pRange->PutItem(iEND-1,8,(_bstr_t)szval);
	}
	catch(_com_error & error){}
}


int SeachCG::readCG()
{
	try
	{
		//Bang tien luong (15)
		_bstr_t sh_TL = name_sheet("shCPTLuong");
		_WorksheetPtr ps_TL = xl->Sheets->GetItem(sh_TL);
		RangePtr pr_TL = ps_TL->Cells;

		int dem=0;
		int iBDAU=6;
		CString szval=L"";

		for (int i = 3; i < 65000; i++)
		{
			szval = pr_TL->GetItem(i,2);
			if(szval==L"STT")
			{
				dem++;
				if(dem==iVMMn)
				{
					iBDAU=i+1;
					break;
				}
			}
		}

		int iEnd = FindCEnd(ps_TL,2,"END",iBDAU+1);
		if(iEnd==0) return 0;

		num=0;
		for(int i=iBDAU;i<iEnd;i++)
		{
			szval = pr_TL->GetItem(i,3);
			if(szval!=L"")
			{
				CGArr[num].macg=pr_TL->GetItem(i,3);
				CGArr[num].hoten=pr_TL->GetItem(i,4);
				CGArr[num].chucd=pr_TL->GetItem(i,5);

				szval = pr_TL->GetItem(i,11);
				szval.Replace(L",",L".");
				CGArr[num].cpnc.Format(L"%f",_ttof(szval));
				CGArr[num].index=i;
				num++;
			}
		}

		return 1;
	}
	catch (...){return 0;}
}


void SeachCG::OnBnClickedBtnchg1()
{
	inp=0;
	POSITION pos=listTCG.GetFirstSelectedItemPosition();
	if(pos!=NULL)
	{
		CString szval=L"";
		for (int i = 0; i < listTCG.GetItemCount();i++ )
			if (listTCG.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
			{
				pCGArr[inp].macg = listTCG.GetItemText(i,0);
				pCGArr[inp].hoten = listTCG.GetItemText(i,1);
				pCGArr[inp].chucd = listTCG.GetItemText(i,2);
				pCGArr[inp].cpnc = listTCG.GetItemText(i,3);

				szval = listTCG.GetItemText(i,4);
				pCGArr[inp].index = _wtoi(szval);

				inp++;
			}

		fillcell(inp);
		CDialog::OnOK();
	}
	else MessageBox(L"Bạn chưa lựa chọn dữ liệu",L"Thông báo",MB_OK);	
}


void SeachCG::OnBnClickedBtnchg2()
{
	OnCancel();
}


void SeachCG::OnLvnKeydownListchg1(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);
	if(pLVKeyDow->wVKey == VK_SPACE) OnBnClickedBtnchg1();

	*pResult = 0;
}


void SeachCG::OnNMDblclkListchg1(NMHDR *pNMHDR, LRESULT *pResult)
{
	OnBnClickedBtnchg1();
}
