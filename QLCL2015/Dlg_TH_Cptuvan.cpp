#include "stdafx.h"
#include "Dlg_TH_Cptuvan.h"
#include "cls_main.h"

Dlg_TH_Cptuvan::Dlg_TH_Cptuvan(CWnd* pParent /*=NULL*/)
	: CDialog(Dlg_TH_Cptuvan::IDD, pParent)
{
	
}

void Dlg_TH_Cptuvan::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, THCPTV_L1, m_list);
}

BEGIN_MESSAGE_MAP(Dlg_TH_Cptuvan, CDialog)
	ON_WM_SIZE()
	ON_WM_SIZING()

	ON_BN_CLICKED(THCPTV_B1, &Dlg_TH_Cptuvan::OnBnClickedB1)
	ON_BN_CLICKED(THCPTV_B2, &Dlg_TH_Cptuvan::OnBnClickedB2)
	ON_NOTIFY(NM_DBLCLK, THCPTV_L1, &Dlg_TH_Cptuvan::OnNMDblclkL1)
	ON_NOTIFY(LVN_KEYDOWN, THCPTV_L1, &Dlg_TH_Cptuvan::OnLvnKeydownL1)
	ON_WM_SYSCOMMAND()
END_MESSAGE_MAP()

BEGIN_EASYSIZE_MAP(Dlg_TH_Cptuvan)
	EASYSIZE(THCPTV_L1,ES_BORDER,ES_BORDER,ES_BORDER,ES_BORDER,0)
	EASYSIZE(THCPTV_B1,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
	EASYSIZE(THCPTV_B2,ES_KEEPSIZE,ES_KEEPSIZE,ES_BORDER,ES_BORDER,0)
END_EASYSIZE_MAP

BOOL Dlg_TH_Cptuvan::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	f_Style_List();
	f_Load_data();

	INIT_EASYSIZE;

	// Set size dialog mfc
	this->SetWindowPos(NULL,0,0,irW2,irH2,SWP_NOMOVE | SWP_NOZORDER);

	return TRUE; 
}

// Bắt sự kiện click Enter
BOOL Dlg_TH_Cptuvan::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if (pMsg->message == WM_KEYDOWN &&
		pMsg->wParam == VK_RETURN &&
		pWndWithFocus == GetDlgItem(THCPTV_L1))
	{
		OnBnClickedB1();
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

		OnBnClickedB2();
		return TRUE; 
	}

	return FALSE;
}

void Dlg_TH_Cptuvan::OnSize(UINT nType, int cx, int cy) 
{
	CDialog::OnSize(nType, cx, cy);
	UPDATE_EASYSIZE;
}

void Dlg_TH_Cptuvan::OnSizing(UINT fwSide, LPRECT pRect) 
{
	CDialog::OnSizing(fwSide, pRect);
	EASYSIZE_MINSIZE(50,50,fwSide,pRect);	// chiều rộng + chiều cao
}


void Dlg_TH_Cptuvan::f_Style_List()
{
	// Định dạng kiểu List Control
	m_list.InsertColumn(0,L"",LVCFMT_CENTER,0);
	m_list.InsertColumn(1,L"Khoản mục chi phí",LVCFMT_LEFT,jW2[1]);
	m_list.InsertColumn(2,L"Cách tính",LVCFMT_CENTER,jW2[2]);
	m_list.InsertColumn(3,L"Cơ sở pháp lý",LVCFMT_LEFT,jW2[3]);
	m_list.InsertColumn(4,L"Ghi chú",LVCFMT_LEFT,jW2[4]);
	m_list.InsertColumn(5,L"Định mức tỷ lệ",LVCFMT_LEFT,jW2[5]);
	m_list.InsertColumn(6,L"Giá trị trươc thuế",LVCFMT_CENTER,jW2[6]);
	m_list.InsertColumn(7,L"VAT",LVCFMT_CENTER,jW2[7]);

	m_list.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_CHECKBOXES | LVS_EX_GRIDLINES);
}


void Dlg_TH_Cptuvan::f_Get_size()
{
	CRect rectCtrl;
	GetWindowRect(&rectCtrl);
	irH2 = rectCtrl.Height();
	irW2 = rectCtrl.Width();
}


void Dlg_TH_Cptuvan::f_Get_width()
{
	for (int i = 0; i < 8; i++)
		jW2[i] = m_list.GetColumnWidth(i);
}


void Dlg_TH_Cptuvan::OnBnClickedB1()
{
	try
	{
		int count = m_list.GetItemCount();
		if(count==0) return;

		POSITION pos=m_list.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			pSheet = pWb->ActiveSheet;
			pRange = pSheet->Cells;
			int _irow = pSheet->Application->ActiveCell->Row;
			int iEnd = FindComment(pSheet,2,"END");

			// Kiểm tra vị trị activate
			if(_irow+1==iEnd)
			{
				PRS = pRange->GetItem(iEnd,1);
				PRS->EntireRow->Insert(xlShiftDown);

				// Định dạng lại khung
				PRS = pSheet->GetRange(pRange->GetItem(iEnd,1),pRange->GetItem(iEnd-1,1));
				PRS->EntireRow->Font->PutBold(0);
				PRS->EntireRow->Font->PutItalic(0);
				PRS->EntireRow->Interior->PutColor(xlNone);

				PRS = pSheet->GetRange(pRange->GetItem(iEnd,3),pRange->GetItem(iEnd-1,3));
				PRS->PutHorizontalAlignment(xlLeft);

				iEnd++;
			}

			int dem=0;
			for (int i = 0; i < count; i++)
				if (m_list.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED) dem++;

			if(dem>=2)
			{
				// Chèn thêm dòng & tính lại công thức
				PRS = pSheet->GetRange(pRange->GetItem(_irow+1,1),pRange->GetItem(_irow+dem,1));
				PRS->EntireRow->Insert(xlShiftDown);
				iEnd+=dem;
			}

			// Xác định tên gọi phương thức lập dự toán Man Month (mặc định sTenMM = "Lập dự toán riêng")
			shTS = name_sheet("shTS");	
			psTS = xl->Sheets->GetItem(shTS);
			prTS = psTS->Cells;

			CString sTenMM = getVCell(psTS,L"CF_LDTR");

			dem=0;
			int iA=0;
			CString vlgan[8]={L"",L"",L"",L"",L"",L"",L"",L""};
			for (int i = 0; i < count; i++)
				if (m_list.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					iA = _irow+dem;
					for (int j = 1; j < 8; j++) vlgan[j] = m_list.GetItemText(i,j);

					// Kiểm tra cách tính
					vlgan[0] = pRange->GetItem(iA,6);
					if(vlgan[0]==sTenMM)
					{
						// Xóa các bảng tính... 
						_bstr_t bx = pRange->GetItem(iA,2);
						_bstr_t by = pRange->GetItem(iA,12);
						if(bx!=(_bstr_t)L"" && by!=(_bstr_t)L"")
						{
							cls_main _fcall;
							_fcall.f_Delete_table(bx);

							// Xóa dữ liệu
							RangePtr PRS = pRange->GetItem(_irow,1);
							PRS->EntireRow->ClearContents();
						}
					}

					// tên khoản mục
					pRange->PutItem(iA,3,(_bstr_t)vlgan[1]);

					if(vlgan[2]!=L"") vlgan[2] = L"'" + vlgan[2];	// cách tính
					pRange->PutItem(iA,6,(_bstr_t)vlgan[2]);

					if(vlgan[5]!=L"") vlgan[5] = L"=" + vlgan[5];	// định mức tỷ lệ
					pRange->PutItem(iA,4,(_bstr_t)vlgan[5]);					

					if(vlgan[6]!=L"") vlgan[6] = L"=" + vlgan[6];	// giá trị trước thuế
					pRange->PutItem(iA,7,(_bstr_t)vlgan[6]);

					// thuế gtgt
					vlgan[0]=L"";
					vlgan[7].Trim();
					if(vlgan[7]!=L"") vlgan[0].Format(L"=%s*%s",vlgan[7],GAR(iA,7,pRange,0));
					pRange->PutItem(iA,8,(_bstr_t)vlgan[0]);

					// giá trị sau thuế
					vlgan[0].Format(L"=%s+%s",GAR(iA,7,pRange,0),GAR(iA,8,pRange,0));
					pRange->PutItem(iA,9,(_bstr_t)vlgan[0]);

					// cơ sở pháp lý
					pRange->PutItem(iA,10,(_bstr_t)vlgan[3]);

					// ghi chú
					PRS = pRange->GetItem(iA,5);
					ValidationPtr vPr = PRS->Validation;
					vPr->Delete();
					vPr->Add(XlDVType::xlValidateInputOnly,XlDVAlertStyle::xlValidAlertStop,
						XlFormatConditionOperator::xlBetween,vtMissing,vtMissing);

					if(vlgan[4].GetLength()>250) vlgan[0] = vlgan[4].Left(250) + L"...";
					else vlgan[0] = vlgan[4];
					vPr->PutInputTitle((_bstr_t)L"");
					vPr->PutInputMessage((_bstr_t)vlgan[0]);

					// Bỏ wraptext 2 cột bên trên
					PRS=pRange->GetItem(iA,10);
					PRS->PutWrapText(0);

					// Lập dự toán riêng
					PRS=pRange->GetItem(iA,12);
					PRS->ClearContents();

					dem++;
				}

			// Đánh lại stt
			for (int i = 5; i < iEnd; i++)
			{
				vlgan[0] = pRange->GetItem(i,3);
				if(vlgan[0]!=L"") pRange->PutItem(i,2,i-4);
			}
		}

		f_Get_width();
		f_Get_size();
		OnOK();
	}
	catch(...){}
}


void Dlg_TH_Cptuvan::OnBnClickedB2()
{
	f_Get_width();
	f_Get_size();
	CDialog::OnCancel();
}


long Dlg_TH_Cptuvan::f_GetLineFile(CString szPath)
{
	FILE * pFile;
	wchar_t *c = (wchar_t *)malloc(2000*sizeof(wchar_t));
	long line = 0;
	pFile=_wfopen(szPath,L"r, ccs=UTF-16LE");
	if (pFile==NULL) perror ("Error opening file");
	else
	{
		while (fgetws(c, 2000, pFile))
		{
			line++;
		}
		fclose (pFile);
		printf ("File contains %d.\n",line);
	}
	free(c);
	return line;
}


void Dlg_TH_Cptuvan::f_Load_data()
{
	try
	{
		CString spath=L"";
		spath.Format(L"%s%s",_fGPathFolder(L"ToolQLCL.dll"),_dFileCPTV);

		wchar_t **dArray;
		int numberOfLine = 0;
		int i;

		long _size;
		_size= f_GetLineFile(spath);
		_size++;
		dArray = (wchar_t **)calloc(_size,sizeof(wchar_t *));
		wchar_t **p;
		p = (wchar_t **)calloc(_size,sizeof(wchar_t *));
		for (i = 0; i < _size; i++)
		{
			*(dArray+i) = (wchar_t *)malloc(SIZE_LINE*sizeof(wchar_t));
			*(p+i) = (wchar_t *)malloc(SIZE_LINE*sizeof(wchar_t));
		}
		FILE *pReadFile;

		if(_wfopen_s(&pReadFile, spath, L"r, ccs=UTF-16LE") != 0)
		{
			_s(L"Không thể mở được tệp. Vui lòng chọn kiểm tra tệp dữ liệu.");
			for (i = 0; i < _size; i++)
			{
				free(*(dArray+i));
			}
			free(dArray);
			return;
		}	
		else
		{
			i=0;	
			CString szVarString=L"";
			while(fgetws(*(p+i), ROW_LINE, pReadFile)) i++;
			
			numberOfLine= i;
			fclose(pReadFile);

			int fdem=0;
			CString vlgan[8]={L"",L"",L"",L"",L"",L"",L"",L""};
			for(int i= 1;i<numberOfLine;i++)
			{
				wcscpy_s(*(dArray+i), wcslen(*(p+i))+1, *(p+i));
				wchar_t *textString = new wchar_t[SIZE_LINE];
				wcscpy_s(textString,wcslen(*(dArray+i))+1, *(dArray+i));
				szVarString= textString;
				long length;
				CString szLeft;
				int nToken= szVarString.Find(L'\t');
			
				fdem=0;
				while(nToken>=0)
				{
					length= szVarString.GetLength();
					szLeft= szVarString.Left(nToken);
					szLeft.TrimLeft();szLeft.TrimRight();
					szVarString=szVarString.Right(length-nToken-1);
					nToken= szVarString.Find(L'\t');

					fdem++;
					szLeft.Replace(L"\"",L"");
					szLeft.Replace(L"'",L"");
					if(fdem==1) vlgan[1] = szLeft;			// Tên khoản mục
					else if(fdem==2) vlgan[2] = szLeft;		// Cách tính
					else if(fdem==3) vlgan[3] = szLeft;		// Cơ sở pháp lý
					else if(fdem==4) vlgan[4] = szLeft;		// Ghi chú
					else if(fdem==5) vlgan[5] = szLeft;		// Định mức tỷ lệ
					else if(fdem==6) vlgan[6] = szLeft;		// Giá trị trước thuế

					// Thêm vào để lấy cột cuối cùng (by Leo 23.06.15)
					vlgan[7] = L"";
					if(nToken<0)
					{
						szVarString.Replace(L"\"",L"");
						vlgan[7]=szVarString;		// Thuế VAT
					}
				}

				// Load lên list
				vlgan[1].Trim();
				if(vlgan[1]!=L"")
				{
					m_list.InsertItem(i-1,L"",0);
					for (int j = 1; j < 8; j++) m_list.SetItemText(i-1,j,vlgan[j]);
				}

				free(textString);
			}

			fclose(pReadFile);
		}
	
		//free memory
		for (i = 0; i < _size; i++)
		{
			free(*(dArray+i));
			free(*(p+i));
		}
		free(dArray);
		free(p);
	}
	catch(_com_error & error){_s(L"Lỗi đọc dữ liệu csv.");}
}

void Dlg_TH_Cptuvan::OnNMDblclkL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	OnBnClickedB1();
}


void Dlg_TH_Cptuvan::OnLvnKeydownL1(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);
	if(pLVKeyDow->wVKey == VK_SPACE) OnBnClickedB1();

	*pResult = 0;
}


void Dlg_TH_Cptuvan::OnSysCommand(UINT nID, LPARAM lParam)
{
	if(nID == SC_CLOSE)
	{
		f_Get_width();
		f_Get_size();
		CDialog::OnCancel();
	}
	else CDialog::OnSysCommand(nID, lParam);
}
