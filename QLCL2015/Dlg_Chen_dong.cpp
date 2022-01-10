#include "Dlg_Chen_dong.h"

Dlg_Chen_dong::Dlg_Chen_dong() : CDialog(Dlg_Chen_dong::IDD)
{
}

void Dlg_Chen_dong::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, CDGTXT1, txt1);
}

BEGIN_MESSAGE_MAP(Dlg_Chen_dong, CDialog)
	ON_BN_CLICKED(CDGBTN2, &Dlg_Chen_dong::OnBnClickedCdgbtn2)
	ON_BN_CLICKED(CDGBTN1, &Dlg_Chen_dong::OnBnClickedCdgbtn1)
END_MESSAGE_MAP()

BOOL Dlg_Chen_dong::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);
	
	if(iKqua==1)
	{
		this->SetWindowText(L"Số tầng");	
		CStatic* stt = (CStatic*)GetDlgItem(CDG_TIEUDE);
		stt->SetWindowText(L"Nhập số tầng");
		CString szval=L"";
		szval.Format(L"%d",iSLTang);
		txt1.SetWindowText(szval);
	}
	else txt1.SetWindowText(L"20");

	return TRUE; 
}

// Bắt sự kiện click Enter
BOOL Dlg_Chen_dong::PreTranslateMessage(MSG* pMsg)
{
	 CWnd* pWndWithFocus = GetFocus();
	if (pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_TAB) &&
		pWndWithFocus == GetDlgItem(CDGTXT1))
	{
		OnBnClickedCdgbtn1();
		return TRUE;
	}
	else if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		CDialog::OnCancel();
		return TRUE; 
	}

	return FALSE;
}



void Dlg_Chen_dong::OnBnClickedCdgbtn2()
{
	CDialog::OnCancel();
}


void Dlg_Chen_dong::OnBnClickedCdgbtn1()
{
	try
	{
		// Kiểm tra giá trị
		CString szval = L"";
		txt1.GetWindowTextW(szval);
		szval.Trim();
		int num = _wtoi(szval);
		if(num>0)
		{
			CDialog::OnOK();
			xl->PutScreenUpdating(VARIANT_FALSE);

			if(iKqua==0)
			{
				pSheet = pWb->ActiveSheet;
				pRange = pSheet->Cells;
				int Rchen = pSheet->Application->ActiveCell->Row;
				pSheet->GetRange(pRange->GetItem(Rchen,1),pRange->GetItem(Rchen+num-1,1))->EntireRow->Insert(xlShiftDown);
			}

			iKqua = num;
		}
		else
		{
			MessageBox(L"Giá trị nhập không hợp lệ.",L"Thông báo",MB_OK | MB_ICONWARNING);
			GotoDlgCtrl(GetDlgItem(CDGTXT1));
		}
	}
	catch(...){}
}

