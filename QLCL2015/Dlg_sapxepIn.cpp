#include "stdafx.h"
#include "Dlg_sapxepIn.h"

// Dlg_sapxepIn dialog

IMPLEMENT_DYNAMIC(Dlg_sapxepIn, CDialog)

Dlg_sapxepIn::Dlg_sapxepIn(CWnd* pParent /*=NULL*/)
	: CDialog(Dlg_sapxepIn::IDD, pParent)
{

}

Dlg_sapxepIn::~Dlg_sapxepIn()
{
}

void Dlg_sapxepIn::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, SXIN_L1, m_list1_VL);
	DDX_Control(pDX, SXIN_L2, m_list2_CV);
	DDX_Control(pDX, SXIN_L3, m_list3_GD);
}


BEGIN_MESSAGE_MAP(Dlg_sapxepIn, CDialog)
	ON_BN_CLICKED(SXIN_OK, &Dlg_sapxepIn::OnBnClickedOk)
	ON_BN_CLICKED(SXIN_CANCEL, &Dlg_sapxepIn::OnBnClickedCancel)
	ON_BN_CLICKED(SXIN_B1, &Dlg_sapxepIn::OnBnClickedB1)
	ON_BN_CLICKED(SXIN_B4, &Dlg_sapxepIn::OnBnClickedB4)
	ON_BN_CLICKED(SXIN_B7, &Dlg_sapxepIn::OnBnClickedB7)
	ON_BN_CLICKED(SXIN_B2, &Dlg_sapxepIn::OnBnClickedB2)
	ON_BN_CLICKED(SXIN_B5, &Dlg_sapxepIn::OnBnClickedB5)
	ON_BN_CLICKED(SXIN_B8, &Dlg_sapxepIn::OnBnClickedB8)
	ON_BN_CLICKED(SXIN_B3, &Dlg_sapxepIn::OnBnClickedB3)
	ON_BN_CLICKED(SXIN_B6, &Dlg_sapxepIn::OnBnClickedB6)
	ON_BN_CLICKED(SXIN_B9, &Dlg_sapxepIn::OnBnClickedB9)
END_MESSAGE_MAP()


// Dlg_sapxepIn message handlers
BOOL Dlg_sapxepIn::OnInitDialog()
{
	CDialog::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(ICO_QLCL));
	SetIcon(hIcon, FALSE);

	// Định dạng kiểu List Control
	m_list1_VL.InsertColumn(0,L"",LVCFMT_LEFT,0);
	m_list1_VL.InsertColumn(1,L"Tên bảng",LVCFMT_LEFT,180);
	m_list1_VL.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);

	m_list2_CV.InsertColumn(0,L"",LVCFMT_LEFT,0);
	m_list2_CV.InsertColumn(1,L"Tên bảng",LVCFMT_LEFT,180);
	m_list2_CV.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);

	m_list3_GD.InsertColumn(0,L"",LVCFMT_LEFT,0);
	m_list3_GD.InsertColumn(1,L"Tên bảng",LVCFMT_LEFT,180);
	m_list3_GD.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);

	f_GetAllSheet();

	if(nccBBin==1)
	{
		isxin[1]=0;
		isxin[2]=0;
	}
	else if(nccBBin==2)
	{
		isxin[2]=0;
		isxin[0]=0;
	}
	else if(nccBBin==3)
	{
		isxin[0]=0;
		isxin[1]=0;
	}

	OnBnClickedB3();
	OnBnClickedB6();
	OnBnClickedB9();

	return TRUE; 
}


BOOL Dlg_sapxepIn::PreTranslateMessage(MSG* pMsg)
{
	CWnd* pWndWithFocus = GetFocus();
	if(pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_ESCAPE)
	{
		CDialog::OnCancel();
		return TRUE; 
	}

	return FALSE;
}

void Dlg_sapxepIn::OnBnClickedOk()
{
	CDialog::OnOK();
	iSapxepIN=1;

	if(isxin[0]>0)
	{
		for (int i = 0; i < isxin[0]; i++)
		{
			SXI_VL_cn[i+1] = m_list1_VL.GetItemText(i,0);
			SXI_VL_name[i+1] = m_list1_VL.GetItemText(i,1);
		}
	}

	if(isxin[1]>0)
	{
		for (int i = 0; i < isxin[1]; i++)
		{
			SXI_VL_cn[i+1] = m_list2_CV.GetItemText(i,0);
			SXI_VL_name[i+1] = m_list2_CV.GetItemText(i,1);
		}
	}

	if(isxin[2]>0)
	{
		for (int i = 0; i < isxin[2]; i++)
		{
			SXI_VL_cn[i+1] = m_list3_GD.GetItemText(i,0);
			SXI_VL_name[i+1] = m_list3_GD.GetItemText(i,1);
		}
	}
}


void Dlg_sapxepIn::OnBnClickedCancel()
{
	CDialog::OnCancel();
}


void Dlg_sapxepIn::f_GetAllSheet()
{
	try
	{
		_WorksheetPtr pSh1;
		nSheet = xl->ActiveWorkbook->Sheets->Count;
		isxin[0]=0,isxin[1]=0,isxin[2]=0;

		CString val=L"",sktr=L"";
		for (int i = 1; i <= nSheet; i++)
		{	
			pSh1 = pWb->Worksheets->GetItem(i);			
			val = (LPCTSTR)pSh1->CodeName;
			if(val.GetLength()<=8) continue;
			sktr = val.Left(8);			
			if(sktr==L"shHSNTVL")
			{
				isxin[0]++;
				SXI_VL_cn[isxin[0]] = val;
				SXI_VL_name[isxin[0]] = (LPCTSTR)pSh1->Name;
			}
			else if(sktr==L"shHSNTCV")
			{
				isxin[1]++;
				SXI_CV_cn[isxin[1]] = val;
				SXI_CV_name[isxin[1]] = (LPCTSTR)pSh1->Name;
			}
			else if(sktr==L"shHSNTGD")
			{
				isxin[2]++;
				SXI_GD_cn[isxin[2]] = val;
				SXI_GD_name[isxin[2]] = (LPCTSTR)pSh1->Name;
			}
		}
	}
	catch(...){}
}


void Dlg_sapxepIn::OnBnClickedB1()
{
	try
	{
		CString val=L"";
		POSITION pos=m_list1_VL.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			int nItem=0;
			for (int i = 0; i < m_list1_VL.GetItemCount(); i++)
				if (m_list1_VL.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					nItem=i;
					break;
				}
			
			if(nItem>0)
			{
				// Codename
				val = m_list1_VL.GetItemText(nItem,0);
				m_list1_VL.SetItemText(nItem,0,m_list1_VL.GetItemText(nItem-1,0));
				m_list1_VL.SetItemText(nItem-1,0,val);

				// Name
				val = m_list1_VL.GetItemText(nItem,1);
				m_list1_VL.SetItemText(nItem,1,m_list1_VL.GetItemText(nItem-1,1));
				m_list1_VL.SetItemText(nItem-1,1,val);

				// Vị trí lựa chọn mới
				m_list1_VL.SetItemState(nItem,0,LVIS_SELECTED);
				m_list1_VL.SetItemState(nItem-1,LVIS_SELECTED,LVIS_SELECTED);
			}
		}
	}
	catch(...){}
}


void Dlg_sapxepIn::OnBnClickedB4()
{
	try
	{
		CString val=L"";
		POSITION pos=m_list2_CV.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			int nItem=0;
			for (int i = 0; i < m_list2_CV.GetItemCount(); i++)
				if (m_list2_CV.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					nItem=i;
					break;
				}
			
			if(nItem>0)
			{
				// Codename
				val = m_list2_CV.GetItemText(nItem,0);
				m_list2_CV.SetItemText(nItem,0,m_list2_CV.GetItemText(nItem-1,0));
				m_list2_CV.SetItemText(nItem-1,0,val);

				// Name
				val = m_list2_CV.GetItemText(nItem,1);
				m_list2_CV.SetItemText(nItem,1,m_list2_CV.GetItemText(nItem-1,1));
				m_list2_CV.SetItemText(nItem-1,1,val);

				// Vị trí lựa chọn mới
				m_list2_CV.SetItemState(nItem,0,LVIS_SELECTED);
				m_list2_CV.SetItemState(nItem-1,LVIS_SELECTED,LVIS_SELECTED);
			}
		}
	}
	catch(...){}
}


void Dlg_sapxepIn::OnBnClickedB7()
{
	try
	{
		CString val=L"";
		POSITION pos=m_list3_GD.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			int nItem=0;
			for (int i = 0; i < m_list3_GD.GetItemCount(); i++)
				if (m_list3_GD.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					nItem=i;
					break;
				}
			
			if(nItem>0)
			{
				// Codename
				val = m_list3_GD.GetItemText(nItem,0);
				m_list3_GD.SetItemText(nItem,0,m_list3_GD.GetItemText(nItem-1,0));
				m_list3_GD.SetItemText(nItem-1,0,val);

				// Name
				val = m_list3_GD.GetItemText(nItem,1);
				m_list3_GD.SetItemText(nItem,1,m_list3_GD.GetItemText(nItem-1,1));
				m_list3_GD.SetItemText(nItem-1,1,val);

				// Vị trí lựa chọn mới
				m_list3_GD.SetItemState(nItem,0,LVIS_SELECTED);
				m_list3_GD.SetItemState(nItem-1,LVIS_SELECTED,LVIS_SELECTED);
			}
		}
	}
	catch(...){}
}


void Dlg_sapxepIn::OnBnClickedB2()
{
	try
	{
		CString val=L"";
		POSITION pos=m_list1_VL.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			int nItem=0;
			for (int i = 0; i < m_list1_VL.GetItemCount(); i++)
				if (m_list1_VL.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					nItem=i;
					break;
				}

			if(nItem<m_list1_VL.GetItemCount()-1)
			{	
				// Codename
				val = m_list1_VL.GetItemText(nItem,0);
				m_list1_VL.SetItemText(nItem,0,m_list1_VL.GetItemText(nItem+1,0));
				m_list1_VL.SetItemText(nItem+1,0,val);

				// Name
				val = m_list1_VL.GetItemText(nItem,1);
				m_list1_VL.SetItemText(nItem,1,m_list1_VL.GetItemText(nItem+1,1));
				m_list1_VL.SetItemText(nItem+1,1,val);

				// Vị trí lựa chọn mới
				m_list1_VL.SetItemState(nItem,0,LVIS_SELECTED);
				m_list1_VL.SetItemState(nItem+1,LVIS_SELECTED,LVIS_SELECTED);
			}
		}
	}
	catch(...){}
}


void Dlg_sapxepIn::OnBnClickedB5()
{
	try
	{
		CString val=L"";
		POSITION pos=m_list2_CV.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			int nItem=0;
			for (int i = 0; i < m_list2_CV.GetItemCount(); i++)
				if (m_list2_CV.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					nItem=i;
					break;
				}

			if(nItem<m_list2_CV.GetItemCount()-1)
			{	
				// Codename
				val = m_list2_CV.GetItemText(nItem,0);
				m_list2_CV.SetItemText(nItem,0,m_list2_CV.GetItemText(nItem+1,0));
				m_list2_CV.SetItemText(nItem+1,0,val);

				// Name
				val = m_list2_CV.GetItemText(nItem,1);
				m_list2_CV.SetItemText(nItem,1,m_list2_CV.GetItemText(nItem+1,1));
				m_list2_CV.SetItemText(nItem+1,1,val);

				// Vị trí lựa chọn mới
				m_list2_CV.SetItemState(nItem,0,LVIS_SELECTED);
				m_list2_CV.SetItemState(nItem+1,LVIS_SELECTED,LVIS_SELECTED);
			}
		}
	}
	catch(...){}
}


void Dlg_sapxepIn::OnBnClickedB8()
{
	try
	{
		CString val=L"";
		POSITION pos=m_list3_GD.GetFirstSelectedItemPosition();
		if(pos!=NULL)
		{
			int nItem=0;
			for (int i = 0; i < m_list3_GD.GetItemCount(); i++)
				if (m_list3_GD.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					nItem=i;
					break;
				}

			if(nItem<m_list3_GD.GetItemCount()-1)
			{	
				// Codename
				val = m_list3_GD.GetItemText(nItem,0);
				m_list3_GD.SetItemText(nItem,0,m_list3_GD.GetItemText(nItem+1,0));
				m_list3_GD.SetItemText(nItem+1,0,val);

				// Name
				val = m_list3_GD.GetItemText(nItem,1);
				m_list3_GD.SetItemText(nItem,1,m_list3_GD.GetItemText(nItem+1,1));
				m_list3_GD.SetItemText(nItem+1,1,val);

				// Vị trí lựa chọn mới
				m_list3_GD.SetItemState(nItem,0,LVIS_SELECTED);
				m_list3_GD.SetItemState(nItem+1,LVIS_SELECTED,LVIS_SELECTED);
			}
		}
	}
	catch(...){}
}


void Dlg_sapxepIn::OnBnClickedB3()
{
	m_list1_VL.DeleteAllItems();
	ASSERT(m_list1_VL.GetItemCount() == 0);

	if(isxin[0]>0)
	{
		m_list1_VL.EnableWindow(1);
		for (int i = 0; i < isxin[0]; i++)
		{
			m_list1_VL.InsertItem(i,SXI_VL_cn[i+1],0);
			m_list1_VL.SetItemText(i,1,SXI_VL_name[i+1]);
		}
	}

}


void Dlg_sapxepIn::OnBnClickedB6()
{
	m_list2_CV.DeleteAllItems();
	ASSERT(m_list2_CV.GetItemCount() == 0);

	if(isxin[1]>0)
	{
		m_list2_CV.EnableWindow(1);
		for (int i = 0; i < isxin[1]; i++)
		{
			m_list2_CV.InsertItem(i,SXI_CV_cn[i+1],0);
			m_list2_CV.SetItemText(i,1,SXI_CV_name[i+1]);
		}
	}	
}


void Dlg_sapxepIn::OnBnClickedB9()
{
	m_list3_GD.DeleteAllItems();
	ASSERT(m_list3_GD.GetItemCount() == 0);

	if(isxin[2]>0)
	{
		m_list3_GD.EnableWindow(1);
		for (int i = 0; i < isxin[2]; i++)
		{
			m_list3_GD.InsertItem(i,SXI_GD_cn[i+1],0);
			m_list3_GD.SetItemText(i,1,SXI_GD_name[i+1]);
		}
	}	
}
