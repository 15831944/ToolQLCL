// PHẦN MỀM QUẢN LÝ CHẤT LƯỢNG 8.0
// Chỉnh sửa lần cuối: 03.04.2018 (LEO)

/*
	29.07.2016
		- Sửa code phần in: In phần mở rộng
		- Sửa nút cập nhật trong tiêu chuẩn
		- Thêm mới tính năng quản lý WB
		- Thêm mới tính năng in các bảng tính khác

	16.08.2016
		- trường hợp tồn tại duy nhất thì vẫn hiển thị hộp thoại
		if(iKqua>=1) ...
		- Replate "Excel::"
		- getPathDB thay đổi đường dẫn CSDL

	07.09.2016
		- Thêm 2 lệnh load nhiều dữ liệu công việc & vật liệu
			call_import_xlsx
			call_import_cviec

	21.09.2016
		- Kiểm tra phần hỗ trợ in nhiều sheet khác (Dlg_Printer)

	23.09.2016
		- Bổ sung thư viện FolderDlg
		- Thêm dialog lưu hồ sơ
		- Code lưu hồ sơ cho phần in hồ sơ, in nhật ký và in khác

	28.09.2016
		- Sửa phần f_setCombobox trong thiết lập chế độ in
		- Code tiếp phần lưu hồ sơ
		- Lỗi xuất bảng vật liệu
		- Lỗi xuất bảng nhật ký

	13.10.2016
		- Hoàn thiện phần in XPS
		- Lưu ý code name shHSNTCVNBCV bị trùng bởi 3 sheet công việc
			 && sEdit.Left(13)!=L"shHSNTCVNBCV1" && sEdit.Left(13)!=L"shHSNTCVNBCV2"

	18.11.2016
		- Kiểm tra lỗi load ảnh công việc: Không tự xóa dòng
			+ float _iwidth = 400;
			+ PRS = pRange->GetItem(irow+3,icol+_inextpic);
			+ pRange->PutItem(irow,1,(_bstr_t)L"picture");
			+ for (int i = irow; i < irow+50; i++)
			{
				msg = GIOR(i,1,pRange,L"");
				if(msg==L"picture")
				{
					pSheet->GetRange(pRange->GetItem(irow,1),
						pRange->GetItem(irow+11,1))->EntireRow->Delete(xlShiftUp);
					break;
				}
			}

		- In thêm phần ghi chú ở cuối mỗi biên bản
		- Lỗi font khi thêm mới VL hoặc CV
			+ Lỗi font xảy ra với office 64, nguyên nhân là khi truy vấn sql update.
			+ khắc phục, viết lại hàm connection.h và 2 file Database.dll & Database.lib
			+ Copy file Database.dll vào thư mục cài đặt (bộ cài QLCL)
		- Autofit

	02.12.2016
		- Chỉnh phần ảnh đổ sai vị trí

	22.12.2016
		- Viết hàm lưu nhật ký thi công				call_save_nky
		- Viết hàm thêm hình ảnh cho nhật ký		call_addimage_nky
		- In hình ảnh đính kèm trong nhật ký		f_Import_sheetpic
		- Sửa .xla, thêm 2 hàm 
			+ call_addimage_nkVB()
			+ call_upCSDL_nkVB()

			+ Hàm gọi từ Macro
				Sub AddImageNKy()
				Sub UpdateCSDLNL()

	24.12.2016
	- Sửa lại hàm GetImageSize (Lấy kích thước ảnh đã qua chỉnh sửa photoshop)
	- Sửa lại cách get rowhight: if(iHThuc>0 && iWThuc>0)...
		+ Trước: y1 = (int)(_iwHinhve*iHThuc/iWThuc);
		+ Sau : y1 = (int)(3*_iwHinhve*iHThuc/(4*iWThuc));

	26.12.2016
	- Xong phần lưu & in ảnh nhật ký
	- Set date time control dt1.SetFormat(ngview)

	11.01.2017
	- Sửa temp, sửa xla (CF_EVDATE,CF_DMCOPY)
	- Set ngày tháng tiếng Anh
	- Kiểm tra điều kiện ngày tháng trong danh mục công việc + vật liệu

	19.01.2017
	- Chỉnh sửa template
	- Viết thêm hàm lưu nội dung nhật ký

	24.01.2017
	- Chỉnh temp, thêm name CF_XNDNK
	- Chỉnh lại option trong folder OpenCLGXD >> project DLLQLCL.sln >> OpenCLGXD.dll

	06.02.2017
	- Sửa quan trọng: Check bản quyền trong project OpenCSV
		if (check_count_key++ % 10 == 9)

	09.02.2017
	-	Phần in hồ sơ, khi chọn in riêng lẻ hồ sơ. Ô text rỗng thì phải lấy giá trị bên list
		f_Xac_dinh_sl_In

		// Trường hợp iKey = 0
		iKey = list1.GetItemCount();
		if(iKey>0) for (int i = 0; i < iKey; i++) iCViec[i+1] = list1.GetItemText(i,1);

	16.02.2017
	-	Viết thêm lệnh xuất hồ sơ theo ngày lấy mẫu hoặc nghiệm thu (trước chỉ xuất theo stt hoặc mã hồ sơ)
		--> Chỉnh cả trong xlam và xla

		Hàm call_truyenDL_sang_sh_dmHoSo bỏ đối, tạo thêm dialog DLG_SRTDMHS

	13.04.2017
	-	Bổ sung phần hiển thị hình ảnh cho các sheet
	-	Viết thêm hàm quy cách lấy mẫu

	18.04.2017
	-	Cho phép đảo vị trí HK_A, B,C,D,E trong sheet nhật ký
	-	Thêm check bản quyền cho tiện ích in ấn
		if(f_Check_ban_quyen()!=1) return;

		// Lưu ý đặc biệt: Khi khai báo biến toàn cục, ví dụ biến count_key giá trị mặc định ban đầu = 0
		// Tuy nhiên khi gọi hàm trung gian thông qua dll khác, giá trị ban đầu có thể chưa được khởi tạo

	22.04.2017
	-	Viết thêm hàm hiển thị các sheet bổ sung từ dự toán: call_sheet_bsdtoan

	27.04.2017
	-	Viết thêm hàm thêm chi phí chuyên gia
	-	Viết thêm hàm chuột phải gọi hộp thoại chi phí tư vấn: open_chiphituvan
	-	Viết hàm dự toán Man Month: call_create_dtmm

	22.06.2017
	-	Phát triển bản QLCL 8.0

	23.06.2017
	-	Sửa temp nhật ký, thay đổi cách lưu nhật ký, lưu thêm nhân công và máy

	07.08.2017
	-	Sửa code xuất nhật ký, cho phép thay đổi "Lấy mẫu" --> thành 1 cụm từ bất kỳ
	-	Đưa ký hiện LM sang config


	29.12.2017
	- Nhiều máy gặp trường hợp tra mã thứ 2 bị treo
	- Nguyên nhân có thể do khởi tạo lớp CListCtrlEx
	- Đưa đoạn khai báo biến khởi tạo vào bên trong thay vì để ngoài
	
	CListCtrlEx::CListCtrlEx()
{
	m_pEditor=NULL;
	m_nEditingRow=-1;
	m_nEditingColumn=-1;
	m_msgHook;
	m_nRow=-1;
	m_nColumn=-1;
	m_bHandleDelete=FALSE;
	m_nSortColumn=-1;
	m_fnCompare=NULL;
	m_hAccel=NULL;
	m_dwSortData=NULL;
	m_bInvokeAddNew=FALSE;

	#ifndef _WIN32_WCE
		EnableActiveAccessibility();
	#endif	
}

	08.02.2018
	- Viết lại hàm xuất khối lượng
	- Thêm lệnh hiện ẩn chi tiết công việc
	- Đưa tiêu chuẩn sang bảng tính mới

	03.04.2018
	- Thay đổi quan trọng: Chuyển hàm Enter_RightClick từ DLLLock sang đây
	--> Hàm chuột phải gọi hộp thoại quy cách lấy mẫu bê tông

*/