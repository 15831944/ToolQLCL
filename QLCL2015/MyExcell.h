#ifdef MYEXCELL_EXPORTS
	#define MYEXCELL_API __declspec(dllexport)
#else
	#define MYEXCELL_API __declspec(dllexport)
#endif

//namespace G_EXCELL {
//	class MyExcell {
//	public:
		
		MYEXCELL_API long getColumn(Excel::RangePtr range, char* name);
		MYEXCELL_API long getRow(Excel::RangePtr range, char* name);
		MYEXCELL_API _variant_t NameSheet(char* sheetName);
		MYEXCELL_API _variant_t GetValueOfRange(Excel::RangePtr range, char* name);
		MYEXCELL_API CString getValueRange(Excel::RangePtr range, char* name);
		MYEXCELL_API bool PutValueToRange(RangePtr range, char* name, _variant_t szVal);
		MYEXCELL_API CString GAR(int nRow,int nCol,RangePtr pRng,int st);        //st=0 kiểu A1 , st=1 kiểu $A1, st=2 kiểu A$1, st=3 kiểu $A$1
		MYEXCELL_API long FindComment(Excel::_WorksheetPtr pSheet, long col, _variant_t comment, long bMes);
		MYEXCELL_API long FindCommentCustom(Excel::_WorksheetPtr pSheet, long col, _variant_t comment, bool bMes, long iRowB, long iRowE);
//	};
//}