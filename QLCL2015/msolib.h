// Office XP Objects (2002)
#import "..\\Office\\MSO.DLL" \
	rename("DocumentProperties", "DocumentPropertiesXL") \
	rename("RGB", "RBGXL")
//Microsoft VBA objects
#import "..\\Office\\VBE6EXT.OLB"
//Excel Application objects
#import "..\\Office\\EXCEL.EXE" \
	rename("DialogBox", "DialogBoxXL") rename("RGB", "RBGXL") \
	rename("DocumentProperties", "DocumentPropertiesXL") \
	rename("ReplaceText", "ReplaceTextXL") \
	rename("CopyFile", "CopyFileXL") \
	exclude("IFont", "IPicture") no_dual_interfaces

using namespace Excel;