# OfficeAutomation
Upp Package to perform Microsoft Office Automation

________________________________________

#include <Core/Core.h>
#include <OfficeAutomation/OfficeAutomation.h>

using namespace Upp;


CONSOLE_APP_MAIN
{
	ExcelApp myExcel;


	try{
		myExcel.FindOrStart();
		myExcel.SetVisible(true);
		Workbook wb = myExcel.OpenWorkbook(" WB PATH ");
		Sheet ws = wb.Sheets("SHEET NAME ");
	
		Range myRange = ws.Range("A1:B5");
		myRange.Value("CECI EST UN TEST");
		Sleep(1000); //Sleep is not required, it only allow you to see what happen
		Cout() << ws.GetLastRow("A") << "\n";
		Sleep(1000);
		Cout() << ws.Cells(1,1).Value() << "\n";
		Sleep(1000);
		Cout() << ws.Cells(1,1).Value("TEST") << "\n";
		Sleep(1000);
		Cout() << ws.Cells(1,1).Value() << "\n";
		Sleep(1000);
		wb.Close();
		Sleep(1000);
		myExcel.Quit();



		
	
	}catch(const OleException & exception){
		Cout() << String(exception.what()) <<"\n";
		myExcel.Quit();
	}

}
