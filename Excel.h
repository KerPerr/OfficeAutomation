#ifndef _OfficeAutomation_Excel_h_
#define _OfficeAutomation_Excel_h_
#include "OfficeAutomation.h"

/* 
 Project created 01/18/2019 
 By Clément Hamon And Pierre Castrec
 Lib used to drive every Microsoft Application's had OLE Compatibility.
 This project have to be used with Ultimate++ FrameWork and required the Core Librairy from it
*/

class ExcelApp; //Class represents an   Excel Application 
class ExcelWorkbook; //Class represents Excel Workbook
class ExcelSheet; //Class represents an Excel WorkSheet
class ExcelRange; //Class represents an Excel Range

class ExcelApp : public Ole {
	private: 
		bool ExcelIsStarted; //Bool to know if we started Excel
		Upp::Vector<ExcelWorkbook> workbooks; //Vector of every workbook
	
	public:
		ExcelApp(); //Initialise COM
		~ExcelApp(); //Unitialise COM
		
		ExcelWorkbook Workbooks(int index); //Allow to retrieve workbook by is index 
		ExcelWorkbook Workbooks(Upp::String name); //Allow to retrieve workbook by is name
		
		bool Start(); //Start new Excel Applicatio
		bool FindOrStart(); //Find running Excel or Start new One
		bool Quit(); //Close current Excel Application
		
		bool FindApplication(); //Find First current Excel Application still openned
		
		bool SetVisible(bool set); //Set or not the application visible 
		
		bool NewWorkbook(); //Create new Workbook and add it to actual excel Running method
		bool OpenWorkbook(Upp::String FilePath); //Find and Open Workbook by FilePath
		
		int GetNumberOfWorkbook(); //Return number of workbook currently openned on this excel App
	
};

class ExcelWorkbook : public Upp::Moveable<ExcelWorkbook>, public Ole {
	private:
		ExcelApp* parent; //Pointer to excelApp
		Upp::Vector<ExcelSheet> sheets; //Vector of every Worksheets
		bool isOpenned; //This bool must be useless But I prefere to have in case of object is still present in memory by a missing unreferenced pointer
		
	public:
		ExcelWorkbook(ExcelApp &parent,VARIANT AppObj);
		~ExcelWorkbook();
		
		ExcelSheet Sheets(int index);//Allow to retrieve worksheet by is index 
		ExcelSheet Sheets(Upp::String name);//Allow to retrieve worksheet by is name
		
		bool AddSheet(); //Create new Sheet with default Name
		bool AddSheet(Upp::String sheetName); //Create new Sheet with defined name 
		
		bool isReadOnly(); //Return true if the workbook is readOnly
		
		bool Save(); //Save current workbook
		bool SaveAs(Upp::String filePath); //Save current workbook at filePath
		bool Close(); //Close current workbook
};

class ExcelSheet :public Upp::Moveable<ExcelSheet>, public Ole  {
	private:
		ExcelWorkbook* parent;//Pointer to excelworkbook
	public:
		ExcelSheet(ExcelWorkbook &parent,VARIANT AppObj);
		~ExcelSheet();
		
		ExcelRange Range(Upp::String range); //Return a Range
		ExcelRange Cells(int ligne, int colonne); //Return a Cells
		
		int GetLastRow(Upp::String Colonne);
		
		bool setName(Upp::String name); //Redefine name of sheet 
};

class ExcelRange : public Ole {
	private:
		ExcelSheet* parent; //Pointer to excelWorkbook
	public:
		
		ExcelRange(ExcelSheet &parent,VARIANT AppObj);
		~ExcelRange();
		
		Upp::String Value(); //Return value of range
		bool Value(Upp::String value);//Set value of range
		
};

#endif
