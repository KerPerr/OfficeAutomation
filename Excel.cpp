#include <Core/Core.h>


#include <windows.h>
#include <ole2.h>
#include <stdio.h>
#include "Excel.h"
#include "OfficeAutomation.h"

#define DISP_FREEARGS 0x01
#define DISP_NOSHOWEXCEPTIONS 0x02

using namespace Upp;

ExcelApp::ExcelApp(){
	this->ExcelIsStarted=false;
	CoInitialize(NULL);
}

ExcelApp::~ExcelApp(){
	CoUninitialize();
}

bool ExcelApp::Start() //Start new Excel Application
{
	if(!this->ExcelIsStarted){
		this->AppObj = this->StartApp(WS_ExcelApp);
		if( this->AppObj.intVal != -1){
			this->ExcelIsStarted=true;
			return true;
		}
		return false;
	}
	return false;
}

bool ExcelApp::Quit() //Close current Excel Application
{
	if(this->ExcelIsStarted){
		try{
			this->ExecuteMethode("Quit",0);	
			return true;
		}catch(...){
			return false;
		}
	}
	return false;
}

bool ExcelApp::SetVisible(bool set)//Set or not the application visible 
{
	if(this->ExcelIsStarted){
		try{
			this->SetAttribute("Visible",(int)set);
			return true;
		}catch(...){
			return false;
		}
	}
	return false;
}

bool ExcelApp::FindOrStart(){
		if(!this->ExcelIsStarted){
		CLSID clsExcelApp;
		VARIANT xlApp = {0};

	   if(FAILED(CLSIDFromProgID(WS_ExcelApp, &clsExcelApp))) {
	      this->ExcelIsStarted=false;
	      return this->Start();
	   }
	   IUnknown *pUnk;
	   HWND hExcelMainWnd = 0;
	   hExcelMainWnd = FindWindow("XLMAIN",NULL);
	   if(hExcelMainWnd) {
		   SendMessage(hExcelMainWnd,WM_USER + 18, 0, 0);
			HRESULT hr2 = GetActiveObject(clsExcelApp,NULL,(IUnknown**)&pUnk);
			if (!FAILED(hr2)) {
				hr2=pUnk->QueryInterface(IID_IDispatch, (void **)&xlApp.pdispVal);
				if (!xlApp.ppdispVal) {
					this->ExcelIsStarted=false;
					return this->Start();
				}
			}
			if (pUnk) pUnk->Release();
		}
		else {
			this->ExcelIsStarted=false;
			return this->Start();
		}
		this->AppObj = xlApp;
		this->ExcelIsStarted=true;
		return true;
			
	}
	return false;
}

bool ExcelApp::FindApplication(){
	if(!this->ExcelIsStarted){
		CLSID clsExcelApp;
		VARIANT xlApp = {0};

	   if(FAILED(CLSIDFromProgID(WS_ExcelApp, &clsExcelApp))) {
	      this->ExcelIsStarted=false;
	      return false;
	   }
	   IUnknown *pUnk;
	   HWND hExcelMainWnd = 0;
	   hExcelMainWnd = FindWindow("XLMAIN",NULL);
	   if(hExcelMainWnd) {
		   SendMessage(hExcelMainWnd,WM_USER + 18, 0, 0);
			HRESULT hr2 = GetActiveObject(clsExcelApp,NULL,(IUnknown**)&pUnk);
			if (!FAILED(hr2)) {
				hr2=pUnk->QueryInterface(IID_IDispatch, (void **)&xlApp.pdispVal);
				if (!xlApp.ppdispVal) {
					this->ExcelIsStarted=false;
					return false;
				}
			}
			if (pUnk) pUnk->Release();
		}
		else {
			this->ExcelIsStarted=false;
			return false;
		}
		this->AppObj = xlApp;
		this->ExcelIsStarted=true;
		return true;
			
	}
	return false;
}

ExcelWorkbook* ExcelApp::NewWorkbook(){
	if(this->ExcelIsStarted){
		return &workbooks.Add(ExcelWorkbook(*this,GetAttribute(GetAttribute("Workbooks"),"Add")));
	}
	return NULL;
}

ExcelWorkbook* ExcelApp::Workbooks(int index){
	if(this->ExcelIsStarted && workbooks.GetCount() > index){
		return &workbooks[index];
	}
	return NULL;
}

int ExcelApp::GetNumberOfWorkbook(){
	return workbooks.GetCount();
}

ExcelWorkbook* ExcelApp::Workbooks(Upp::String name){
	if(this->ExcelIsStarted){
		for(int i = 0; i< workbooks.GetCount(); i++){
			if (BSTRtoString(workbooks[i].GetAttribute("Name").bstrVal).Compare(name) ==0){
				return &workbooks[i];
			}
		}
	}
	return NULL;
}

ExcelWorkbook* ExcelApp::OpenWorkbook(Upp::String name){
  	if( !FileExists(name.ToStd().c_str())) {
      return NULL;
    }
    ExcelWorkbook* myExcel=  &workbooks.Add(ExcelWorkbook(*this,ExecuteMethode(GetAttribute(L"Workbooks"),L"Open",1,AllocateString(name))));
    myExcel->ResolveSheet();
    return myExcel;
}

bool ExcelWorkbook::ResolveSheet(){
	int nbrSheet = this->GetAttribute( this->GetAttribute("Sheets"),"Count").intVal;
	for(int i = 0; i < nbrSheet; i++){
		sheets.Add(ExcelSheet(*this, this->GetAttribute("Sheets",1,AllocateInt(i +1))));
	}
	return true;
}


ExcelWorkbook::ExcelWorkbook(ExcelApp& myApp, VARIANT appObj){
	this->AppObj = appObj;
	this->parent = &myApp;
	this->isOpenned = true;
}

ExcelSheet* ExcelWorkbook::Sheets(int index){
	if(this->isOpenned && sheets.GetCount() > index){
		return &sheets[index];
	}
	return NULL;
}

ExcelSheet* ExcelWorkbook::Sheets(Upp::String name){
	if(this->isOpenned){
		for(int i = 0; i< sheets.GetCount(); i++){
			if (BSTRtoString(sheets[i].GetAttribute("Name").bstrVal).Compare(name) ==0){
				return &sheets[i];
			}
		}
	}
	return NULL;
}

ExcelSheet* ExcelWorkbook::AddSheet(){ //Create new Sheet with default Name
	if(this->isOpenned){
		return &sheets.Add(ExcelSheet(*this,GetAttribute(GetAttribute("Sheets"),"Add")));
	}
	return NULL;
}


ExcelSheet* ExcelWorkbook::AddSheet(Upp::String sheetName){ //Create new Sheet with defined name 
	if(this->isOpenned){
		ExcelSheet* mySheet;
		try{			
			mySheet = (&this->sheets.Add(ExcelSheet(*this,GetAttribute(GetAttribute("Sheets"),"Add"))));
			mySheet->SetName(sheetName);
			return mySheet;
		}catch(...){
			
		}
		return mySheet;
	}
	return NULL;
}



bool ExcelWorkbook::isReadOnly(){
	if(this->isOpenned){
		return (bool)GetAttribute("ReadOnly").lVal;
	}
	return false;
}

bool ExcelWorkbook::Save(){ //Save current workbook
	if(this->isOpenned){
		try{
			ExecuteMethode("Save",0);
			return true;
		}catch(...){
			 throw;
		}
	}
	return false;
}
bool ExcelWorkbook::SaveAs(Upp::String filePath){//Save current workbook at filePath
	if(this->isOpenned){
		try{
			ExecuteMethode("SaveAs",1,AllocateString(filePath));
			return true;
		}catch(...){
			 throw;
		}
	}
	return false;
}
bool ExcelWorkbook::Close(){//Close current workbook
	if(this->isOpenned){
		try{
			ExecuteMethode("Close",1,AllocateInt(0));
			return true;
		}catch(...){
			 throw;
		}
	}
	return false;
}

bool ExcelSheet::SetName(Upp::String sheetName){
	try{
		return SetAttribute(this->AppObj,"Name",sheetName);
	}catch(...){
		throw;	
	}
	return false;
}

ExcelWorkbook::~ExcelWorkbook(){

}

ExcelRange ExcelSheet::Range(Upp::String range){
	return ExcelRange(*this,this->ExecuteMethode("Range",1,AllocateString(range)));
}

ExcelRange ExcelSheet::Cells(int ligne, int colonne){
	char range[50];
	IndToStr(ligne,colonne,range);
	return this->Range(Upp::String(range));
}

int ExcelSheet::GetRowNumberOfMySheet(){
	return this->GetAttribute(this->GetAttribute("Rows"),"Count").intVal;
}

int ExcelSheet::GetLastRow(Upp::String Colonne){
	char range[10];
	ltoa(this->GetRowNumberOfMySheet(),range,10);
	Upp::String finalRange = Colonne + "1:"+ Colonne + Upp::String(range);
	/*
	Here we use some excel const
	xlDown		-4121	Down.
	xlToLeft	-4159	To left.
	xlToRight	-4161	To right.
	xlUp		-4162	Up.
	*/
	
	return this->GetAttribute(this->GetAttribute(this->GetAttribute("Range",1,AllocateString(finalRange)),"End",1,AllocateInt(-4121)),L"Row").intVal;	
}



ExcelSheet::ExcelSheet(ExcelWorkbook& parent, VARIANT appObj){
	this->AppObj = appObj;
	this->parent = &parent;	
}

ExcelSheet::~ExcelSheet(){
	
}

ExcelRange::ExcelRange(ExcelSheet &parent,VARIANT appObj){
	this->AppObj = appObj;
	this->parent = &parent;	
}
ExcelRange::~ExcelRange(){
	
}