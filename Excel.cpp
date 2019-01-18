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
	this->AppObj = this->StartApp(WS_ExcelApp);
	if( this->AppObj.intVal != -1){
		this->ExcelIsStarted=true;
		return true;
	}
	return false;
}

bool ExcelApp::Quit() //Close current Excel Application
{
	try{
		this->ExecuteMethode("Quit",0);	
		return true;
	}catch(...){
		return false;
	}
}

bool ExcelApp::SetVisible(bool set)//Set or not the application visible 
{
	try{
		this->SetAttribute("Visible",(int)set);
		return true;
	}catch(...){
		return false;
	}
}

ExcelWorkbook::~ExcelWorkbook(){
	
}
ExcelSheet::~ExcelSheet(){
	
}
ExcelRange::~ExcelRange(){
	
}