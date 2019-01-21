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

bool ExcelApp::NewWorkbook(){
	if(this->ExcelIsStarted){
		GetAttribute(GetAttribute("Workbooks"),"Add");
		return true;
	}
	return false;
}

ExcelWorkbook::~ExcelWorkbook(){
	
}
ExcelSheet::~ExcelSheet(){
	
}
ExcelRange::~ExcelRange(){
	
}