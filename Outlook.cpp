#include <Core/Core.h>
#include <windows.h>
#include <ole2.h>
#include <stdio.h>
#include "Outlook.h"
#include "OfficeAutomation.h"

#define DISP_FREEARGS 0x01
#define DISP_NOSHOWEXCEPTIONS 0x02



OutlookApp::OutlookApp(){ //Initialise COM
	CoInitialize(NULL);
}

OutlookApp::~OutlookApp(){ //Unitialise COM
	CoUninitialize();
}

OutlookSession* OutlookApp::GetSession(){
	return this->session;
}

bool OutlookApp::Start(){ //Start new Outlook Application
		if(!this->OutlookIsStarted){
		this->AppObj = this->StartApp(WS_OutlookApp);
		if( this->AppObj.intVal != -1){
			this->OutlookIsStarted=true;
			return true;
		}
		return false;
	}
	return false;
}

bool OutlookApp::FindOrStart(){ //Find running Outlook or Start new One
	if(!this->OutlookIsStarted){
		CLSID clsOutlookApp;
		VARIANT xlApp = {0};

		if(FAILED(CLSIDFromProgID(WS_ExcelApp, &clsOutlookApp))) {
		  this->OutlookIsStarted=false;
		  return this->Start();
		}
	   IUnknown *pUnk;
	   HWND hExcelMainWnd = 0;
	   hExcelMainWnd = FindWindow("OLMAIN",NULL);
	   if(hExcelMainWnd) {
		   SendMessage(hExcelMainWnd,WM_USER + 18, 0, 0);
			HRESULT hr2 = GetActiveObject(clsOutlookApp,NULL,(IUnknown**)&pUnk);
			if (!FAILED(hr2)) {
				hr2=pUnk->QueryInterface(IID_IDispatch, (void **)&xlApp.pdispVal);
				if (!xlApp.ppdispVal) {
					this->OutlookIsStarted=false;
					return this->Start();
				}
			}
			if (pUnk) pUnk->Release();
		}
		else {
			this->OutlookIsStarted=false;
			return this->Start();
		}
		this->AppObj = xlApp;
		this->OutlookIsStarted=true;
		return true;
			
	}
	return false;
}

bool OutlookApp::Quit(){//Close current Outlook Application
	if(this->OutlookIsStarted){
		try{
			this->ExecuteMethode("Quit",0);	
			return true;
		}catch(OleException const& exception){
			throw exception;
		}
	}
	return false;
}

bool OutlookApp::FindApplication(){ //Find First current Outlook Application still openned
	
}

bool OutlookApp::SetVisible(bool set){ //Set or not the application visible 
	
}

OutlookSession::OutlookSession(OutlookApp& parent, VARIANT appObj){
	
}

OutlookSession::~OutlookSession(){
	
}

