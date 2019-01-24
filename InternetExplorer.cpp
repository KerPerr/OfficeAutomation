#include "InternetExplorer.h"

InternetExplorer::InternetExplorer(){
	this->isStarted=false;
	CoInitialize(NULL);
}

InternetExplorer::~InternetExplorer(){
	CoUninitialize();
}

bool InternetExplorer::Start() //Start new InternetExplorer Application
{
	if(!this->isStarted){
		this->AppObj = this->StartApp(WS_InternetExplorerApp);
		if( this->AppObj.intVal != -1){
			this->isStarted=true;
			return true;
		}
		return false;
	}
	return false;
}

bool InternetExplorer::FindApplication(){ //Find First current Excel Application still openned
	if(!this->isStarted){
		CLSID clsExcelApp;
		VARIANT xlApp = {0};

		if(FAILED(CLSIDFromProgID(WS_InternetExplorerApp, &clsExcelApp))) {
			this->isStarted=false;
			return false;
		}
		IUnknown *pUnk;
		HWND hExcelMainWnd = 0;
		hExcelMainWnd = FindWindow("TabThumbnailWindow",NULL);
		if(hExcelMainWnd) {
			SendMessage(hExcelMainWnd,WM_USER + 18, 0, 0);
			HRESULT hr2 = GetActiveObject(clsExcelApp,NULL,(IUnknown**)&pUnk);
			if (!FAILED(hr2)) {
			hr2=pUnk->QueryInterface(IID_IDispatch, (void **)&xlApp.pdispVal);
				if (!xlApp.ppdispVal) {
					this->isStarted=false;
					return false;
				}
			}
			if (pUnk) pUnk->Release();
		} else {
			this->isStarted=false;
			return false;
		}
		this->AppObj = xlApp;
		this->isStarted=true;
		return true;
	}
	return false;
}

bool InternetExplorer::Quit() //Close current InternetExplorer Application
{
	if(this->isStarted){
		try{
			this->ExecuteMethode("Quit",0);
			return true;
		}catch(...){
			return false;
		}
	}
	return false;
}

bool InternetExplorer::SetVisible(bool set)//Set or not the application visible
{
	if(this->isStarted){
		try{
			this->SetAttribute("Visible",(int)set);
			return true;
		}catch(...){
			return false;
		}
	}
	return false;
}

Upp::String InternetExplorer::GetCookie()
{
	VARIANT var;
	Upp::String prop = "Cookies";
	if(!this->isStarted) {
		this->ExecuteMethode("GetProperty", 2, AllocateString(prop), &var);
		return BSTRtoString(var.bstrVal);
	} else {
		return "Erreur : Application not running";
	}
}