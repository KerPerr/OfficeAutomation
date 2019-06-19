#include <Core/Core.h>
#include <windows.h>
#include <ole2.h>
#include <stdio.h>
#include "Outlook.h"
#include "OfficeAutomation.h"

#define DISP_FREEARGS 0x01
#define DISP_NOSHOWEXCEPTIONS 0x02

using namespace Upp;

OutlookApp::OutlookApp(){ //Initialise COM
	CoInitialize(NULL);
}

OutlookApp::~OutlookApp(){ //Unitialise COM
//	~Ole();
	CoUninitialize();
}


MailItem OutlookApp::CreateMail(){
	if(this->OutlookIsStarted){
		try{
			return MailItem(*this, this->ExecuteMethode("CreateItem",1,AllocateInt(0)));
		}catch(OleException const& exception){
			throw exception;
		}
	}
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
		this->AppObj = this->FindApp(WS_OutlookApp, false);
		if( this->AppObj.intVal != -1) {
			this->OutlookIsStarted=true;
			return true;
		}else{
			return Start();	
		}
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
		if(!this->OutlookIsStarted){
		CLSID clsExcelApp;
		VARIANT xlApp = {0};

	   if(FAILED(CLSIDFromProgID(WS_OutlookApp, &clsExcelApp))) {
	      this->OutlookIsStarted=false;
	      return false;
	   }
	   IUnknown *pUnk;
	   HWND hExcelMainWnd = 0;
	   hExcelMainWnd = FindWindow("OLMAIN",NULL);
	   if(hExcelMainWnd) {
		   SendMessage(hExcelMainWnd,WM_USER + 18, 0, 0);
			HRESULT hr2 = GetActiveObject(clsExcelApp,NULL,(IUnknown**)&pUnk);
			if (!FAILED(hr2)) {
				hr2=pUnk->QueryInterface(IID_IDispatch, (void **)&xlApp.pdispVal);
				if (!xlApp.ppdispVal) {
					this->OutlookIsStarted=false;
					return false;
				}
			}
			if (pUnk) pUnk->Release();
		}
		else {
			this->OutlookIsStarted=false;
			return false;
		}
		this->AppObj = xlApp;
		this->OutlookIsStarted=true;
		return true;
			
	}
	return false;
}

bool OutlookApp::SetVisible(bool set){ //Set or not the application visible 
	if(this->OutlookIsStarted){
		try{
			this->SetAttribute("Visible",(int)set);
			return true;
		}catch(OleException const& exception){
			throw exception;
		}
	}
	return false;
}


bool MailItem::AddRecever(Upp::String email){
	try{
		auto recip = this->GetAttribute("Recipients");
		auto Added = ExecuteMethode(recip,"Add",1, AllocateString(email));
	    SetAttribute(Added ,L"Type",1 );
	    return true;
	}catch(OleException const& exception){
		throw exception;
	}
	return false;
}

bool MailItem::AddRecipients(Upp::String email){
	try{
		//2 correspond to  olCC constant indicate the type of recipients added (here cc )
	    SetAttribute(ExecuteMethode(this->GetAttribute("Recipients"),"Add",1, AllocateString(email)) ,L"Type"  ,2  );
	}catch(OleException const& exception){
		throw exception;
	}
	return false;
}

bool MailItem::SetSubject(Upp::String subject){
	try{
	    this->SetAttribute("subject",subject);
	}catch(OleException const& exception){
		throw exception;
	}
	return false;
}

bool MailItem::SetBody(Upp::String body){
	try{
	    this->SetAttribute("Body",body);
	}catch(OleException const& exception){
		throw exception;
	}
	return false;
}

bool MailItem::SetHTMLBody(Upp::String body){
	try{
	    this->SetAttribute("HTMLBody",body);
	}catch(OleException const& exception){
		throw exception;
	}
	return false;
}

bool MailItem::setHighImportance(){
	try{
	    this->SetAttribute("Importance",2);
	}catch(OleException const& exception){
		throw exception;
	}
	return false;
}

bool MailItem::DisplayMail(){
	try{
	    this->ExecuteMethode("Display",0);
	}catch(OleException const& exception){
		throw exception;
	}
	return false;
}
bool MailItem::AddItem(String PathToData){
	try{
		if(Upp::FileExists(PathToData)){
			auto recip = this->GetAttribute("Attachments");
			auto Added = ExecuteMethode(recip,"Add",1, AllocateString(PathToData));
		    return true;
		}
		Cout()<<"Fichier introuvable !"<<"\n";
		return false;
	}catch(OleException const& exception){
		throw exception;
	}
	return false;
}


MailItem::~MailItem(){
	
}

MailItem::MailItem(OutlookApp& parent, VARIANT appObj){
	this->AppObj = appObj;
	this->parent = &parent;
} 



