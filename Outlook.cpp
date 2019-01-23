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
	
}

bool OutlookApp::Quit(){//Close current Outlook Application
	
}

bool OutlookApp::FindApplication(){ //Find First current Outlook Application still openned
	
}

bool OutlookApp::SetVisible(bool set){ //Set or not the application visible 
	
}

OutlookSession::OutlookSession(OutlookApp& parent, VARIANT appObj){
	
}

OutlookSession::~OutlookSession(){
	
}

