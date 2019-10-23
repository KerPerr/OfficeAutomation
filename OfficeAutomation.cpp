#include <Core/Core.h>
#define _WIN32_WINNT 0x0501
#include <windows.h>
#include <ole2.h>
#include <stdio.h>
#include "OfficeAutomation.h"
#define DISP_FREEARGS 0x01
#define DISP_NOSHOWEXCEPTIONS 0x02

#define DLLFILENAME "C:\\Program Files (x86)\\Attachmate\\Reflection\\AgfSoft.dll"
#define DLIMODULE   LOGON
#define DLIHEADER   <OfficeAutomation/logon.dli>
#include <Core/dli.h>

using namespace Upp;

const int MAX = 255;
struct LogonInfoRecord{
	char users[MAX];
	int lgUser=MAX;
	char Pswd[MAX];
	int lgPsdw=MAX;
	char domaine[MAX];
	int lgDomaine=MAX;
};

//This function come from MSDN and have been Change By Clément Hamon
HRESULT Ole::AutoWrap(int autoType, VARIANT *pvResult, IDispatch *pDisp, LPOLESTR ptName, DISPPARAMS dp)
{
	HRESULT hr;
    if(!pDisp)
        {
       // MessageBox(NULL, "NULL IDispatch passed to AutoWrap()", "Error", 0x10010);
        throw OleException(0,"NULL IDispatch passed to AutoWrap()",0);
        _exit(0);
    }
    // Variables used...
    DISPID dispID;
    char buf[200];
    char szName[200];
    // Convert down to ANSI
    WideCharToMultiByte(CP_ACP, 0, ptName, -1, szName, 256, NULL, NULL);
    // Get DISPID for name passed...
    hr = pDisp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if(FAILED(hr))
        {
        sprintf(buf, "Action \"%s\" unreachable... err 0x%08lx", szName, hr);
        throw OleException(1,String(buf),0);
        _exit(0);
    }
    // Make the call!
    hr = pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, autoType, &dp, pvResult, NULL, NULL);
    if(FAILED(hr))
        {
              //  MessageBox(NULL, buf, "AutoWrap()", 0x10010);
                switch(hr)
                {
                case DISP_E_BADPARAMCOUNT:
                        sprintf(buf, "IDispatch::Invoke(\"%s\"=%08lx) failed w/err 0x%08lx : DISP_E_BADPARAMCOUNT",szName, dispID, hr);
                         MessageBox(NULL, buf, "AutoWrap()", 0x10010);
                        throw OleException(2,String(buf),0);
                        break;
                case DISP_E_BADVARTYPE:
                        sprintf(buf, "IDispatch::Invoke(\"%s\"=%08lx) failed w/err 0x%08lx : DISP_E_BADVARTYPE",szName, dispID, hr);
                         MessageBox(NULL, buf, "AutoWrap()", 0x10010);
                        throw OleException(3,String(buf),0);
                        break;
                case DISP_E_EXCEPTION:
                        sprintf(buf, "IDispatch::Invoke(\"%s\"=%08lx) failed w/err 0x%08lx : DISP_E_EXCEPTION",szName, dispID, hr);
                         MessageBox(NULL, buf, "AutoWrap()", 0x10010);
                        throw OleException(4,String(buf),0);
                        break;
                case DISP_E_MEMBERNOTFOUND:
                        sprintf(buf, "IDispatch::Invoke(\"%s\"=%08lx) failed w/err 0x%08lx : DISP_E_MEMBERNOTFOUND",szName, dispID, hr);
                         MessageBox(NULL, buf, "AutoWrap()", 0x10010);
                        throw OleException(5,String(buf),0);
                        break;
                case DISP_E_NONAMEDARGS:
                        sprintf(buf, "IDispatch::Invoke(\"%s\"=%08lx) failed w/err 0x%08lx : DISP_E_NONAMEDARGS",szName, dispID, hr);
                         MessageBox(NULL, buf, "AutoWrap()", 0x10010);
                        throw OleException(6,String(buf),0);
                        break;
                case DISP_E_OVERFLOW:
                        sprintf(buf, "IDispatch::Invoke(\"%s\"=%08lx) failed w/err 0x%08lx : DISP_E_OVERFLOW",szName, dispID, hr);
                         MessageBox(NULL, buf, "AutoWrap()", 0x10010);
                        throw OleException(7,String(buf),0);
                        break;
                case DISP_E_PARAMNOTFOUND:
                    	sprintf(buf, "IDispatch::Invoke(\"%s\"=%08lx) failed w/err 0x%08lx : DISP_E_PARAMNOTFOUND",szName, dispID, hr);
                    	 MessageBox(NULL, buf, "AutoWrap()", 0x10010);
                        throw OleException(8,String(buf),0);
                        break;
                case DISP_E_TYPEMISMATCH:
                    	sprintf(buf, "IDispatch::Invoke(\"%s\"=%08lx) failed w/err 0x%08lx : DISP_E_TYPEMISMATCH",szName, dispID, hr);
                    	 MessageBox(NULL, buf, "AutoWrap()", 0x10010);
                        throw OleException(9,String(buf),0);
                        break;
                case DISP_E_UNKNOWNINTERFACE:
                        sprintf(buf, "IDispatch::Invoke(\"%s\"=%08lx) failed w/err 0x%08lx : DISP_E_UNKNOWNINTERFACE",szName, dispID, hr);
                        throw OleException(10,String(buf),0);
                        break;
                case DISP_E_UNKNOWNLCID:
                    	sprintf(buf, "IDispatch::Invoke(\"%s\"=%08lx) failed w/err 0x%08lx : DISP_E_UNKNOWNLCID",szName, dispID, hr);
                        throw OleException(11,String(buf),0);
                        break;
                case DISP_E_PARAMNOTOPTIONAL:
                   		sprintf(buf, "IDispatch::Invoke(\"%s\"=%08lx) failed w/err 0x%08lx : DISP_E_PARAMNOTOPTIONAL",szName, dispID, hr);
                        throw OleException(12,String(buf),0);
                        break;
                }
                 MessageBox(NULL, buf, "AutoWrap()", 0x10010);
                return hr;
        }
    return hr;
}

//conversion BSTR to CHAR
Upp::String Ole::BSTRtoString (BSTR bstr)
{
    std::wstring ws(bstr);
    std::string str(ws.begin(), ws.end());
    return Upp::String(str);
}

//translating row and column number into the string name of the cell.
void Ole::IndToStr(int row,int col,char* strResult) {
	if(col>26) {
          sprintf(strResult,"%c%c%d\0",'A'+(col-1)/26-1,'A'+(col-1)%26,row);
    }
	else {
		sprintf(strResult,"%c%d\0",'A'+(col-1)%26,row);
    }
}
void Ole::InitSinkCommunication(const Upp::WString appName){
	/*	EventListened = true;
		eventListener = new Upp::Thread;

		eventListener->Run([this,appName](){
			CoInitialize(NULL);
			IID id;  
			HRESULT hr;
		    //COfficeEventHandler sink(this);
			IUnknown* iu;
			IConnectionPoint* pConnPoint;
			IConnectionPointContainer* pConnPntCont;
			
			CLSID clsApp;  
			hr = CLSIDFromProgID(appName, &clsApp);
			
			IUnknown* punk;
			hr = GetActiveObject( clsApp, NULL, &punk );

			hr = punk->QueryInterface(IID_IConnectionPointContainer, (void FAR* FAR*)&pConnPntCont); 
			hr = IIDFromString( this->CLSIDbyName(appName) ,&id); //(wchar_t *)~
		/*
			IEnumConnectionPoints* myEnum;
			hr = pConnPntCont->EnumConnectionPoints(&myEnum);
			
			ULONG itemRetrieved;
			LPCONNECTIONPOINT* myConPoint;
			hr = myEnum->Reset();
			hr = myEnum->Next((ULONG)2, myConPoint,&itemRetrieved);
			if(hr == 0){
				for(int e= 0; e < itemRetrieved;e++){
					IID theID;  
				   myConPoint[e]->GetConnectionInterface(&theID);
				
				}
			}
			
			hr = pConnPntCont->FindConnectionPoint( id, &pConnPoint );
			hr = sink.QueryInterface( IID_IUnknown, (void FAR* FAR*)&iu);
			pConnPoint->Advise( iu, &sink.m_dwEventCookie );
			MSG msg;
			BOOL  bRet;
			while( !Thread::IsShutdownThreads() ){
	 			GetMessage( &msg, NULL, 0, 0 );
	 			DispatchMessage(&msg);
			}	
			pConnPoint->Unadvise( sink.m_dwEventCookie );
			CoUninitialize();
		});*/
}

const Upp::WString Ole::CLSIDbyName(const Upp::WString appName) {
	if(appName.Compare(this->WS_ExcelApp)==0)
		return WS_CLSID_ExcelApp;
	else if(appName.Compare(this->WS_WordApp)==0)
		return WS_CLSID_WordApp;
	return WS_CLSID_ExcelApp;
}


VARIANT Ole::FindApp(const Upp::WString appName,bool startEventListener ,bool isFindOnly){
	CLSID clsApp;
	VARIANT App = {0};
	IUnknown* punk;
	HRESULT hr = CLSIDFromProgID(appName, &clsApp); 
	if(!FAILED(hr)){
		HRESULT hr2 =GetActiveObject( clsApp, NULL, &punk );
		if (!FAILED(hr2)) {
			hr2=punk->QueryInterface(IID_IDispatch, (void **)&App.pdispVal);
			if (!App.ppdispVal) {
				if(isFindOnly)
					App.intVal = -1;
				else 
					return this->StartApp(appName,startEventListener);
			}
		}else
		{
			if(isFindOnly)
				App.intVal = -1;
			else 
				return this->StartApp(appName,startEventListener);	
		}
	}
	else{
		if(isFindOnly)
			App.intVal = -1;
		else 
			return this->StartApp(appName,startEventListener);	
	}
	if(startEventListener && App.intVal != -1) {
		InitSinkCommunication(appName);
		this->EventListened = true;
	}
		
	return App;	
}



VARIANT Ole::StartApp(const Upp::WString appName,bool startEventListener ){
	CLSID clsApp;
	VARIANT App = {0}; //Variant who's contain the app, have -1 into App.intVal if something went wrong
	IUnknown* punk;
   /* Obtain the CLSID that identifies the app
   * This value is universally unique to Excel versions 5 and up, and
   * is used by OLE to identify which server to start.  We are obtaining
   * the CLSID from the ProgID.
   */
   if(FAILED(CLSIDFromProgID(appName, &clsApp))) {
      MessageBox(NULL, "CLSIDFromProgID() failed", "Error", 0x10010);
      throw OleException(13,"CLSIDFromProgID() => App Named " + appName.ToString() +" Can't be find",1);
   }
	if (FAILED(CoCreateInstance(clsApp, NULL, CLSCTX_SERVER,  IID_IUnknown, (void FAR* FAR*)&punk)))
	{
		MessageBox(NULL, "this App's not registered properly", "Error", 0x10010);
		throw OleException(14,"CoCreateInstance() => this App's ("+ appName.ToString()  +")not registered properly",1);
	}
	
	punk->QueryInterface(IID_IDispatch, (void **)&App.pdispVal);
	if(startEventListener) {
		InitSinkCommunication(appName);
		this->EventListened = true;
	}
	return App;
}

Ole::~Ole(){
//	EventListener->Detach();

}

VARIANT Ole::AllocateString(Upp::String myStr){
	VARIANT buffer = {0};
	VariantInit(&buffer);
	buffer.vt= VT_BSTR;
	buffer.bstrVal =SysAllocString((wchar_t*)~(myStr.ToWString()));
	return buffer;
}
VARIANT Ole::AllocateString(Upp::WString myStr){
	VARIANT buffer = {0};
	VariantInit(&buffer);
	buffer.vt= VT_BSTR;
	buffer.bstrVal =SysAllocString((wchar_t*)~(myStr));
	return buffer;
}
VARIANT Ole::AllocateInt(int myVar){
	VARIANT buffer = {0};
	VariantInit(&buffer);
	buffer.vt= VT_I4;
	buffer.lVal = myVar;
	return buffer;
}

VARIANT Ole::AllocateShort(short myVar){
	VARIANT buffer = {0};
	VariantInit(&buffer);
	buffer.vt= VT_I2;
	buffer.iVal = myVar;
	return buffer;
}

VARIANT Ole::AllocateFloat(float myVar){
	VARIANT buffer = {0};
	VariantInit(&buffer);
	buffer.vt= VT_R4;
	buffer.fltVal = myVar;
	return buffer;
}
VARIANT Ole::AllocateDouble(double myVar){
	VARIANT buffer = {0};
	VariantInit(&buffer);
	buffer.vt= VT_R8;
	buffer.dblVal = myVar;
	return buffer;
}
VARIANT Ole::AllocateLong(long myVar){
	VARIANT buffer = {0};
	VariantInit(&buffer);
	buffer.vt= VT_I8;
	buffer.iVal = myVar;
	return buffer;
}

VARIANT Ole::GetAttribute(Upp::WString attributeName) //Allow to retrieve attribute Value By VARIANT
{
	VARIANT *pArgs;
	try{
		VARIANT buffer={0};
	    DISPPARAMS dp = { NULL, NULL, 0, 0 };

	    // Allocate memory for arguments...
	    pArgs = new VARIANT[0];
	
	    // Build DISPPARAMS
	    dp.cArgs = 0;
	    dp.rgvarg = pArgs;

		AutoWrap(DISPATCH_PROPERTYGET,&buffer,this->AppObj.pdispVal,(wchar_t*)~attributeName,dp);
		delete [] pArgs;
		return buffer;
	}catch(const OleException & exception){
		delete [] pArgs;
		throw;
	}
}
		
bool Ole::SetAttribute(Upp::WString attributeName, Upp::String value)//Allow to set attribute Value
{
	VARIANT *pArgs;
	try{
		VARIANT buffer={0};
	    DISPPARAMS dp = { NULL, NULL, 0, 0 };
	    DISPID dispidNamed = DISPID_PROPERTYPUT;
	
	    // Allocate memory for arguments...
	    pArgs = new VARIANT[1];
	    pArgs[0] = AllocateString(value);
	
	    // Build DISPPARAMS
	    dp.cArgs = 1;
	    dp.rgvarg = pArgs;
        dp.cNamedArgs = 1;
        dp.rgdispidNamedArgs = &dispidNamed;
	    
		AutoWrap(DISPATCH_PROPERTYPUT,&buffer,this->AppObj.pdispVal,(wchar_t*)~attributeName,dp);
		delete [] pArgs;
		return true;
	}catch(const OleException & exception){
		delete [] pArgs;
		throw;
	}
}

bool Ole::SetAttribute(Upp::WString attributeName, int value)//Allow to set attribute Value
{
	VARIANT *pArgs;
	try{
		VARIANT buffer={0};
	    DISPPARAMS dp = { NULL, NULL, 0, 0 };
	    DISPID dispidNamed = DISPID_PROPERTYPUT;
	
	    // Allocate memory for arguments...
	    pArgs = new VARIANT[1];
	    pArgs[0] =AllocateInt(value);
	
	    // Build DISPPARAMS
	    dp.cArgs = 1;
	    dp.rgvarg = pArgs;
	    dp.cNamedArgs = 1;
	    dp.rgdispidNamedArgs = &dispidNamed;
	    
		AutoWrap(DISPATCH_PROPERTYPUT,&buffer,this->AppObj.pdispVal,(wchar_t*)~attributeName,dp);
		delete [] pArgs;
		return true;
	}catch(const OleException & exception){
		delete [] pArgs;
		throw;
	}
}
		
VARIANT Ole::ExecuteMethode(Upp::WString methodName,int cArgs...)//Allow to execute methode attribute retrieve VARIANT
{
	VARIANT *pArgs;
	try{
		va_list marker;
    	va_start(marker, cArgs);
		VARIANT buffer={0};
	    DISPPARAMS dp = { NULL, NULL, 0, 0 };
   
	    // Allocate memory for arguments...
	    pArgs = new VARIANT[cArgs+1];
	    // Extract arguments...
	    for(int i=0; i<cArgs; i++)
	    {
	        pArgs[i] = va_arg(marker, VARIANT);
	    }
	    va_end(marker);
	    // Build DISPPARAMS
	    dp.cArgs = cArgs;
	    dp.rgvarg = pArgs;
		AutoWrap(DISPATCH_METHOD,&buffer,this->AppObj.pdispVal,(wchar_t*)~methodName,dp);
		delete [] pArgs;
		return buffer;
	}catch(OleException const& exception){
		delete [] pArgs;
		throw;
	}
}


VARIANT Ole::GetAttribute(VARIANT variant,Upp::WString attributeName) //Allow to retrieve attribute Value By VARIANT
{
	VARIANT *pArgs;
	try{
		VARIANT buffer={0};
	    DISPPARAMS dp = { NULL, NULL, 0, 0 };

	    // Allocate memory for arguments...
	    pArgs = new VARIANT[0];
	
	    // Build DISPPARAMS
	    dp.cArgs = 0;
	    dp.rgvarg = pArgs;

		AutoWrap(DISPATCH_PROPERTYGET,&buffer,variant.pdispVal,(wchar_t*)~attributeName,dp);
		delete [] pArgs;
		return buffer;
	}catch(const OleException & exception){
		delete [] pArgs;
		throw;
	}
}
VARIANT Ole::GetAttribute(IDispatch* disp,Upp::WString attributeName) //Allow to retrieve attribute Value By VARIANT
{
	VARIANT *pArgs;
	try{
		VARIANT buffer={0};
	    DISPPARAMS dp = { NULL, NULL, 0, 0 };

	    // Allocate memory for arguments...
	    pArgs = new VARIANT[0];
	
	    // Build DISPPARAMS
	    dp.cArgs = 0;
	    dp.rgvarg = pArgs;

		AutoWrap(DISPATCH_PROPERTYGET,&buffer,disp,(wchar_t*)~attributeName,dp);
		delete [] pArgs;
		return buffer;
	}catch(const OleException & exception){
		delete [] pArgs;
		throw;
	}
}

		
bool Ole::SetAttribute(VARIANT variant,Upp::WString attributeName, Upp::String value)//Allow to set attribute Value
{
	VARIANT *pArgs;
	try{
		VARIANT buffer={0};
	    DISPPARAMS dp = { NULL, NULL, 0, 0 };
	    DISPID dispidNamed = DISPID_PROPERTYPUT;
	
	    // Allocate memory for arguments...
	    pArgs = new VARIANT[1];
	    pArgs[0] = AllocateString(value);
	
	    // Build DISPPARAMS
	    dp.cArgs = 1;
	    dp.rgvarg = pArgs;
        dp.cNamedArgs = 1;
        dp.rgdispidNamedArgs = &dispidNamed;
	    
		AutoWrap(DISPATCH_PROPERTYPUT,&buffer,variant.pdispVal,(wchar_t*)~attributeName,dp);
		delete [] pArgs;
		return true;
	}catch(const OleException & exception){
		delete [] pArgs;
		throw;
	}
}

bool Ole::SetAttribute(IDispatch* disp,Upp::WString attributeName, Upp::String value)//Allow to set attribute Value
{
	VARIANT *pArgs;
	try{
		VARIANT buffer={0};
	    DISPPARAMS dp = { NULL, NULL, 0, 0 };
	    DISPID dispidNamed = DISPID_PROPERTYPUT;
	
	    // Allocate memory for arguments...
	    pArgs = new VARIANT[1];
	    pArgs[0] = AllocateString(value);
	
	    // Build DISPPARAMS
	    dp.cArgs = 1;
	    dp.rgvarg = pArgs;
        dp.cNamedArgs = 1;
        dp.rgdispidNamedArgs = &dispidNamed;
	    
		AutoWrap(DISPATCH_PROPERTYPUT,&buffer,disp,(wchar_t*)~attributeName,dp);
		delete [] pArgs;
		return true;
	}catch(const OleException & exception){
		delete [] pArgs;
		throw;
	}
}

bool Ole::SetAttribute(VARIANT variant,Upp::WString attributeName, int value)//Allow to set attribute Value
{
	VARIANT *pArgs;
	try{
		VARIANT buffer={0};
	    DISPPARAMS dp = { NULL, NULL, 0, 0 };
	    DISPID dispidNamed = DISPID_PROPERTYPUT;
	
	    // Allocate memory for arguments...
	    pArgs = new VARIANT[1];
	    pArgs[0] =AllocateInt(value);
	
	    // Build DISPPARAMS
	    dp.cArgs = 1;
	    dp.rgvarg = pArgs;
        dp.cNamedArgs = 1;
        dp.rgdispidNamedArgs = &dispidNamed;
        
		AutoWrap(DISPATCH_PROPERTYPUT,&buffer,variant.pdispVal,(wchar_t*)~attributeName,dp);
		delete [] pArgs;
		return true;
	}catch(const OleException & exception){
		delete [] pArgs;
		throw;
	}
}
bool Ole::SetAttribute(IDispatch* disp,Upp::WString attributeName, int value)//Allow to set attribute Value
{
	VARIANT *pArgs;
	try{
		VARIANT buffer={0};
	    DISPPARAMS dp = { NULL, NULL, 0, 0 };
	    DISPID dispidNamed = DISPID_PROPERTYPUT;
	
	    // Allocate memory for arguments...
	    pArgs = new VARIANT[1];
	    pArgs[0] =AllocateInt(value);
	
	    // Build DISPPARAMS
	    dp.cArgs = 1;
	    dp.rgvarg = pArgs;
        dp.cNamedArgs = 1;
        dp.rgdispidNamedArgs = &dispidNamed;
        
		AutoWrap(DISPATCH_PROPERTYPUT,&buffer,disp,(wchar_t*)~attributeName,dp);
		delete [] pArgs;
		return true;
	}catch(const OleException & exception){
		delete [] pArgs;
		throw;
	}
}

		
VARIANT Ole::ExecuteMethode(VARIANT variant,Upp::WString methodName,int cArgs...)//Allow to execute methode attribute retrieve VARIANT
{
	VARIANT *pArgs;
	try{
		va_list marker;
    	va_start(marker, cArgs);
		VARIANT buffer={0};
	    DISPPARAMS dp = { NULL, NULL, 0, 0 };

	    // Allocate memory for arguments...
	    pArgs = new VARIANT[cArgs+1];
	    // Extract arguments...
	    for(int i=0; i<cArgs; i++)
	    {
	        pArgs[i] = va_arg(marker, VARIANT);
	    }
	    va_end(marker);
	    // Build DISPPARAMS
	    dp.cArgs = cArgs;
	    dp.rgvarg = pArgs;
		AutoWrap(DISPATCH_METHOD,&buffer,variant.pdispVal,(wchar_t*)~methodName,dp);
		delete [] pArgs;
		return buffer;
	}catch(const OleException & exception){
		delete [] pArgs;
		throw;
	}
}

VARIANT Ole::ExecuteMethode(IDispatch* disp,Upp::WString methodName,int cArgs...)//Allow to execute methode attribute retrieve VARIANT
{
	VARIANT *pArgs;
	try{
		va_list marker;
    	va_start(marker, cArgs);
		VARIANT buffer={0};
	    DISPPARAMS dp = { NULL, NULL, 0, 0 };

	    // Allocate memory for arguments...
	    pArgs = new VARIANT[cArgs+1];
	    // Extract arguments...
	    for(int i=0; i<cArgs; i++)
	    {
	        pArgs[i] = va_arg(marker, VARIANT);
	    }
	    va_end(marker);
	    // Build DISPPARAMS
	    dp.cArgs = cArgs;
	    dp.rgvarg = pArgs;
		AutoWrap(DISPATCH_METHOD,&buffer,disp,(wchar_t*)~methodName,dp);
		delete [] pArgs;
		return buffer;
	}catch(const OleException & exception){
		delete [] pArgs;
		throw;
	}
}

VARIANT Ole::GetAttribute(Upp::WString attributeName,int cArgs...){
	VARIANT *pArgs;
	try{
		va_list marker;
    	va_start(marker, cArgs);
		VARIANT buffer={0};
	    DISPPARAMS dp = { NULL, NULL, 0, 0 };
	    // Allocate memory for arguments...
	    pArgs = new VARIANT[cArgs+1];
	    // Extract arguments...
	    for(int i=0; i<cArgs; i++)
	    {
	        pArgs[i] = va_arg(marker, VARIANT);
	    }
	    va_end(marker);
	    // Build DISPPARAMS
	    dp.cArgs = cArgs;
	    dp.rgvarg = pArgs;
		AutoWrap(DISPATCH_PROPERTYGET|DISPATCH_METHOD,&buffer,AppObj.pdispVal,(wchar_t*)~attributeName,dp);
		delete [] pArgs;
		return buffer;
	}catch(const OleException & exception){
		delete [] pArgs;
		throw;
	}
}
VARIANT Ole::GetAttribute(VARIANT variant,Upp::WString attributeName,int cArgs...){
	VARIANT *pArgs;
  	try{
  		va_list marker;
    	va_start(marker, cArgs);
  		// Variables used...
  		VARIANT buffer={0};
	    DISPPARAMS dp = { NULL, NULL, 0, 0 };
	    // Allocate memory for arguments...
	    pArgs = new VARIANT[cArgs+1];
	    // Extract arguments...
	    for(int i=0; i<cArgs; i++)
	    {
	        pArgs[i] = va_arg(marker, VARIANT);
	    }
	    va_end(marker);
	    // Build DISPPARAMS
	    dp.cArgs = cArgs;
	    dp.rgvarg = pArgs;
		AutoWrap(DISPATCH_PROPERTYGET|DISPATCH_METHOD,&buffer,variant.pdispVal,(wchar_t*)~attributeName,dp);
		delete [] pArgs;
		return buffer;
	}catch(const OleException & exception){
		delete [] pArgs;
		throw;
	}
}
VARIANT Ole::GetAttribute(IDispatch* disp,Upp::WString attributeName,int cArgs...){
	VARIANT *pArgs;
  	try{
  		va_list marker;
    	va_start(marker, cArgs);
  		// Variables used...
  		VARIANT buffer={0};
	    DISPPARAMS dp = { NULL, NULL, 0, 0 };
	    // Allocate memory for arguments...
	    pArgs = new VARIANT[cArgs+1];
	    // Extract arguments...
	    for(int i=0; i<cArgs; i++)
	    {
	        pArgs[i] = va_arg(marker, VARIANT);
	    }
	    va_end(marker);
	    // Build DISPPARAMS
	    dp.cArgs = cArgs;
	    dp.rgvarg = pArgs;
		AutoWrap(DISPATCH_PROPERTYGET|DISPATCH_METHOD,&buffer,disp,(wchar_t*)~attributeName,dp);
		delete [] pArgs;
		return buffer;
	}catch(const OleException & exception){
		delete [] pArgs;
		throw;
	}
}


int Ole::ColStrToInt(Upp::String target){
	int resultat= 0;
	for(int i = 0; i < target.GetCount(); i++){
		if((int)toupper(target[i]) >64 && (int)toupper(target[i]) < 91){
			if (i>0) {
				resultat+= 26 *((int)toupper(target[i]) -64);
			}
			else
			{
				resultat+= ((int)toupper(target[i]) -64);
			}
		}
	}
	return resultat;
}

int Ole::ExtractRow(Upp::String target)
{
	char myRow[target.GetCount()];
	int iterator = 0;
	for(int i = 0; i < target.GetCount(); i++) {
		if( int(target[i]) >47 && int(target[i]) < 58){
			myRow[iterator] = target[i];
			iterator++;
		}
	}
	return atoi(myRow);
}

Upp::String Ole::StringWOZ(Upp::String str) {
	str.Trim(str.Find('.'));
	return str;
}
