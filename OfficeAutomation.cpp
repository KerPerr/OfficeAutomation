#include <Core/Core.h>

#define _WIN32_WINNT 0x0501
#include <windows.h>
#include <ole2.h>
#include <stdio.h>
#include "OfficeAutomation.h"

#define DISP_FREEARGS 0x01
#define DISP_NOSHOWEXCEPTIONS 0x02

using namespace Upp;
/*
//Fonction reprise de MSDN
HRESULT Ole::AutoWrap(int autoType, VARIANT *pvResult, IDispatch *pDisp, LPOLESTR ptName, int cArgs...)
{
    // Begin variable-argument list...
    va_list marker;
    va_start(marker, cArgs);

    if(!pDisp)
        {
        MessageBox(NULL, "NULL IDispatch passed to AutoWrap()", "Error", 0x10010);
        _exit(0);
    }

    // Variables used...
    DISPPARAMS dp = { NULL, NULL, 0, 0 };
    DISPID dispidNamed = DISPID_PROPERTYPUT;
    DISPID dispID;
    HRESULT hr;
    char buf[200];
    char szName[200];

    // Convert down to ANSI
    WideCharToMultiByte(CP_ACP, 0, ptName, -1, szName, 256, NULL, NULL);

    // Get DISPID for name passed...
    hr = pDisp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if(FAILED(hr))
        {
        sprintf(buf, "IDispatch::GetIDsOfNames(\"%s\") failed w/err 0x%08lx", szName, hr);
        MessageBox(NULL, buf, "AutoWrap()", 0x10010);
      	// throw 1;
        //_exit(0);
        return hr;
    }

    // Allocate memory for arguments...
    VARIANT *pArgs = new VARIANT[cArgs+1];
    // Extract arguments...
    for(int i=0; i<cArgs; i++)
        {
        pArgs[i] = va_arg(marker, VARIANT);
    }

    // Build DISPPARAMS
    dp.cArgs = cArgs;
    dp.rgvarg = pArgs;

    // Handle special-case for property-puts!
    if(autoType & DISPATCH_PROPERTYPUT)
        {
        dp.cNamedArgs = 1;
        dp.rgdispidNamedArgs = &dispidNamed;
    }

    // Make the call!
    hr = pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, autoType, &dp, pvResult, NULL, NULL);
    if(FAILED(hr))
        {
                sprintf(buf, "IDispatch::Invoke(\"%s\"=%08lx) failed w/err 0x%08lx",
                        szName, dispID, hr);
                       // throw 1;
                MessageBox(NULL, buf, "AutoWrap()", 0x10010);
                switch(hr)
                {
                case DISP_E_BADPARAMCOUNT:
                     //	throw 1;
                        MessageBox(NULL, "DISP_E_BADPARAMCOUNT", "Error:", 0x10010);
                        break;
                case DISP_E_BADVARTYPE:
                    //	 throw 1;
                        MessageBox(NULL, "DISP_E_BADVARTYPE", "Error:", 0x10010);
                        break;
                case DISP_E_EXCEPTION:
                    //	 throw 1;
                        MessageBox(NULL, "DISP_E_EXCEPTION", "Error:", 0x10010);
                        break;
                case DISP_E_MEMBERNOTFOUND:
                     //	throw 1;
                        MessageBox(NULL, "DISP_E_MEMBERNOTFOUND", "Error:", 0x10010);
                        break;
                case DISP_E_NONAMEDARGS:
                    	// throw 1;
                        MessageBox(NULL, "DISP_E_NONAMEDARGS", "Error:", 0x10010);
                        break;
                case DISP_E_OVERFLOW:
                     //	throw 1;
                        MessageBox(NULL, "DISP_E_OVERFLOW", "Error:", 0x10010);
                        break;
                case DISP_E_PARAMNOTFOUND:
                    	// throw 1;
                        MessageBox(NULL, "DISP_E_PARAMNOTFOUND", "Error:", 0x10010);
                        break;
                case DISP_E_TYPEMISMATCH:
                    	// throw 1;
                        MessageBox(NULL, "DISP_E_TYPEMISMATCH", "Error:", 0x10010);
                        break;
                case DISP_E_UNKNOWNINTERFACE:
                     //	throw 1;
                        MessageBox(NULL, "DISP_E_UNKNOWNINTERFACE", "Error:", 0x10010);
                        break;
                case DISP_E_UNKNOWNLCID:
                    //	 throw 1;
                        MessageBox(NULL, "DISP_E_UNKNOWNLCID", "Error:", 0x10010);
                        break;
                case DISP_E_PARAMNOTOPTIONAL:
                    //	 throw 1;
                        MessageBox(NULL, "DISP_E_PARAMNOTOPTIONAL", "Error:", 0x10010);
                        break;
                }
                // _exit(0);
                return hr;
        }
    // End variable-argument section...
    va_end(marker);

    delete [] pArgs;

    return hr;
}*/


//This function come from MSDN and have been Change By Clément Hamon
HRESULT Ole::AutoWrap(int autoType, VARIANT *pvResult, IDispatch *pDisp, LPOLESTR ptName, DISPPARAMS dp)
{
    if(!pDisp)
        {
        MessageBox(NULL, "NULL IDispatch passed to AutoWrap()", "Error", 0x10010);
        _exit(0);
    }

    // Variables used...
    DISPID dispID;
    HRESULT hr;
    char buf[200];
    char szName[200];

    // Convert down to ANSI
    WideCharToMultiByte(CP_ACP, 0, ptName, -1, szName, 256, NULL, NULL);

    // Get DISPID for name passed...
    hr = pDisp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if(FAILED(hr))
        {
        sprintf(buf, "IDispatch::GetIDsOfNames(\"%s\") failed w/err 0x%08lx", szName, hr);
        MessageBox(NULL, buf, "AutoWrap()", 0x10010);
      	// throw 1;
        //_exit(0);
        return hr;
    }

    // Make the call!
    hr = pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, autoType, &dp, pvResult, NULL, NULL);
    if(FAILED(hr))
        {
                sprintf(buf, "IDispatch::Invoke(\"%s\"=%08lx) failed w/err 0x%08lx",
                        szName, dispID, hr);
                       // throw 1;
                MessageBox(NULL, buf, "AutoWrap()", 0x10010);
                switch(hr)
                {
                case DISP_E_BADPARAMCOUNT:
                     //	throw 1;
                        MessageBox(NULL, "DISP_E_BADPARAMCOUNT", "Error:", 0x10010);
                        break;
                case DISP_E_BADVARTYPE:
                    //	 throw 1;
                        MessageBox(NULL, "DISP_E_BADVARTYPE", "Error:", 0x10010);
                        break;
                case DISP_E_EXCEPTION:
                    //	 throw 1;
                        MessageBox(NULL, "DISP_E_EXCEPTION", "Error:", 0x10010);
                        break;
                case DISP_E_MEMBERNOTFOUND:
                     //	throw 1;
                        MessageBox(NULL, "DISP_E_MEMBERNOTFOUND", "Error:", 0x10010);
                        break;
                case DISP_E_NONAMEDARGS:
                    	// throw 1;
                        MessageBox(NULL, "DISP_E_NONAMEDARGS", "Error:", 0x10010);
                        break;
                case DISP_E_OVERFLOW:
                     //	throw 1;
                        MessageBox(NULL, "DISP_E_OVERFLOW", "Error:", 0x10010);
                        break;
                case DISP_E_PARAMNOTFOUND:
                    	// throw 1;
                        MessageBox(NULL, "DISP_E_PARAMNOTFOUND", "Error:", 0x10010);
                        break;
                case DISP_E_TYPEMISMATCH:
                    	// throw 1;
                        MessageBox(NULL, "DISP_E_TYPEMISMATCH", "Error:", 0x10010);
                        break;
                case DISP_E_UNKNOWNINTERFACE:
                     //	throw 1;
                        MessageBox(NULL, "DISP_E_UNKNOWNINTERFACE", "Error:", 0x10010);
                        break;
                case DISP_E_UNKNOWNLCID:
                    //	 throw 1;
                        MessageBox(NULL, "DISP_E_UNKNOWNLCID", "Error:", 0x10010);
                        break;
                case DISP_E_PARAMNOTOPTIONAL:
                    //	 throw 1;
                        MessageBox(NULL, "DISP_E_PARAMNOTOPTIONAL", "Error:", 0x10010);
                        break;
                }
                // _exit(0);
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

VARIANT Ole::StartApp(const Upp::WString appName){
	CLSID clsApp;
	VARIANT App = {0}; //Variant who's contain the app, have -1 into App.intVal if something went wrong
	
   /* Obtain the CLSID that identifies the app
   * This value is universally unique to Excel versions 5 and up, and
   * is used by OLE to identify which server to start.  We are obtaining
   * the CLSID from the ProgID.
   */
   if(FAILED(CLSIDFromProgID(appName, &clsApp))) {
      MessageBox(NULL, "CLSIDFromProgID() failed", "Error", 0x10010);
      App.intVal =-1;
      return App;
   }	
	
	if (FAILED(CoCreateInstance(clsApp, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void **)&App.pdispVal)))
	{
		MessageBox(NULL, "this App's not registered properly", "Error", 0x10010);
		App.intVal =-1;
		return App;
	}
	
	return App;
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

VARIANT Ole::GetAttribute(Upp::WString attributeName) //Allow to retrieve attribute Value By VARIANT
{
	try{
		VARIANT buffer={0};
	    DISPPARAMS dp = { NULL, NULL, 0, 0 };

	    // Allocate memory for arguments...
	    VARIANT *pArgs = new VARIANT[0];
	
	    // Build DISPPARAMS
	    dp.cArgs = 0;
	    dp.rgvarg = pArgs;

		AutoWrap(DISPATCH_PROPERTYGET,&buffer,this->AppObj.pdispVal,(wchar_t*)~attributeName,dp);
		delete [] pArgs;
		return buffer;
	}catch(...){
		throw;	
	}
}
		
bool Ole::SetAttribute(Upp::WString attributeName, Upp::String value)//Allow to set attribute Value
{
	try{
		VARIANT buffer={0};
	    DISPPARAMS dp = { NULL, NULL, 0, 0 };
	    DISPID dispidNamed = DISPID_PROPERTYPUT;
	
	    // Allocate memory for arguments...
	    VARIANT *pArgs = new VARIANT[1];
	    pArgs[0] = AllocateString(value);
	
	    // Build DISPPARAMS
	    dp.cArgs = 1;
	    dp.rgvarg = pArgs;
        dp.cNamedArgs = 1;
        dp.rgdispidNamedArgs = &dispidNamed;
	    
		AutoWrap(DISPATCH_PROPERTYPUT,&buffer,this->AppObj.pdispVal,(wchar_t*)~attributeName,dp);
		delete [] pArgs;
		return true;
	}catch(...){
		throw;	
	}
}

bool Ole::SetAttribute(Upp::WString attributeName, int value)//Allow to set attribute Value
{
	try{
		VARIANT buffer={0};
	    DISPPARAMS dp = { NULL, NULL, 0, 0 };
	    DISPID dispidNamed = DISPID_PROPERTYPUT;
	
	    // Allocate memory for arguments...
	    VARIANT *pArgs = new VARIANT[1];
	    pArgs[0] =AllocateInt(value);
	
	    // Build DISPPARAMS
	    dp.cArgs = 1;
	    dp.rgvarg = pArgs;
	    dp.cNamedArgs = 1;
	    dp.rgdispidNamedArgs = &dispidNamed;
	    
		AutoWrap(DISPATCH_PROPERTYPUT,&buffer,this->AppObj.pdispVal,(wchar_t*)~attributeName,dp);
		delete [] pArgs;
		return true;
	}catch(...){
		throw;	
	}
}
		
VARIANT Ole::ExecuteMethode(Upp::WString methodName,int cArgs...)//Allow to execute methode attribute retrieve VARIANT
{
	try{
		va_list marker;
    	va_start(marker, cArgs);
		VARIANT buffer={0};
	    DISPPARAMS dp = { NULL, NULL, 0, 0 };
   
	    // Allocate memory for arguments...
	    VARIANT *pArgs = new VARIANT[cArgs+1];
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
	}catch(...){
		throw;
	}
}


VARIANT Ole::GetAttribute(VARIANT variant,Upp::WString attributeName) //Allow to retrieve attribute Value By VARIANT
{
	try{
		VARIANT buffer={0};
	    DISPPARAMS dp = { NULL, NULL, 0, 0 };

	    // Allocate memory for arguments...
	    VARIANT *pArgs = new VARIANT[0];
	
	    // Build DISPPARAMS
	    dp.cArgs = 0;
	    dp.rgvarg = pArgs;

		AutoWrap(DISPATCH_PROPERTYGET,&buffer,variant.pdispVal,(wchar_t*)~attributeName,dp);
		delete [] pArgs;
		return buffer;
	}catch(...){
		throw;	
	}
}
		
bool Ole::SetAttribute(VARIANT variant,Upp::WString attributeName, Upp::String value)//Allow to set attribute Value
{
	try{
		VARIANT buffer={0};
	    DISPPARAMS dp = { NULL, NULL, 0, 0 };
	    DISPID dispidNamed = DISPID_PROPERTYPUT;
	
	    // Allocate memory for arguments...
	    VARIANT *pArgs = new VARIANT[1];
	    pArgs[0] = AllocateString(value);
	
	    // Build DISPPARAMS
	    dp.cArgs = 1;
	    dp.rgvarg = pArgs;
        dp.cNamedArgs = 1;
        dp.rgdispidNamedArgs = &dispidNamed;
	    
		AutoWrap(DISPATCH_PROPERTYPUT,&buffer,variant.pdispVal,(wchar_t*)~attributeName,dp);
		delete [] pArgs;
		return true;
	}catch(...){
		throw;	
	}
}
bool Ole::SetAttribute(VARIANT variant,Upp::WString attributeName, int value)//Allow to set attribute Value
{
	try{
		VARIANT buffer={0};
	    DISPPARAMS dp = { NULL, NULL, 0, 0 };
	    DISPID dispidNamed = DISPID_PROPERTYPUT;
	
	    // Allocate memory for arguments...
	    VARIANT *pArgs = new VARIANT[1];
	    pArgs[0] =AllocateInt(value);
	
	    // Build DISPPARAMS
	    dp.cArgs = 1;
	    dp.rgvarg = pArgs;
        dp.cNamedArgs = 1;
        dp.rgdispidNamedArgs = &dispidNamed;
        
		AutoWrap(DISPATCH_PROPERTYPUT,&buffer,variant.pdispVal,(wchar_t*)~attributeName,dp);
		delete [] pArgs;
		return true;
	}catch(...){
		throw;	
	}
}
		
VARIANT Ole::ExecuteMethode(VARIANT variant,Upp::WString methodName,int cArgs...)//Allow to execute methode attribute retrieve VARIANT
{
	try{
		va_list marker;
    	va_start(marker, cArgs);
		VARIANT buffer={0};
	    DISPPARAMS dp = { NULL, NULL, 0, 0 };

	    // Allocate memory for arguments...
	    VARIANT *pArgs = new VARIANT[cArgs+1];
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
	}catch(...){
		throw;
	}
}

VARIANT Ole::GetAttribute(Upp::WString attributeName,int cArgs...){
	try{
		va_list marker;
    	va_start(marker, cArgs);
		VARIANT buffer={0};
	    DISPPARAMS dp = { NULL, NULL, 0, 0 };
	    // Allocate memory for arguments...
	    VARIANT *pArgs = new VARIANT[cArgs+1];
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
	}catch(...){
		throw;
	}
}
VARIANT Ole::GetAttribute(VARIANT variant,Upp::WString attributeName,int cArgs...){
  	try{
  		va_list marker;
    	va_start(marker, cArgs);
  		// Variables used...
  		VARIANT buffer={0};
	    DISPPARAMS dp = { NULL, NULL, 0, 0 };
	    // Allocate memory for arguments...
	    VARIANT *pArgs = new VARIANT[cArgs+1];
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
	}catch(...){
		throw;
	}
}
