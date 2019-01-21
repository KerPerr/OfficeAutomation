#include <Core/Core.h>

#define _WIN32_WINNT 0x0501
#include <windows.h>
#include <ole2.h>
#include <stdio.h>
#include "OfficeAutomation.h"

#define DISP_FREEARGS 0x01
#define DISP_NOSHOWEXCEPTIONS 0x02

using namespace Upp;

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
                MessageBox(NULL, buf, "AutoWrap()", 0x10010);
                switch(hr)
                {
                case DISP_E_BADPARAMCOUNT:
                        MessageBox(NULL, "DISP_E_BADPARAMCOUNT", "Error:", 0x10010);
                        break;
                case DISP_E_BADVARTYPE:
                        MessageBox(NULL, "DISP_E_BADVARTYPE", "Error:", 0x10010);
                        break;
                case DISP_E_EXCEPTION:
                        MessageBox(NULL, "DISP_E_EXCEPTION", "Error:", 0x10010);
                        break;
                case DISP_E_MEMBERNOTFOUND:
                        MessageBox(NULL, "DISP_E_MEMBERNOTFOUND", "Error:", 0x10010);
                        break;
                case DISP_E_NONAMEDARGS:
                        MessageBox(NULL, "DISP_E_NONAMEDARGS", "Error:", 0x10010);
                        break;
                case DISP_E_OVERFLOW:
                        MessageBox(NULL, "DISP_E_OVERFLOW", "Error:", 0x10010);
                        break;
                case DISP_E_PARAMNOTFOUND:
                        MessageBox(NULL, "DISP_E_PARAMNOTFOUND", "Error:", 0x10010);
                        break;
                case DISP_E_TYPEMISMATCH:
                        MessageBox(NULL, "DISP_E_TYPEMISMATCH", "Error:", 0x10010);
                        break;
                case DISP_E_UNKNOWNINTERFACE:
                        MessageBox(NULL, "DISP_E_UNKNOWNINTERFACE", "Error:", 0x10010);
                        break;
                case DISP_E_UNKNOWNLCID:
                        MessageBox(NULL, "DISP_E_UNKNOWNLCID", "Error:", 0x10010);
                        break;
                case DISP_E_PARAMNOTOPTIONAL:
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
	return this->GetAttribute(this->AppObj,attributeName);
}
		
bool Ole::SetAttribute(Upp::WString attributeName, Upp::String value)//Allow to set attribute Value
{
	try{
	this->SetAttribute(this->AppObj,attributeName,value);
	return true;
	}catch(...){
		return false;	
	}
}
bool Ole::SetAttribute(Upp::WString attributeName, int value)//Allow to set attribute Value
{
	try{
	VARIANT buffer={0};
	this->SetAttribute(this->AppObj,attributeName,value);
	return true;
	}catch(...){
		return false;	
	}
}
		
VARIANT Ole::ExecuteMethode(Upp::WString methodName,int cArgs...)//Allow to execute methode attribute retrieve VARIANT
{
	va_list vl;
	va_start(vl,cArgs);
	return this->ExecuteMethode(this->AppObj,methodName,cArgs,va_arg(vl,VARIANT));
}


VARIANT Ole::GetAttribute(VARIANT variant,Upp::WString attributeName) //Allow to retrieve attribute Value By VARIANT
{
	VARIANT buffer={0};
	AutoWrap(DISPATCH_PROPERTYGET,&buffer,variant.pdispVal,(wchar_t*)~attributeName,0);
	return buffer;
}
		
bool Ole::SetAttribute(VARIANT variant,Upp::WString attributeName, Upp::String value)//Allow to set attribute Value
{
	try{
	VARIANT buffer={0};
	AutoWrap(DISPATCH_PROPERTYPUT,&buffer,variant.pdispVal,(wchar_t*)~attributeName,1,AllocateString(value));
	return true;
	}catch(...){
		return false;	
	}
}
bool Ole::SetAttribute(VARIANT variant,Upp::WString attributeName, int value)//Allow to set attribute Value
{
	try{
	VARIANT buffer={0};
	AutoWrap(DISPATCH_PROPERTYPUT,&buffer,variant.pdispVal,(wchar_t*)~attributeName,1,AllocateInt(value));
	return true;
	}catch(...){
		return false;	
	}
}
		
VARIANT Ole::ExecuteMethode(VARIANT variant,Upp::WString methodName,int cArgs...)//Allow to execute methode attribute retrieve VARIANT
{
	va_list vl;
	va_start(vl,cArgs);
	VARIANT buffer={0};
	AutoWrap(DISPATCH_METHOD,&buffer,variant.pdispVal,(wchar_t*)~methodName,cArgs,va_arg(vl,VARIANT ));
	return buffer;
}
