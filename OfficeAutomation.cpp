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
      throw OleException(13,"CLSIDFromProgID() => App Named " + appName.ToString() +" Can't be find",1);
   }
	
	if (FAILED(CoCreateInstance(clsApp, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void **)&App.pdispVal)))
	{
		MessageBox(NULL, "this App's not registered properly", "Error", 0x10010);
		throw OleException(14,"CoCreateInstance() => this App's ("+ appName.ToString()  +")not registered properly",1);
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
	}catch(OleException const& exception){
		delete [] pArgs;
		throw exception;
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
	}catch(OleException const& exception){
		delete [] pArgs;
		throw exception;
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
	}catch(OleException const& exception){
		delete [] pArgs;
		throw exception;
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
	}catch(OleException const& exception){
		delete [] pArgs;
		throw exception;
	}
}


VARIANT Ole::GetAttribute(VARIANT variant,Upp::WString attributeName) //Allow to retrieve attribute Value By VARIANT
{
	VARIANT *pArgs;
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
	}catch(OleException const& exception){
		delete [] pArgs;
		throw exception;
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
	}catch(OleException const& exception){
		delete [] pArgs;
		throw exception;
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
	}catch(OleException const& exception){
		delete [] pArgs;
		throw exception;
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
	}catch(OleException const& exception){
		delete [] pArgs;
		throw exception;
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
	}catch(OleException const& exception){
		delete [] pArgs;
		throw exception;
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
	}catch(OleException const& exception){
		delete [] pArgs;
		throw exception;
	}
}

int Ole::ColStrToInt(Upp::String target){
	int resultat= 0;
	for(int i = 0; i < target.GetCount(); i++){	
		if((int)toupper(target[i]) >64 && 	(int)toupper(target[i]) < 91){
			if (i>0) {
				resultat+=25;
			}
			resultat+= ((int)toupper(target[i]) -64);
		}
	}
	return resultat;
}

int Ole::ExtractRow(Upp::String target)
{
	char myRow[target.GetCount()];
	int iterator = 0;
	for(int i = 0; i < target.GetCount(); i++){	
		if( int(target[i]) >47 && int(target[i]) < 57){
			myRow[iterator] = target[i];
			iterator++;
		}
	}
	return atoi(myRow);
}

void Ole::DumpVariant(){
	
	VARIANT variant = this->AppObj;
//	Cout()<<(long) variant.lVal<<"\n";
//	Cout()<<(byte) variant.bVal<<"\n";
//	Cout()<< (short)variant.iVal<<"\n";
//	Cout()<< (float)variant.fltVal<<"\n";
//	Cout()<< (double)variant.dblVal<<"\n";
//	Cout()<< variant.boolVal<<"\n";
//	Cout()<< variant.scode<<"\n";
//	Cout()<< variant.cyVal<<"\n";
//	Cout()<< variant.date<<"\n";
//	Cout()<< BSTRtoString(variant.bstrVal ) <<"\n";
/*	Cout()<< *variant.punkVal<<"\n";
	Cout()<< *variant.pdispVal<<"\n";
	Cout()<< *variant.parray<<"\n";
	Cout()<< *variant.pbVal<<"\n";
	Cout()<< *variant.piVal<<"\n";
	Cout()<< *variant.plVal<<"\n";
	Cout()<< *variant.pllVal<<"\n";
	Cout()<< *variant.pfltVal<<"\n";
	Cout()<< *variant.pdblVal<<"\n";
	Cout()<< *variant.pboolVal<<"\n";
	Cout()<< *variant.pscode<<"\n";
	Cout()<< *variant.pcyVal<<"\n";
	Cout()<< *variant.pdate<<"\n"; 
	Cout()<< *variantj.pbstrVal<<"\n";
	Cout()<< **variant.ppunkVal<<"\n";
	Cout()<< **variantj.ppdispVal<<"\n";
	Cout()<< **variant.pparray<<"\n";
	Cout()<< *variant.pvarVal<<"\n";*/
//	Cout()<< variant.byref<<"\n";
//	Cout()<< (char)variant.cVal<<"\n";
//	Cout()<< (unsigned short)variant.uiVal<<"\n";
//	Cout()<< (unsigned long) variant.ulVal<<"\n";
//	Cout()<< variant.ullVal<<"\n";
//	Cout()<< (int)variant.intVal<<"\n";
//	Cout()<< (unsigned int)variant.uintVal<<"\n";
//	Cout()<< *variant.pdecVal<<"\n";
/*	while(variant.pcVal++){
	Cout()<<"| "  <<*variant.pcVal<<"\n";	
	}
	*/
/*	Cout()<< *variant.puiVal<<"\n";
	Cout()<< *variant.pulVal<<"\n";
	Cout()<< *variant.pullVal<<"\n";
	Cout()<< *variant.pintVal<<"\n";
	Cout()<< *variant.puintVal<<"\n";*/
}

void Ole::DumpVariant(VARIANT variant){
	/*
	Cout()<< variant.llVal <<"\n";
	Cout()<< tvariant.lVal<<"\n";
	Cout()<< variant.bVal<<"\n";
	Cout()<< variant.iVal<<"\n";
	Cout()<< variant.fltVal<<"\n";
	Cout()<< variant.dblVal<<"\n";
	Cout()<< variant.boolVal<<"\n";
	Cout()<< variant.scode<<"\n";
	Cout()<< variant.cyVal<<"\n";
	Cout()<< variant.date<<"\n";
	Cout()<< variant.bstrVal<<"\n";
	Cout()<< *variant.punkVal<<"\n";
	Cout()<< *variant.pdispVal<<"\n";
	Cout()<< *variant.parray<<"\n";
	Cout()<< *variant.pbVal<<"\n";
	Cout()<< *variant.piVal<<"\n";
	Cout()<< *variant.plVal<<"\n";
	Cout()<< *variant.pllVal<<"\n";
	Cout()<< *variant.pfltVal<<"\n";
	Cout()<< *variant.pdblVal<<"\n";
	Cout()<< *variant.pboolVal<<"\n";
	Cout()<< *variant.pscode<<"\n";
	Cout()<< *variant.pcyVal<<"\n";
	Cout()<< *variant.pdate<<"\n"; 
	Cout()<< *variantj.pbstrVal<<"\n";
	Cout()<< **variant.ppunkVal<<"\n";
	Cout()<< **variantj.ppdispVal<<"\n";
	Cout()<< **variant.pparray<<"\n";
	Cout()<< *variant.pvarVal<<"\n";
	Cout()<< variant.byref<<"\n";
	Cout()<< variantj.cVal<<"\n";
	Cout()<< variant.uiVal<<"\n";
	Cout()<< variant.ulVal<<"\n";
	Cout()<< variant.ullVal<<"\n";
	Cout()<< variant.intVal<<"\n";
	Cout()<< variant.uintVal<<"\n";
	Cout()<< *variant.pdecVal<<"\n";
	Cout()<< *variant.pcVal<<"\n";
	Cout()<< *variant.puiVal<<"\n";
	Cout()<< *variant.pulVal<<"\n";
	Cout()<< *variant.pullVal<<"\n";
	Cout()<< *variant.pintVal<<"\n";
	Cout()<< *variant.puintVal<<"\n";*/
}