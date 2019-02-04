#ifndef _OfficeAutomation_OfficeAutomation_h_
#define _OfficeAutomation_OfficeAutomation_h_
#include <Core/Core.h>
#include <windows.h> 
#include <exception>
#include <ocidl.h>
#include <typeinfo>

static const GUID IID_IApplicationEvents2Word =  {0x000209FE,0x0000,0x0000, {0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};
static const GUID IID_IApplicationEvents2Excel =  {0x00024500,0x0000,0x0000, {0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};
/* 
Project created 01/18/2019 
By Clément Hamon And Pierre Castrec
Lib used to drive every Microsoft Application's had OLE Compatibility.
This project have to be used with Ultimate++ FrameWork and required the Core Librairy from it
http://www.ultimatepp.org
Copyright © 1998, 2019 Ultimate++ team
All those sources are contained in "plugin" directory. Refer there for licenses, however all libraries have BSD-compatible license.
Ultimate++ has BSD license:
License : https://www.ultimatepp.org/app$ide$About$en-us.html
Thanks to UPP team
*/

class COfficeEventHandler;
class Ole;
class OleException;
struct IApplicationEvents2;

class Ole {
	private: 
		virtual HRESULT AutoWrap(int autoType, VARIANT *pvResult, IDispatch *pDisp, LPOLESTR ptName, DISPPARAMS dp);//Allow code execution on whatever object 
		
	public:
		typedef Ole CLASSNAME;
		~Ole();
		
		bool EventListened = false;
		Upp::Thread* eventListener;

		VARIANT AppObj;
		
		const Upp::WString WS_ExcelApp = L"Excel.Application"; //MS Excel
		const Upp::WString WS_WordApp = L"Word.Application"; //MS Word
		const Upp::WString WS_OutlookApp = L"Outlook.Application"; //MS Outlook
		const Upp::WString WS_PowerPointApp = L"PowerPoint.Application"; //MS PowerPoint
		const Upp::WString WS_InternetExplorerApp = L"InternetExplorer.Application"; //MS IE
		const Upp::WString WS_ProdApp = L"InternetExplorer.Application"; // this one is to use in my context, you'r supposed to never use it :p
		
		const Upp::WString WS_CLSID_WordApp = L"{000209FE-0000-0000-C000-000000000046}"; //Clsid of Word
		const Upp::WString WS_CLSID_ExcelApp = L"{00024500-0000-0000-C000-000000000046}"; //Clsid of Excel
		
		virtual VARIANT StartApp(const Upp::WString appName,bool startEventListener = false); 
		virtual VARIANT FindApp(const Upp::WString appName,bool startEventListener = false,bool isFindOnly = false);
		virtual void InitSinkCommunication(const Upp::WString appName);
		
		virtual Upp::String BSTRtoString (BSTR bstr); //Converting VARIANT.BSTR to Upp::String
		virtual void IndToStr(int row,int col,char* strResult);//translating row and column number into the string name of the cell.
		virtual int ColStrToInt(Upp::String target); //Return int represent col. The arg is a range (Example : "AB15")
		virtual int ExtractRow(Upp::String target); //Return int represent row. The arg is a range (Example : "AB15")
	
		virtual VARIANT AllocateString(Upp::String arg); //Easy way to allocate some data into variant to use it as arg
		virtual VARIANT AllocateString(Upp::WString arg);//Easy way to allocate some data into variant to use it as arg
		virtual VARIANT AllocateInt(int arg);//Easy way to allocate some data into variant to use it as arg
		
		const Upp::WString CLSIDbyName(const Upp::WString appName); //Get App name and return CLsid name is he can
		/****************************************************************************/
		// This section allow Free hand execution of code on every object
		// It mean you must know how VARIANT work to retrieve information you want
		/****************************************************************************/
		virtual VARIANT GetAttribute(Upp::WString attributeName); //Allow to retrieve attribute Value By VARIANT
		virtual VARIANT GetAttribute(VARIANT variant,Upp::WString attributeName);//Allow to retrieve attribute Value By VARIANT
		virtual VARIANT GetAttribute(IDispatch* pdisp, Upp::WString attributeName);//Allow to retrieve attribute Value By DISPATCHER
		
		virtual VARIANT GetAttribute(Upp::WString attributeName,int cArgs...);//Allow to retrieve attribute Value By VARIANT
		virtual VARIANT GetAttribute(VARIANT variant,Upp::WString attributeName,int cArgs...);//Allow to retrieve attribute Value By VARIANT
		virtual VARIANT GetAttribute(IDispatch* pdisp, Upp::WString attributeName,int cArgs...);//Allow to retrieve attribute Value By DISPATCHER
		
		virtual bool SetAttribute(Upp::WString attributeName, Upp::String value);//Allow to set attribute Value
		virtual bool SetAttribute(Upp::WString attributeName, int value);//Allow to set attribute Value
		virtual bool SetAttribute(VARIANT variant,Upp::WString attributeName, Upp::String value);//Allow to set attribute Value
		virtual bool SetAttribute(VARIANT variant,Upp::WString attributeName, int value);//Allow to set attribute Value
		virtual bool SetAttribute(IDispatch* pdisp,Upp::WString attributeName, int value);//Allow to set attribute Value
		virtual bool SetAttribute(IDispatch* pdisp,Upp::WString attributeName, Upp::String value);//Allow to set attribute Value
		
		virtual VARIANT ExecuteMethode(Upp::WString methodName,int cArgs...);//Allow to execute methode attribute retrieve VARIANT
		virtual VARIANT ExecuteMethode(VARIANT variant,Upp::WString methodName,int cArgs...);//Allow to execute methode attribute retrieve VARIANT
		virtual VARIANT ExecuteMethode(IDispatch* pdisp,Upp::WString methodName,int cArgs...);//Allow to execute methode attribute retrieve VARIANT
};

#include "Excel.h"
#include "Word.h"
#include "Outlook.h"
#include "IExplorer.h"

class OleException : public std::exception { //classe to managed every OLE exception
	private:
	    int m_numero;               //Id of Error
	    Upp::String m_phrase;       //Error summaries
	    int m_niveau;               //level of Error  0=> Invoque problem; 1 => Exception from OLE ; 2 => Exception from VARIANT Wrapper
	    char* myChar=NULL;

	public:
	    OleException(int numero=0, Upp::String phrase="", int niveau=0){
	        m_numero = numero;
	        m_phrase = phrase;
	        m_niveau = niveau;
	       	myChar =  new char[m_phrase.GetCount()+1];
	        strcpy(myChar,this->m_phrase.ToStd().c_str());
	    }
	    
	    virtual const char* what() const throw() {
	       	return  (const char *)  myChar;
	    }
	    int getNiveau() const throw(){
	    	return m_niveau;
	    }
		virtual ~OleException(){
			delete [] myChar;
		}
};



struct IApplicationEvents2 : public IDispatch // Pretty much copied from typelib
{
/*
 * IDispatch methods
 */
STDMETHODIMP QueryInterface(REFIID riid, void ** ppvObj) = 0; 
STDMETHODIMP_(ULONG) AddRef()  = 0;  
STDMETHODIMP_(ULONG) Release() = 0;

STDMETHODIMP GetTypeInfoCount(UINT *iTInfo) = 0;
STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) = 0;
STDMETHODIMP GetIDsOfNames(REFIID riid, OLECHAR **rgszNames, 
                              UINT cNames,  LCID lcid, DISPID *rgDispId) = 0;
STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid,
                              WORD wFlags, DISPPARAMS* pDispParams,
                              VARIANT* pVarResult, EXCEPINFO* pExcepInfo,
                              UINT* puArgErr) = 0;


};


class COfficeEventHandler : public IApplicationEvents2
{
	protected:
	LONG m_cRef;
	ExcelApp * excelInstance=NULL;
	WordApp * wordInstance=NULL;
	
	public:
	DWORD m_dwEventCookie;
	
	COfficeEventHandler(){
		m_cRef={1};
		m_dwEventCookie={0};
	}
	
template<class Type> COfficeEventHandler(Type* instance){
		m_cRef={1};
		m_dwEventCookie={0};
		Upp::Cout() << typeid(instance).name() <<"\n";
		Upp::Cout() << typeid(WordApp*).name() <<"\n";
		if(typeid(instance).name()==typeid(ExcelApp*).name()){
			excelInstance =  dynamic_cast<ExcelApp*>(instance);
		}
		else if(typeid(instance).name()==typeid(WordApp*).name()){
			wordInstance =  dynamic_cast<WordApp*>(instance);
		}
	}

	STDMETHOD_(ULONG, AddRef)()
	{
		InterlockedIncrement(&m_cRef);
		return m_cRef;  
	}
	
	STDMETHOD_(ULONG, Release)(){
		InterlockedDecrement(&m_cRef);
		if (m_cRef == 0)
		{
		    delete this;
		    return 0;
	}
		return m_cRef;
	}
	
	STDMETHOD(QueryInterface)(REFIID riid, void ** ppvObj)
	{
	 if (riid == IID_IUnknown){
	    *ppvObj = static_cast<IApplicationEvents2*>(this);
	}
	
	else if (riid == IID_IApplicationEvents2Word){
	    *ppvObj = static_cast<IApplicationEvents2*>(this);
	}
	else if (riid == IID_IApplicationEvents2Excel){
	    *ppvObj = static_cast<IApplicationEvents2*>(this);
	}
	else if (riid == IID_IDispatch){
	    *ppvObj = static_cast<IApplicationEvents2*>(this);
	}
	else
	{
	    char clsidStr[256];
	    WCHAR wClsidStr[256];
	    char txt[512];
	    StringFromGUID2(riid, (LPOLESTR)&wClsidStr, 256);
	    // Convert down to ANSI
	    WideCharToMultiByte(CP_ACP, 0, wClsidStr, -1, clsidStr, 256, NULL, NULL);
	    sprintf_s(txt, 512, "riid is : %s: Unsupported Interface", clsidStr);
	    Upp::Cout() << clsidStr <<"\n";
	    *ppvObj = NULL;
	    return E_NOINTERFACE;
	}
	
	static_cast<IUnknown*>(*ppvObj)->AddRef();
		return S_OK;
	}
	
	STDMETHOD(GetTypeInfoCount)(UINT* pctinfo){
		return E_NOTIMPL;
	}
	
	STDMETHOD(GetTypeInfo)(UINT itinfo, LCID lcid, ITypeInfo** pptinfo){
		return E_NOTIMPL;
	}
	
	STDMETHOD(GetIDsOfNames)(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgdispid){
		return E_NOTIMPL;
	}
	
	STDMETHOD(Invoke)(DISPID dispidMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pdispparams, VARIANT* pvarResult,EXCEPINFO* pexcepinfo, UINT* puArgErr){
		  //Validate arguments
	    if ((riid != IID_NULL))
	        return E_INVALIDARG;
	    HRESULT hr = S_OK;  // Initialize
	    /* To see what Word sends as dispid values */
	    static char myBuf[80];
	    memset( &myBuf, '\0', 80 );
	    sprintf_s( (char*)&myBuf, 80, " Dispid: %d :", dispidMember );
		Upp::Cout() << dispidMember <<"\n";
	    switch(dispidMember){
	    case 0x01:    // Startup
	       Upp::Cout() <<"Word Demarer" <<"\n";
	    break;
	    case 0x02:    // Quit
	       Upp::Cout() <<"Document Quit" <<"\n";
			if(excelInstance) Upp::Cout() <<" Document Excel !" << "\n";
			if(wordInstance) Upp::Cout() <<" Document Word !" << "\n";
	    break;
	    case 0x03:    // DocumentChange
	        Upp::Cout() <<"Document Change" <<"\n";
	    break;
	    }
	
	    return S_OK;
	}
	

};




#endif
