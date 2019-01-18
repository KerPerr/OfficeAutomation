#ifndef _OfficeAutomation_OfficeAutomation_h_
#define _OfficeAutomation_OfficeAutomation_h_

#include <Core/Core.h>
#include <windows.h> 

/* 
 Project created 01/18/2019 
 By Clément Hamon And Pierre Castrec
 Lib used to drive every Microsoft Application's had OLE Compatibility.
 This project have to be used with Ultimate++ FrameWork and required the Core Librairy from it
*/

class Ole;

class Ole {
	public:
		VARIANT AppObj;
		
		const Upp::WString WS_ExcelApp = L"Excel.Application"; //MS Excel
		const Upp::WString WS_WordApp = L"Word.Application"; //MS Word
		const Upp::WString WS_OutlookApp = L"Outlook.Application"; //MS Outlook
		const Upp::WString WS_PowerPointApp = L"PowerPoint.Application"; //MS PowerPoint
		const Upp::WString WS_InternetExplorerApp = L"InternetExplorer.Application"; //MS IE
		const Upp::WString WS_ProdApp = L"InternetExplorer.Application"; // this one is to use in my context, you'r supposed to never use it :p
		
		virtual VARIANT StartApp(const Upp::WString appName); 
		virtual HRESULT AutoWrap(int autoType, VARIANT *pvResult, IDispatch *pDisp, LPOLESTR ptName, int cArgs...);//Allow code execution on whatever object 
		virtual Upp::String BSTRtoString (BSTR bstr); //Converting VARIANT.BSTR to Upp::String
		virtual void IndToStr(int row,int col,char* strResult);//translating row and column number into the string name of the cell.
		
		virtual VARIANT AllocateString(Upp::String arg); //Easy way to allocate some data into variant to use it as arg
		virtual VARIANT AllocateString(Upp::WString arg);//Easy way to allocate some data into variant to use it as arg
		virtual VARIANT AllocateInt(int arg);//Easy way to allocate some data into variant to use it as arg
		
		/*******************************************************************/
		// This section allow Free hand execution of code on every object
		// It mean you must know how VARIANT work to retrieve information you want
		/*******************************************************************/
		virtual VARIANT GetAttribute(Upp::WString attributeName); //Allow to retrieve attribute Value By VARIANT
		
		virtual bool SetAttribute(Upp::WString attributeName, Upp::String value);//Allow to set attribute Value
		virtual bool SetAttribute(Upp::WString attributeName, int value);//Allow to set attribute Value
		
		virtual VARIANT ExecuteMethode(Upp::WString methodName,int cArgs...);//Allow to execute methode attribute retrieve VARIANT
};

#include "Excel.h"



#endif
