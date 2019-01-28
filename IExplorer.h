#ifndef _OfficeAutomation_IExplorer_h_
#define _OfficeAutomation_IExplorer_h_

#include "OfficeAutomation.h"
#include <mshtml.h>
#include <exdisp.h>
#include <time.h>

class IExplorer;

class IExplorer : public Ole {
private:
	IWebBrowser2* ptrWebBrowser;
	IHTMLDocument3 *htmlDocPtr;
	IDispatch* ptrDispatch;
	bool isStarted; //Bool to know if we started InternetExplorer
public:
	Upp::String GetCookie();
	void Navigate(Upp::WString url);
	void WaitUntilNotBusy();
	void WaitUntilReady();
	bool Stop();
	bool SetVisible(bool isVisible);
	bool SetSilent(bool isVisible);
	bool Start();
	bool Quit();
	typedef IExplorer CLASSNAME;
	IExplorer();
	~IExplorer();
};

#endif
