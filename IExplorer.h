#ifndef _OfficeAutomation_IExplorer_h_
#define _OfficeAutomation_IExplorer_h_

#include "OfficeAutomation.h"
#include <mshtml.h>
#include <exdisp.h>
#include <time.h>

class IExplorer;

class IExplorer : public Ole {
private:
	IWebBrowser2* browser;
	IHTMLDocument2 *html;
	bool isStarted;
	
	void UpdateHTMLDocPtr();
	void WaitUntilNotBusy();
	void WaitUntilReady();
public:
	bool Start();
	bool Stop();
	bool Quit();
	
	bool SetVisible(bool set);
	bool SetSilent(bool set);
	bool SetToolBar(bool set);
	bool SetStatusBar(bool set);
	bool SetAddressBar(bool set);
	bool SetMenuBar(bool set);

	bool Navigate(Upp::WString url);
	
	Upp::String FindClass();
	Upp::String FindTitle();
	Upp::String GetURL();
	Upp::String GetCookie();
	
	typedef IExplorer CLASSNAME;
	IExplorer();
	~IExplorer();
};

#endif
