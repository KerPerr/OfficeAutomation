#ifndef _OfficeAutomation_IExplorer_h_
#define _OfficeAutomation_IExplorer_h_

#include "OfficeAutomation.h"
#include <mshtml.h>
#include <exdisp.h>
#include <time.h>

class IExplorer;

class IExplorer : public Ole {
public:
	IWebBrowser2* browser;
	IHTMLDocument2 *html;
	bool isStarted;
	
	void UpdateHTMLDocPtr();
	void WaitUntilNotBusy();
	void WaitUntilReady();
private:
	bool Search();
	bool Search(Upp::WString url);
	
	bool Start();
	bool Stop();
	bool Quit();
	
	bool SetVisible(bool set);
	bool SetSilent(bool set);
	bool SetFullScreen(bool set);
	bool SetToolBar(bool set);
	bool SetStatusBar(bool set);
	bool SetAddressBar(bool set);
	bool SetMenuBar(bool set);

	bool Navigate(Upp::WString url);
	bool Open(Upp::WString url);
	
	Upp::String GetURL();
	Upp::String GetCookie();
	Upp::String FindTitle();
	
	typedef IExplorer CLASSNAME;
	IExplorer();
	~IExplorer();
};

#endif
