#include "IExplorer.h"

#include <time.h>
#include <comdef.h>

#define TIMEOUT_SECONDS 30
#define SLEEP_MILLISECONDS 500

IExplorer::IExplorer()
{
	this->isStarted=false;
	OleInitialize (NULL);
}

IExplorer::~IExplorer() { CoUninitialize(); }

bool IExplorer::Start() //Start new InternetExplorer Application
{
	SHANDLE_PTR handle;
	if(!this->isStarted && SUCCEEDED(CoCreateInstance(CLSID_InternetExplorer, NULL, CLSCTX_LOCAL_SERVER, IID_IWebBrowser2, (void**)&browser))) {
		this->isStarted = true;
		browser->get_HWND(&handle);
		WaitUntilNotBusy();
		return this->isStarted;
	} return false;
}

bool IExplorer::Search()
{
	IShellWindows *psw;
	HRESULT hr = CoCreateInstance(CLSID_ShellWindows,NULL,CLSCTX_ALL,IID_IShellWindows,(void**)&psw);
	if (FAILED(hr)) return false;
	IWebBrowser2* pBrowser2 = 0;
	long nCount = 0;
	hr = psw->get_Count(&nCount);
	if (SUCCEEDED(hr)) {
		for (long i = nCount - 1; (i >= 0); i--) {
			// get interface to item no i
			_variant_t va(i, VT_I4);
			IDispatch * spDisp;
			hr = psw->Item(va,&spDisp);
			hr = spDisp->QueryInterface(IID_IWebBrowserApp,(void **)&pBrowser2);
			if (SUCCEEDED(hr)) {
				BSTR name;
				pBrowser2->get_FullName(&name);
				Upp::String n = BSTRtoString(name);
				if (n.Find("IEXPLORE") == -1) {
					pBrowser2->Release();
					return false;
				} else {
					this->browser = pBrowser2;
					this->isStarted = true;
					return true;
				}
			}
		}
		psw->Release();
	}
	return false;
}

bool IExplorer::Search(Upp::WString url)
{
	IShellWindows *psw;
	HRESULT hr = CoCreateInstance(CLSID_ShellWindows,NULL,CLSCTX_ALL,IID_IShellWindows,(void**)&psw);
	if (FAILED(hr)) return false;
	IWebBrowser2* pBrowser2 = 0;
	long nCount = 0;
	hr = psw->get_Count(&nCount);
	if (SUCCEEDED(hr)) {
		for (long i = nCount - 1; i >= 0; i--) {
			// get interface to item no i
			_variant_t va(i, VT_I4);
			IDispatch * spDisp;
			hr = psw->Item(va,&spDisp);
			hr = spDisp->QueryInterface(IID_IWebBrowserApp,(void **)&pBrowser2);
			if (SUCCEEDED(hr)) {
				BSTR name;
				pBrowser2->get_LocationURL(&name);
				Upp::WString n(name);
				if (n.Find(url) == -1) {
					pBrowser2->Release();
				} else {
					this->isStarted = true;
					this->browser = pBrowser2;
					return true;
				}
			}
		}
		psw->Release();
	}
	return false;
}

Upp::String IExplorer::GetHtml(){
	UpdateHTMLDocPtr(); //c'est une phylosophie
		 
    IHTMLElement *lpBodyElm;
    html->get_html(&lpBodyElm);
    BSTR    bstr;
    lpBodyElm->get_outerHTML(&bstr);
    return BSTRtoString(bstr);
}

bool IExplorer::Stop()
{
	if(SUCCEEDED(browser->Stop()))
		return true;
	return false;
}

bool IExplorer::Quit()
{
	if(SUCCEEDED(browser->Quit())) {
		this->isStarted = false;
		return true;
	} return false;
}

bool IExplorer::SetVisible(bool set)
{
	if(this->isStarted && SUCCEEDED(browser->put_Visible(set))) {
		return true;
	} return false;
}

bool IExplorer::SetFullScreen(bool set)
{
	if(this->isStarted && SUCCEEDED(browser->put_TheaterMode(set))) {
		return true;
	} return false;
}

bool IExplorer::SetSilent(bool set)
{
	if(this->isStarted && SUCCEEDED(browser->put_Silent(set))) {
		return true;
	} return false;
}

bool IExplorer::SetToolBar(bool set)
{
	if (this->isStarted && SUCCEEDED(browser->put_ToolBar(set))) {
		return true;
	} return false;
}
bool IExplorer::SetStatusBar(bool set)
{
	if (this->isStarted && SUCCEEDED(browser->put_StatusBar(set))) {
		return true;
	} return false;
}
bool IExplorer::SetAddressBar(bool set)
{
	if (this->isStarted && SUCCEEDED(browser->put_AddressBar(set))) {
		return true;
	} return false;
}
bool IExplorer::SetMenuBar(bool set)
{
	if (this->isStarted && SUCCEEDED(browser->put_MenuBar(set))) {
		return true;
	} return false;
}

bool IExplorer::Navigate (Upp::WString url) {
	// Ex : url = L"http://castrec-pierre.netlify.com"
	if(this->isStarted && SUCCEEDED(browser->Navigate(AllocateString(url).bstrVal, NULL, NULL, NULL, NULL))) {
		WaitUntilReady();
		return true;
	} return false;
}

Upp::String IExplorer::GetURL()
{
	BSTR url;
	this->WaitUntilNotBusy();
	try {
		if(this->isStarted && SUCCEEDED(browser->get_LocationURL(&url))) {
			return BSTRtoString(url);
		} return "Error : getUrl()";
	} catch (const OleException &e) {
		throw OleException(30, "get_LocationURL", 1);
	}
}

Upp::String IExplorer::GetCookie()
{
	BSTR cookie;
	//this->WaitUntilNotBusy();
	this->UpdateHTMLDocPtr();
	try {
		if(this->isStarted && SUCCEEDED(html->get_cookie(&cookie))) {
			if(cookie) return BSTRtoString(cookie);
		} return "Error : getCookie()";
	} catch (const OleException &e) {
		throw OleException(32, "get_Cookie", 1);
	}
}

Upp::String IExplorer::FindTitle()
{
	char str[1024];
	SHANDLE_PTR lg;
	if(SUCCEEDED(browser->get_HWND(&lg))) {
		char* titleName = str;
		GetWindowTextA((HWND)lg, (LPSTR)titleName, sizeof(str));
		return titleName;
	} return "Error";
}

void IExplorer::WaitUntilNotBusy () {
	VARIANT_BOOL busy;
	time_t startTime = time(NULL);
	do {
		Sleep(SLEEP_MILLISECONDS);
		if(FAILED(browser->get_Busy(&busy)))
			throw OleException(23, "get_Busy", 1);
	} while ((busy==VARIANT_TRUE) && (difftime(time(NULL), startTime)<TIMEOUT_SECONDS));
	if (busy == VARIANT_TRUE)
		throw OleException(24, "Timeout while waiting 'get_Busy=false'", 1);
}

void IExplorer::WaitUntilReady () {
	READYSTATE isReady;
	time_t startTime = time(NULL);
	do {
		browser->get_ReadyState(&isReady);
		Sleep (SLEEP_MILLISECONDS);
	} while ((isReady!=READYSTATE_COMPLETE) && (difftime(time(NULL), startTime)<TIMEOUT_SECONDS));
	if (isReady != READYSTATE_COMPLETE)
		throw OleException(25, "Timeout while waiting READYSTATE_COMPLETE", 1);
}

void IExplorer::UpdateHTMLDocPtr () {
	IDispatch* pdisp;
	if (FAILED(browser->get_Document(&pdisp)))
		throw OleException (26, "get_Document", 2);
	if (FAILED(pdisp->QueryInterface(IID_IHTMLDocument2, (void **)&html)))
		throw OleException (27, "QueryInterface, IID_IHTMLDocument2", 2);
}

