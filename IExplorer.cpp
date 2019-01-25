#include "IExplorer.h"

#include <time.h>
#include <comdef.h>

#define TIMEOUT_SECONDS 30
#define SLEEP_MILLISECONDS 500

IExplorer::IExplorer(){
	this->isStarted=false;
	CLSID rclsid;
	OleInitialize (NULL);
}

IExplorer::~IExplorer(){
	CoUninitialize();
}

bool IExplorer::Start() //Start new InternetExplorer Application
{
	if(!this->isStarted){
		CLSID rclsid;
		if (CLSIDFromProgID(WS_InternetExplorerApp, &rclsid) != S_OK)
			throw OleException(21, "CLSIDFromProgID", 0);
		if (CoCreateInstance(rclsid, NULL, CLSCTX_SERVER, IID_IWebBrowser2, (LPVOID*)&ptrWebBrowser) != S_OK)
			throw OleException(22, "CoCreateInstance", 0);
		this->isStarted = true;
		WaitUntilNotBusy();
		return this->isStarted;
	} return false;
}

bool IExplorer::SetVisible(bool set)//Set or not the application visible
{
	if(this->isStarted){
		if (ptrWebBrowser->put_Visible(VARIANT_TRUE) != S_OK)
			return false;
		else
			return true;
	}
	return false;
}

bool IExplorer::SetSilent(bool set)
{
	if(this->isStarted){
		if (ptrWebBrowser->put_Silent(VARIANT_TRUE) != S_OK)
			return false;
		else
			return true;
	}
	return false;
}

void IExplorer::Navigate (Upp::WString url) {
	if(this->isStarted){
		// Ex : url = L"http://castrec-pierre.netlify.com"
		ptrWebBrowser->Navigate(AllocateString(url).bstrVal, NULL, NULL, NULL, NULL);
		WaitUntilReady();
	}
}
/*
Upp::String IExplorer::GetURL()
{
	if(this->isStarted){
		ptrWebBrowser->
	}
}
*/
void IExplorer::WaitUntilNotBusy () {
	VARIANT_BOOL busy;
	time_t startTime = time(NULL);
	do {
		Sleep (SLEEP_MILLISECONDS);
		if (ptrWebBrowser->get_Busy (&busy) != S_OK)
			throw OleException(24, "get_Busy", 1);
	} while ((busy==VARIANT_TRUE) && (difftime(time(NULL), startTime)<TIMEOUT_SECONDS));
	if (busy == VARIANT_TRUE)
		throw OleException(23, "Timeout while waiting 'get_Busy=false'", 1);
}

void IExplorer::WaitUntilReady () {
	READYSTATE isReady;
	time_t startTime = time(NULL);
	do {
		ptrWebBrowser->get_ReadyState(&isReady);
		Sleep (SLEEP_MILLISECONDS);
	} while ((isReady!=READYSTATE_COMPLETE) && (difftime(time(NULL), startTime)<TIMEOUT_SECONDS));
	if (isReady != READYSTATE_COMPLETE)
		throw OleException(26, "Timeout while waiting READYSTATE_COMPLETE", 1);
}