#ifndef _OfficeAutomation_IExplorer_h_
#define _OfficeAutomation_IExplorer_h_

#include "OfficeAutomation.h"

class IExplorer;

class IExplorer : public Ole {
private:
	bool isStarted; //Bool to know if we started InternetExplorer
public:
	Upp::String GetCookie();
	bool SetVisible(bool isVisible);
	bool Start();
	bool FindApplication();
	bool Quit();
	typedef IExplorer CLASSNAME;
	IExplorer();
	~IExplorer();
};

#endif
