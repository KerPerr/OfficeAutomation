#ifndef _OfficeAutomation_InternetExplorer_h_
#define _OfficeAutomation_InternetExplorer_h_

#include "OfficeAutomation.h"

class InternetExplorer;

class InternetExplorer : public Ole {
private:
	bool isStarted; //Bool to know if we started InternetExplorer
public:
	Upp::String GetCookie();
	bool SetVisible(bool isVisible);
	bool Start();
	bool FindApplication();
	bool Quit();
	typedef InternetExplorer CLASSNAME;
	InternetExplorer();
	~InternetExplorer();
};

#endif
