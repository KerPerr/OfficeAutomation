#ifndef _OfficeAutomation_Outlook_h_
#define _OfficeAutomation_Outlook_h_
#include "OfficeAutomation.h"
using namespace Upp;
/* 
Project created 01/18/2019 
By Clément Hamon Email: hamon.clement@outlook.fr
Lib used to drive every Microsoft Application's had OLE Compatibility.
This project have to be used with Ultimate++ FrameWork and required the Core Librairy from it
http://www.ultimatepp.org
Copyright © 1998, 2019 Ultimate++ team
All those sources are contained in "plugin" directory. Refer there for licenses, however all libraries have BSD-compatible license.
Ultimate++ has BSD license:
License : https://www.ultimatepp.org/app$ide$About$en-us.html
Thanks to UPP team
*/

class OutlookApp;//Class represents an  Outlook Application 
class MailItem;//Class represents Outlook Session

class OutlookApp : public Ole {
	private: 
		bool OutlookIsStarted; //Bool to know if we started Excel
		Upp::Thread myThread;
	public:
		OutlookApp(); //Initialise COM
		~OutlookApp(); //Unitialise COM
		
		MailItem CreateMail();
		
		bool Start(); //Start new Outlook Application
		bool FindOrStart(); //Find running Outlook or Start new One
		bool Quit(); //Close current Outlook Application
		
		bool FindApplication(); //Find First current Outlook Application still openned

		bool SetVisible(bool set); //Set or not the application visible 
};

class MailItem : public Ole{
	private:
		OutlookApp* parent; //pointer to OutlookApp	
		//olTo 
		
	
		
		//Atachement would come later
	public:
		MailItem(OutlookApp& parent, VARIANT appObj); 
		~MailItem();
		bool AddRecever(Upp::String email);
		//olCC
		bool AddRecipients(Upp::String email);
		bool SetSubject(Upp::String subject);
		bool SetBody(Upp::String body);
		bool setHighImportance();
		bool AddItem(String PathToData);
		bool DisplayMail();
};


#endif
