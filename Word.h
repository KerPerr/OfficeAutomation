#ifndef _OfficeAutomation_Word_h_
#define _OfficeAutomation_Word_h_

#include <Core/Core.h>
#include <stdio.h>
#include <windows.h>
#include "OfficeAutomation.h"

class WordApp;
class WordDocument;

class WordDocument : public Ole, Upp::Moveable<WordDocument> {
	WordApp* app;
public:
	Upp::String GetText();
	bool Close(int save);
	typedef WordDocument CLASSNAME;
	WordDocument(WordApp &app, VARIANT);
};

class WordApp : public Ole {
private:
	bool WordIsStarted; //Bool to know if we started Word
public:
	int Count();
	Upp::Vector<WordDocument> docs;
	WordDocument* AddDocument();
	WordDocument* OpenDocument(Upp::String path);
	bool RemoveDocument(WordDocument* wdoc);
	bool SetVisible(bool isVisible);
	bool Start();
	bool Quit();
	typedef WordApp CLASSNAME;
	WordApp();
	~WordApp();
};

#endif
