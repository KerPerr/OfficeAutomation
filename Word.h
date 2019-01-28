#ifndef _OfficeAutomation_Word_h_
#define _OfficeAutomation_Word_h_

#include <Core/Core.h>
#include <stdio.h>
#include <windows.h>
#include "OfficeAutomation.h"

class WordApp;
class WordDocument;

class WordDocument : public Ole, Upp::Moveable<WordDocument> {
public:
	WordApp* app;
	Upp::String GetText();
	void SetText(Upp::String text);
	bool Close();
	bool Close(bool save);
	typedef WordDocument CLASSNAME;
	WordDocument(WordApp &app, VARIANT);
	WordDocument(const WordDocument&);
	bool operator==(const WordDocument&);
};

class WordApp : public Ole {
private:
	bool isStarted; //Bool to know if we started Word
	Upp::Thread myThread;
public:
	Upp::Vector<WordDocument> docs;
	int Count();
	WordDocument AddDocument();
	WordDocument OpenDocument(Upp::String path);
	bool FindOrStart(); //Find running Excel or Start new One
	bool RemoveDocument(WordDocument wdoc);
	bool SetVisible(bool isVisible);
	bool Start();
	bool Quit();
	typedef WordApp CLASSNAME;
	WordApp();
	~WordApp();
};

#endif
