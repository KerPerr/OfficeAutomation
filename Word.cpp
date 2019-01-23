#include "Word.h"
#include <ole2.h>

WordDocument::WordDocument(WordApp &app, VARIANT doc)
{
	this->app = &app;
	this->AppObj = doc;
}

Upp::String WordDocument::GetText()
{
	return BSTRtoString(this->GetAttribute(this->GetAttribute(L"Content"), L"Text").bstrVal);
}

bool WordDocument::Close(int save) // ALWAYS SAVE !
{
	try {
		if(app->RemoveDocument(this))
			this->ExecuteMethode(L"Close", 1, AllocateInt(save));
		return true;
	} catch(...) {
		return false;
	}
}

WordApp::WordApp(){
	this->WordIsStarted=false;
	CoInitialize(NULL);
}

WordApp::~WordApp(){
	CoUninitialize();
}

bool WordApp::Start() //Start new Word Application
{
	if(!this->WordIsStarted){
		this->AppObj = this->StartApp(WS_WordApp);
		if( this->AppObj.intVal != -1){
			this->WordIsStarted=true;
			return true;
		}
		return false;
	}
	return false;
}

bool WordApp::Quit() //Close current Word Application
{
	if(this->WordIsStarted){
		try{
			this->ExecuteMethode("Quit",0);
			return true;
		}catch(...){
			return false;
		}
	}
	return false;
}

int WordApp::Count()
{
	return docs.GetCount();
}

WordDocument* WordApp::AddDocument()
{
	try {
		return &this->docs.Add(WordDocument(*this, this->ExecuteMethode(this->GetAttribute(L"Documents"), L"Add", 0)));
	} catch (...) {
		Upp::Cout() << "Error Add Document";
	}
}

WordDocument* WordApp::OpenDocument(Upp::String path)
{
	try {
		return &this->docs.Add(WordDocument(*this, this->ExecuteMethode(this->GetAttribute(L"Documents"), L"Open", 1, AllocateString(path))));
	} catch (...) {
		Upp::Cout() << "Error Open Document";
	}
}

bool WordApp::RemoveDocument(WordDocument* wdoc){
    bool trouver = false;
    Upp::Cout() << docs.GetCount() << '\n';
    for(int i=0;i<docs.GetCount();i++){
        Upp::Cout() << wdoc <<  ":" << &docs[i] <<"\n";
        if( wdoc == &docs[i]){
			docs.Remove(i);
            break;
        }
    }
    return trouver;
}

bool WordApp::SetVisible(bool set)//Set or not the application visible
{
	if(this->WordIsStarted){
		try{
			this->SetAttribute("Visible",(int)set);
			return true;
		}catch(...){
			return false;
		}
	}
	return false;
}
