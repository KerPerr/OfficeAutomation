#include "OfficeAutomation.h"

WordApp::WordApp(){
	this->isStarted=false;
	CoInitialize(NULL);
}

WordApp::~WordApp(){
//	~Ole();

	CoUninitialize();
}

bool WordApp::Start(bool startEventListener ) //Start new Word Application
{
	if(!this->isStarted){
		this->AppObj = this->StartApp(WS_WordApp,startEventListener);
		if( this->AppObj.intVal != -1){
			this->isStarted=true;
			return true;
		}
		return false;
	}
	return false;
}

bool WordApp::Quit() //Close current Word Application
{
	if(this->isStarted){
		try{
			if(EventListened){
				eventListener->ShutdownThreads();
				delete eventListener;
				EventListened = false;
			}
			this->isStarted = false;
			this->ExecuteMethode("Quit",0);
			return true;
		}catch(...){
			return false;
		}
	}
	return false;
}

bool WordApp::FindOrStart(bool startEventListener){
	if(!this->isStarted){
		this->AppObj = this->FindApp(WS_WordApp,startEventListener);
		if( this->AppObj.intVal != -1){
			this->isStarted=true;
			return true;
		}
	}
	return false;	
}

int WordApp::Count()
{
	return docs.GetCount();
	
}

WordDocument WordApp::AddDocument()
{
	try {
		return this->docs.Add(WordDocument(*this, this->ExecuteMethode(this->GetAttribute(L"Documents"), L"Add", 0)));
	} catch (...) {
		Upp::Cout() << "Error Add Document";
	}
}

WordDocument WordApp::OpenDocument(Upp::String path)
{
	try {
		return this->docs.Add(WordDocument(*this, this->ExecuteMethode(this->GetAttribute(L"Documents"), L"Open", 1, AllocateString(path))));
	} catch (...) {
		Upp::Cout() << "Error Open Document";
	}

}

bool WordApp::RemoveDocument(WordDocument wdoc){
    bool trouver = false;
    for(int i=0;i<docs.GetCount();i++){
        if(wdoc == docs[i]){
            trouver = true;
			docs.Remove(i);
            break;
        }
    }
    return trouver;
}

bool WordApp::SetVisible(bool set)//Set or not the application visible
{
	if(this->isStarted){
		try{
			this->SetAttribute("Visible",(int)set);
			return true;
		}catch(...){
			return false;
		}
	}
	return false;
}

WordDocument::WordDocument(WordApp &app, VARIANT doc)
{
	this->app = &app;
	this->AppObj = doc;
}

WordDocument::WordDocument(const WordDocument& a){
    this->app = a.app;
    this->AppObj = a.AppObj;
}

bool WordDocument::operator==(const WordDocument& wdoc)
{
	if(this->AppObj.pdispVal == wdoc.AppObj.pdispVal) {
		return true;
	} else {
		return false;
	}
}

Upp::String WordDocument::GetText()
{
	return BSTRtoString(this->GetAttribute(this->GetAttribute(L"Content"), L"Text").bstrVal);
}

void WordDocument::SetText(Upp::String text)
{
	this->SetAttribute(this->GetAttribute(L"Content"), L"Text", text);
}

bool WordDocument::Close()
{
	try {
		if(app->RemoveDocument(*this))
			this->ExecuteMethode(L"Close", 1, AllocateInt(-2));
		return true;
	} catch(...) {
		return false;
	}
}

bool WordDocument::Close(bool save)
{
	int arg = save ? -1 : 0;
	try {
		if(app->RemoveDocument(*this))
			this->ExecuteMethode(L"Close", 1, AllocateInt(arg));
		return true;
	} catch(...) {
		return false;
	}
}
