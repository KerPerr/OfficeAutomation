#include "OfficeAutomation.h"

using namespace Upp;

Imsp::Imsp(){//Initialise COM
	this->IsStarted=false;
	CoInitialize(NULL);
}

Imsp::~Imsp(){//Unitialise COM
//	~Ole();
	VariantClear(&this->AppObj);
	Thread::ShutdownThreads();
	CoUninitialize();
}

bool Imsp::Find()
{
	if(!this->IsStarted){
		this->AppObj = FindApp(WS_ProdApp);
		if( this->AppObj.intVal != -1) {
			//String path =   BSTRtoString(GetAttribute("FullName").bstrVal);
			
		    //ExecuteMethode(GetAttribute("Sessions"),"Open",1,AllocateString(path));
			
			VARIANT buffer = this->GetAttribute("ActiveSession");
			if(buffer.pdispVal) {
				this->AppObj = buffer;
				Wait();
				this->IsStarted=true;
				return true;
			}
			return false;
		}
		return false;
	}
	return false;
}

bool Imsp::Start(bool threading)
{
	if(!this->IsStarted) {
				
		this->AppObj = StartApp(WS_ProdApp);
		if( this->AppObj.intVal != -1) {
			VARIANT buffer = this->ExecuteMethode(this->GetAttribute("Sessions"), L"Open", 1, AllocateString(path));
			if(buffer.pdispVal) {
				this->AppObj = buffer;
				this->IsStarted=true;
				CLSID clsApp;
				IUnknown* punk;
				HRESULT hr = CLSIDFromProgID(L"Attachmate_Reflection_Objects_Framework.ApplicationObject", &clsApp);
				if(!FAILED(hr)){
					HRESULT hr2 = GetActiveObject( clsApp, NULL, &punk );
					if (!FAILED(hr2)) {
						hr2 = punk->QueryInterface(IID_IDispatch, (void **)&terminal.pdispVal);
						if (terminal.ppdispVal){
							frame = ExecuteMethode(terminal, "GetObject", 1, AllocateString(L"Frame")),
							terminal = GetAttribute(GetAttribute(frame, L"SelectedView"), L"control");
						}
					} else {
						Cout() << "FAILED" << EOL;
					}
				}
				if(threading && !worker.IsOpen()) {
					worker.Run([=] {
						CoInitialize(NULL);
						for(int i = 0; i < 100; i++){
							if(Thread::IsShutdownThreads())break;
							String status = BSTRtoString(GetAttribute(frame, L"StatusBarText").bstrVal);
							if(status.StartsWith("Ex")){
								Cout() << "Ohlalala une macro s'execute !" << EOL;
								ExecuteMethode(GetAttribute(terminal, L"Macro"), L"StopMacro", 0);
								break;
							}
							Sleep(100);
							if(Thread::IsShutdownThreads())break;
						}
						CoUninitialize();
					});
				}
				return true;
			}
			return false;
		}
		return false;
	}
	return false;
}

void Imsp::Wait()
{
	ExecuteMethode(GetAttribute("Screen"),L"WaitHostQuiet", 1, AllocateInt(10));
}

void Imsp::SetText(String text)
{
	ExecuteMethode(GetAttribute("Screen"),L"SendKeys", 1, AllocateString(text));
}

void Imsp::SetCmd(String cmd)
{
	ExecuteMethode(GetAttribute("Screen"),L"SendKeys", 1, AllocateString("<"+cmd+">"));
}

String Imsp::GetString(int row, int col, int len)
{
	return BSTRtoString(ExecuteMethode(GetAttribute("Screen"), L"getString", 3, AllocateInt(len),AllocateInt(col),AllocateInt(row)).bstrVal);
}

bool Imsp::NextPage()
{
	if(GetString(1, 12, 31-12+1) == "PAS DE PAGE SUIVANTE") {
		return false;
	} else {
		SetCmd("PF8");
		return true;
	}
}

bool Imsp::PrevPage()
{
	if(GetString(1, 12, 33-12+1) == "PAS DE PAGE PRECEDENTE") {
		return false;
	} else {
		SetCmd("PF7");
		return true;
	}
}