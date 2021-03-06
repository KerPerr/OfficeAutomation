#include <Core/Core.h>
#include <windows.h>
#include <ole2.h>
#include <stdio.h>

#include "OfficeAutomation.h"

#define DISP_FREEARGS 0x01
#define DISP_NOSHOWEXCEPTIONS 0x02

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

/************************************************************************************************************************/

ExcelApp::ExcelApp(){//Initialise COM
	this->ExcelIsStarted=false;
	CoInitialize(NULL);
}

ExcelApp::~ExcelApp(){//Unitialise COM
//	~Ole();
	for(ExcelWorkbook& wb : workbooks){
		wb.Close();	
	}
	VariantClear(&this->AppObj);
	CoUninitialize();
}

bool ExcelApp::Start(bool startEventListener ) //Start new Excel Application
{
	if(!this->ExcelIsStarted){
		this->AppObj = this->StartApp(WS_ExcelApp, startEventListener);
		if( this->AppObj.intVal != -1){
			ResolveWorkbook();
			this->ExcelIsStarted=true;
			return true;
		}
		return false;
	}
	return false;
}

bool ExcelApp::Find(bool startEventListener){
	if(!this->ExcelIsStarted){
		this->AppObj = this->FindApp(WS_ExcelApp, startEventListener,true);
		if( this->AppObj.intVal != -1) {
			try {
				SetAttribute("EditDirectlyInCell", 0);
			} catch (OleException const& exception) {
				throw;
			}
			ResolveWorkbook();
			this->ExcelIsStarted=true;
			return true;
		}else{
			return false;
		}
	}
	return false;
}

bool ExcelApp::FindOrStart(bool startEventListener ){//Find running Excel or Start new One
	if(!this->ExcelIsStarted){
		this->AppObj = this->FindApp(WS_ExcelApp, startEventListener);
		if( this->AppObj.intVal != -1) {
			try {
				SetAttribute("EditDirectlyInCell", 0);
			} catch (OleException const& exception) {
				throw;
			}
			ResolveWorkbook();
			this->ExcelIsStarted=true;
			return true;
		}else{
			return Start(startEventListener);	
		}
	}
	return false;
}

bool ExcelApp::FindOrStartPredictedWorkbook(Upp::String name,bool startEventListener){
	Cout() <<"FindOrStartPredictedWorkbook "<<"\n";
	if(Find(startEventListener)){
		Cout() <<"Excel trouvé, Verification des workbook "<<"\n";
		String wbName = "";
		for(ExcelWorkbook &w : workbooks){
			wbName = name.Right(name.GetCount() - (name.ReverseFind("\\") +1));
			if (w.Name().Compare(name)==0){
				w.ResolveSheet();
				Cout() <<"Workbook trouvé"<<"\n";
				return true;
			}
		}
		Cout() <<"Workbook introuvable, ouverture d'un nouveau excel"<<"\n";
		this->ExcelIsStarted=false;
		return Start(startEventListener);
		Cout() <<"nouveau excel ouvert"<<"\n";	
		return true;
	}else{
		Cout() <<"Aucun excel, ouverture d'un nouveau"<<"\n";
		return Start(startEventListener);	
	}
}

bool ExcelApp::Quit() //Close current Excel Application
{
	if(this->ExcelIsStarted){
		try{
			if(EventListened){
				Upp::Thread::ShutdownThreads();
				delete eventListener;
				EventListened = false;
			}
			this->ExcelIsStarted = false;
			this->ExecuteMethode("Quit",0);
			return true;
		}catch(OleException const& exception){
			throw;
		}
	}
	return false;
}
	
bool ExcelApp::FindApplication(bool startEventListener){ //Find First current Excel Application still openned
	if(!this->ExcelIsStarted){
		this->AppObj = this->FindApp(WS_ExcelApp,startEventListener,true);
		if( this->AppObj.intVal != -1){
			this->ExcelIsStarted=true;
			ResolveWorkbook();
			return true;
		}
	}
	return false;
}
	
bool ExcelApp::SetVisible(bool set)//Set or not the application visible
{
	if(this->ExcelIsStarted){
		try{
			this->SetAttribute("Visible",(int)set);
			return true;
		}catch(OleException const& exception){
			throw;
		}
	}
	return false;
}

bool ExcelApp::DisplayAlerts(bool set){
	if(this->ExcelIsStarted){
		try{
			this->SetAttribute("DisplayAlerts",(int)set);
			return true;
		}catch(OleException const& exception){
			throw;
		}
	}
	return false;
}
		
ExcelWorkbook ExcelApp::NewWorkbook(){ //Create new Workbook and add it to actual excel Running method
	if(this->ExcelIsStarted){
		workbooks.Add(ExcelWorkbook(*this,GetAttribute(GetAttribute("Workbooks"),"Add"))).ResolveSheet();
		return workbooks[workbooks.GetCount()-1];
	}
	return ExcelWorkbook();
}

ExcelWorkbook ExcelApp::OpenWorkbook(Upp::String name){//Find and Open Workbook by FilePath
	if( !FileExists(name.ToStd().c_str())) {
      return ExcelWorkbook();
    }
    workbooks.Add(ExcelWorkbook(*this,ExecuteMethode(GetAttribute(L"Workbooks"),L"Open",1,AllocateString(name)))).ResolveSheet();
    return workbooks[workbooks.GetCount()-1];
}

bool ExcelApp::FindWorkbook(Upp::String name){
	for(ExcelWorkbook &wb : workbooks){
		String wbName = name.Right(name.GetCount() - (name.ReverseFind("\\") +1));
		if(wbName.Compare(wb.Name()) == 0){
			return true;
		}
	}
	return false;
}

ExcelWorkbook ExcelApp::FindOrOpenWorkBook(Upp::String name){//Look at current openned workbook and open it if not open
	for(ExcelWorkbook &wb : workbooks){
		String wbName = name.Right(name.GetCount() - (name.ReverseFind("\\") +1));
		if(wbName.Compare(wb.Name()) == 0){
			wb.ResolveSheet();
			return wb;
		}
	}
	return this->OpenWorkbook(name);
}

ExcelWorkbook ExcelApp::Workbooks(int index){//Allow to retrieve workbook by is index 
	if(this->ExcelIsStarted && workbooks.GetCount() > index){
		return workbooks[index];
	}
	return ExcelWorkbook();
}

ExcelWorkbook ExcelApp::Workbooks(Upp::String name){//Allow to retrieve workbook by is name
	if(this->ExcelIsStarted){
		for(int i = 0; i< workbooks.GetCount(); i++){
			if (BSTRtoString(workbooks[i].GetAttribute("Name").bstrVal).Compare(name) ==0){
				return workbooks[i];
			}
		}
	}
	return ExcelWorkbook();
}

int ExcelApp::GetNumberOfWorkbook(){//Return number of workbook currently openned on this excel App
	return workbooks.GetCount();
}
	
bool ExcelApp::RemoveAWorkbookFromVector(ExcelWorkbook* wb){// remove workbook from vector
	bool trouver = false;
	int i =0;
	for(i= 0; i < workbooks.GetCount(); i++){
		// Cout() << wb <<  ":" << &workbooks[i] <<"\n";
		if( wb == &workbooks[i]){
			trouver = true;
			break;
		}
	}
	if(trouver) workbooks.Remove(i);
	return trouver;
}

bool ExcelApp::ResolveWorkbook(){//Function that calculate all the workbook on openned Excel
	int nbrworkbook = this->GetAttribute(this->GetAttribute("Workbooks"), "Count").intVal;
	for(int i = 0; i < nbrworkbook; i++){
		workbooks.Add(ExcelWorkbook(*this, this->GetAttribute("Workbooks",1,AllocateInt(i +1))));
	}
	for(ExcelWorkbook &w:workbooks){
		w.ResolveSheet();
	}
	return true;
}
		
/************************************************************************************************************************/
ExcelApp*const ExcelWorkbook::GetParent()const{//Getter on parent pointer
	return parent;
}
const Vector<ExcelSheet>& ExcelWorkbook::GetVector()const{
	return this->sheets;
}

const bool ExcelWorkbook::GetOpen()const
{
	return this->isOpenned;
}

ExcelWorkbook::ExcelWorkbook(){//Classic constructor
}

ExcelWorkbook::~ExcelWorkbook(){//Classic destructor
}

ExcelWorkbook::ExcelWorkbook(const ExcelWorkbook& e){ //Copy constructor.
	this->AppObj = e.AppObj;
	this->parent = e.GetParent();
	this->isOpenned = true;
	this->sheets = Vector<ExcelSheet>(e.GetVector(),e.GetVector().GetCount());
}

ExcelWorkbook::ExcelWorkbook(ExcelApp& myApp, VARIANT appObj){//Basic constructor
	this->AppObj = appObj;
	this->parent = &myApp;
	this->isOpenned = true;
}

ExcelWorkbook& ExcelWorkbook::operator=(const ExcelWorkbook &wb)
{
	this->AppObj = wb.AppObj;
	this->parent = wb.GetParent();
	this->isOpenned = true;
	this->sheets = Vector<ExcelSheet>(wb.GetVector(),wb.GetVector().GetCount());
	return *this;
}

bool ExcelWorkbook::Save(){ //Save current workbook
	if(this->isOpenned){
		try{
			ExecuteMethode("Save",0);
			return true;
		}catch(OleException const& exception){
			throw;
		}
	}
	return false;
}

bool ExcelWorkbook::SaveAs(Upp::String filePath){//Save current workbook at filePath
	if(this->isOpenned){
		try{
			ExecuteMethode("SaveAs",1,AllocateString(filePath));
			return true;
		}catch(OleException const& exception){
			throw;
		}
	}
	return false;
}

bool ExcelWorkbook::Close(){//Close current workbook
	if(this->isOpenned){
		try{
			parent->RemoveAWorkbookFromVector(this);
			ExecuteMethode("Close",1,AllocateInt(0));
			return true;
		}catch(OleException const& exception){
			throw;
		}
	}
	return false;
}
Upp::String ExcelWorkbook::Name(){ //Return wb name
	if(this->isOpenned){
		try{
			return BSTRtoString(GetAttribute("Name").bstrVal);
		}catch(OleException const& exception){
			throw;
		}
	}
	return "";
}

Upp::String ExcelWorkbook::Path(){ //Return wb path
	if(this->isOpenned){
		try{
			return BSTRtoString(GetAttribute("Path").bstrVal) +"\\"+ Name();
		}catch(OleException const& exception){
			throw;
		}
	}
	return "";
}

bool ExcelWorkbook::isReadOnly(){//Return true if the workbook is readOnly
	if(this->isOpenned){
		return (bool)GetAttribute("ReadOnly").lVal;
	}
	return false;
}

ExcelSheet ExcelWorkbook::Sheets(int index){//Allow to retrieve worksheet by is index 
	if(this->isOpenned && sheets.GetCount() > index){
		return sheets[index];
	}
	return ExcelSheet();
}

ExcelSheet ExcelWorkbook::Sheets(Upp::String name){//Allow to retrieve worksheet by is name
	if(this->isOpenned){
		for(int i = 0; i< sheets.GetCount(); i++){
			if (BSTRtoString(sheets[i].GetAttribute("Name").bstrVal).Compare(name) ==0){
				return sheets[i];
			}
		}
	}
	return ExcelSheet();
}

ExcelSheet ExcelWorkbook::AddSheet(){ //Create new Sheet with default Name
	if(this->isOpenned){
		return sheets.Add(ExcelSheet(*this,GetAttribute(GetAttribute("Sheets"),"Add")));
	}
	return ExcelSheet();
}

ExcelSheet ExcelWorkbook::AddSheet(Upp::String sheetName){ //Create new Sheet with defined name 
	if(this->isOpenned){
		try {
			this->sheets.Add(ExcelSheet(*this,GetAttribute(GetAttribute("Sheets"),"Add")));
			this->sheets[this->sheets.GetCount()-1].SetName(sheetName);
			return this->sheets[this->sheets.GetCount()-1];
		} catch(OleException const& exception) {
			throw;
		}
		return ExcelSheet();
	}
	return ExcelSheet();
}

bool ExcelWorkbook::ResolveSheet(){//Function that calculate all the sheet on openned workbook
	int nbrSheet = this->GetAttribute( this->GetAttribute("Sheets"),"Count").intVal;
	for(int i = 0; i < nbrSheet; i++){
		sheets.Add(ExcelSheet(*this, this->GetAttribute("Sheets",1,AllocateInt(i +1))));
	}
	return true;
}
/************************************************************************************************************************/
ExcelWorkbook*const ExcelSheet::GetParent()const{//Getter on parent pointer
	return parent;
}

ExcelSheet::ExcelSheet(){//Classic constructor
}

ExcelSheet::~ExcelSheet(){//Classic desctructor
}

ExcelSheet::ExcelSheet(ExcelWorkbook& parent, VARIANT appObj){//Classic constructor
	this->AppObj = appObj;
	this->parent = &parent;
}

ExcelSheet& ExcelSheet::operator=(const ExcelSheet &ws)
{
	this->AppObj = ws.AppObj;
	this->parent = ws.GetParent();
	return *this;
}

bool ExcelSheet::SetName(Upp::String sheetName){//Redefine name of sheet
	try{
		return this->SetAttribute(this->AppObj,"Name",sheetName);
	}catch(OleException const& exception){
		throw;
	}
	return false;
}
int ExcelSheet::GetLastRow(Upp::String Colonne){//Retrieve last row of a colonne
	char range[10];
	ltoa(this->GetRowNumberOfMySheet(),range,10);
	Upp::String finalRange = Colonne + "1:"+ Colonne + Upp::String(range);
	/*
	Here we use some excel const
	xlDown		-4121	Down.
	xlToLeft	-4159	To left.
	xlToRight	-4161	To right.
	xlUp		-4162	Up.
	*/
	return this->GetAttribute(this->GetAttribute(this->GetAttribute("Range",1,AllocateString(finalRange)),"End",1,AllocateInt(-4121)),L"Row").intVal;
}

int  ExcelSheet::GetLastColumn(){// Retrieve the last Column
	//TODO
}

int ExcelSheet::GetRowNumberOfMySheet(){//Retrieve the max number generated by excel. It's usefull to make a huge range that wrap entire sheet
	return this->GetAttribute(this->GetAttribute("Rows"),"Count").intVal;
}

ExcelRange ExcelSheet::Range(Upp::String range){//Return a Range
 	return ExcelRange(*this,this->GetAttribute("Range",1,AllocateString(range)),range);
}

ExcelRange ExcelSheet::GetCurrentRegion(){//Return ExcelRange that's represente the entire active range of the actual sheet
	return ExcelRange(*this,GetAttribute(GetAttribute("Range",1,AllocateString(L"A1:A1")),"CurrentRegion"),"");
}

ExcelCell ExcelSheet::Cells(int ligne, int colonne){//Return a Cells
	char range[50];
	IndToStr(ligne,colonne,range);
	return ExcelCell(GetAttribute(GetAttribute(GetAttribute("Range",1,AllocateString(L"A1:A1")),"CurrentRegion"),"Cells",2,  AllocateInt(colonne),AllocateInt(ligne)));
}
/************************************************************************************************************************/
ExcelSheet*const ExcelRange::GetParent()const{//Getter on parent pointer
	return parent;
}

ExcelRange::ExcelRange(){
}

ExcelRange::~ExcelRange(){
}

ExcelRange::ExcelRange(ExcelSheet &parent,VARIANT appObj){//allow to create ExcelRange on current Variant 
	this->AppObj = appObj;
	this->parent = &parent;
}
																	   
ExcelRange::ExcelRange(ExcelSheet &parent,VARIANT appObj,Upp::String range){//This constructor allow user to pass the range used to get this object. 
	this->AppObj = appObj;													//It's very important if you want to be able tu use every function that 
	this->parent = &parent;													//do job on vector or return vector of Cells
	this->range = range;
}
														   
Upp::String ExcelRange::GetTheRange(){//Return the range used to get the Item, it can be empty
	return this->range;	
}

ExcelCell ExcelRange::Cells(int ligne, int colonne){//Return a Cells by is column and row
	char range[50];
	IndToStr(ligne,colonne,range);
	return ExcelCell(*this,GetAttribute("Cells",2, AllocateInt(colonne),AllocateInt(ligne)));
}

/*
// From NOW you must have a ExcelRange where Upp::String range is initialized
*/

Upp::Vector<ExcelCell> ExcelRange::Value(){//Return every  Cells on a Vector of Cells
	Upp::Vector<ExcelCell> allTheCells;
	if (!this->GetTheRange().GetCount() < 1){
		if( this->GetTheRange().Find(":") != -1){
			Upp::String debut = this->GetTheRange().Left(this->GetTheRange().Find(":"))	;
			Upp::String fin =  this->GetTheRange().Right(this->GetTheRange().GetCount() - (this->GetTheRange().Find(":")+1));
			int lDebut = ExtractRow(debut);
			lDebut--;
			int lFin = ExtractRow(fin);
			int cDebut = ColStrToInt(debut);
			cDebut--;
			int cFin = ColStrToInt(fin);
			for (int c = 1; c <= cFin -cDebut; c++){
				for(int l = 1; l <= lFin - lDebut; l++){
					allTheCells.Add(ExcelCell(*this,GetAttribute("Cells",2, AllocateInt(c),AllocateInt(l))));
				}
			}
		}
		else
		{
			int ligne = ExtractRow(this->GetTheRange());
			int colonne = ColStrToInt(this->GetTheRange());
			allTheCells.Add(ExcelCell(*this,GetAttribute("Cells",2,AllocateInt(1), AllocateInt(1))));
		}

	}
	return allTheCells;
}

bool ExcelRange::Value(Upp::String value){ //set this value to every cells of the range
	if (!this->GetTheRange().GetCount() < 1){
		Upp::Vector<ExcelCell> myVector = this->Value();
		for(int i = 0; i < myVector.GetCount(); i++){
			myVector[i].Value(value);
		}
		return true;
	}
	return false;
}

bool ExcelRange::Value(int value){ //set this value to every cells of the range
	if (!this->GetTheRange().GetCount() < 1){
		Upp::Vector<ExcelCell> myVector = this->Value();
		for(int i = 0; i < myVector.GetCount(); i++){
			myVector[i].Value(value);
		}
		return true;
	}
	return false;
}
/************************************************************************************************************************/
ExcelRange*const ExcelCell::GetParent()const{//Getter on parent pointer 
	return parent;
}

ExcelCell::~ExcelCell(){
}

ExcelCell::ExcelCell(ExcelRange &parent,VARIANT appObj){//Classic constructor
	this->AppObj = appObj;
	this->parent = &parent;
}

ExcelCell::ExcelCell(VARIANT appObj){//Constructor if parent not important (Some ExcelSheet function directly return cells without range setted)
	this->AppObj = appObj;
}
/*
	Here we must add every method a cell could land 
*/
Upp::String ExcelCell::Value(){ //Get the Value of the cells
	switch(GetAttribute("Value").vt) {
		case VT_BSTR: {
			return BSTRtoString(GetAttribute("Value").bstrVal);
		}
		case VT_R8: {
			return StringWOZ(String(std::to_string(GetAttribute("Value").dblVal)));
		}
		case VT_DATE: {
			BSTR variant_date_string;
			if (FAILED(VarBstrFromDate(GetAttribute("Value").date, 0, 0, &variant_date_string)))
			    Cout() <<"Error"<<"\n";
			String madate = BSTRtoString(variant_date_string);
			return madate;
		}
		case VT_EMPTY : {
			return "";
		}
		default:
			throw OleException(2,"UNKNOWN CELL.VALUE VARTYPE: " + String(std::to_string(GetAttribute("Value").vt)), 0);
	}
}

bool ExcelCell::Value(Upp::String value){//Set value of a Cell
	return SetAttribute("Value",value);
}

bool ExcelCell::Value(int value){//Set value of Cells
	return SetAttribute("Value",value);
}
/************************************************************************************************************************/