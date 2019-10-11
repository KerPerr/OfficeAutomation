#ifndef _OfficeAutomation_Imsp_h_
#define _OfficeAutomation_Imsp_h_

/* 
Project created 10/09/2019 
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

class Imsp : public Ole , public Upp::Moveable<Imsp> {
	private:
		// VARIANT portant la frame actuelle
		VARIANT frame;
		// VARIANT portant le terminal
		VARIANT terminal;
		// Thread pour vérifier l'execution d'une macro
		Thread worker;
		// Bool to know if IMSP is started
		bool IsStarted;
		String path = "C:\\Users\\Public\\Documents\\Attachmate\\Reflection\\Built-Ins\\Sessions\\Session Reflection PROD.rd3x";
	public:
		// Lance une instance de IMSP
		bool Start(bool threading = false);
		// Recherche une instance IMSP
		bool Find();
		
		// Renvoie true si il existe une page suivante
		bool NextPage();
		// Renvoie true si il existe une page précédente
		bool PrevPage();
		
		// Attend que IMSP soit disponible
		void Wait(); 
		
		// Ecrit le texte dans IMSP
		void SetText(String text);
		// Envoie la commande dans IMSP
		void SetCmd(String cmd);
		
		// Recupère le texte
		String GetString(int row, int col, int len);
		
		Imsp();
		~Imsp();
};
#endif
