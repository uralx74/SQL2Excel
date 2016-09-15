/*
  Программа для выгрузки данных из Баз Данных в:
  1. MS Excel
  2. MS Word
  3. DBF

  Автор: Инженер-программист 2 кат. ПАО ЦФ ЧЭС
  В.С.Овчинников


  Для следующих версий:
  1. Разработать механизм для возможности экспорта в одинаковые форматы, но с разными
  настройками.
  2. Разработать механизм для возможности запуска нескольких потоков экспорта.
  3. Разработать механизм для возможности прерывания потока пользователем.
  4. Разработать механизм для возможности запуска программы по расписанию с возможностью
  отправки готового отчета по почте.

*/
//---------------------------------------------------------------------------
#include <vcl.h>
#pragma hdrstop
#include "fmain.h"

USERES("SQL2Excel.res");
USEFORM("FMain.cpp", Form1);
USEUNIT("ThreadSelect.cpp");
USEFORM("FormWait.cpp", Form_Wait);
USEFORM("FShowQuery.cpp", FormShowQuery);
USEFORM("..\util\FormLogin\FormLogin.cpp", LoginForm);
USEUNIT("..\util\MSExcelWorks.cpp");
USEUNIT("..\util\MSXMLWorks.cpp");
USEUNIT("..\util\MSWordWorks.cpp");
USEUNIT("..\util\CommandLine.cpp");
USEUNIT("Parameter.cpp");
USEUNIT("Variables.cpp");
USEUNIT("ParameterizedText.cpp");
//---------------------------------------------------------------------------
WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int)
{
/*    map  <String,String> startparams;
    startparams["-username"] = "";
    startparams["-password"] = "";
    startparams["-query"] = "";
    startparams["-mail"] = "";
    startparams["-file"] = "false";
    startparams["-failure"] = "";
    startparams["-subject"] = "";
    startparams["-param"] = "";

    map <String,int> argcount;
    argcount["-username"] = 1;
    argcount["-password"] = 1;
    argcount["-query"] = 1;
    argcount["-mail"] = 1;
    argcount["-file"] = 1;
    argcount["-failure"] = 0;
    argcount["-subject"] = 1;
    argcount["-param"] = 2;

    map <String,String> startparamsshort;
    startparamsshort["-u"] = "-username";
    startparamsshort["-p"] = "-password";
    startparamsshort["-q"] = "-query";
    startparamsshort["-m"] = "-mail";
    startparamsshort["-f"] = "-file";
    startparamsshort["-s"] = "-subject";

    bool bKeyAuto = false;
    int n = ParamCount();


    AnsiString paramname = "";

    for (int i = 1; i <= n; i++) {
        AnsiString sParamStr = Trim(ParamStr(i));
        if (paramname == "") {
            //if (startparams.find(sParamStr)!=startparams.end())
            if (startparams.count(sParamStr) > 0 )
                paramname = sParamStr;
            else if (startparamsshort.count(sParamStr) > 0)
                paramname = startparamsshort[sParamStr];

            if (paramname == "")
                continue;

            if (argcount[paramname] == 0) {
                paramname = "";
                startparams[paramname] = "true";
            }
        } else {
            argcount[paramname];
            startparams[paramname]=sParamStr;
            paramname = "";
        }
    }

/*    for(std::map<String,String>::iterator it = startparams.begin(); it != startparams.end(); ++it) {
        String s = it->first; //it->second << endl;///????? ?? ?????
    }
*/

  /*  AnsiString sParam = startparams["-username"];
    AnsiString sParam2 = startparams["-password"];                         */


        /*RunParameters rp;

    rp.AddParam("-username", "-u", 1);
    rp.AddParam("-password", "-p", 1);
    rp.AddParam("-query", "-q", 1);
    rp.AddParam("-mail", "-m", 1);
    rp.AddParam("-file", "-f", 1);
    rp.AddParam("-failure", 0);
    rp.AddParam("-subject", "-s", 1);
    rp.AddParam("-param", 2); 

    rp.ParseExecute(); */

        /*AnsiString ExecDir = ExtractFileDir(ParamStr(0));
    TIniFile *ini = new TIniFile(ExecDir + "\\record.data");
    TStringList *Sections = new TStringList;
    try {
        ini->ReadString("Section","Ident3","Def");
        ini->ReadSections(Sections);
    }
    __finally {
        delete Sections;
        delete ini;
    }*/
    /*BringWindowToTop(Form1->Handle);
    SetActiveWindow(Form1->Handle);
    ShowWindow(Application->Handle, SW_SHOWNA);
    SetForegroundWindow(Form1->Handle);         // На передний план    Кнопка на панели задач отжата*/



	try
	{
        Application->Initialize();
		Application->CreateForm(__classid(TForm1), &Form1);
         Application->CreateForm(__classid(TFormShowQuery), &FormShowQuery);
         Application->Run();
	}
	catch (Exception &exception)
	{
		Application->ShowException(&exception);
	}
	return 0;
}
//---------------------------------------------------------------------------



