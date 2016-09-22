#ifndef FMainH
#define FMainH

//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ComCtrls.hpp>
#include "MemDS.hpp"
#include "Ora.hpp"
#include "..\MSExcelWorks.h"
#include "..\MSWordWorks.h"
#include "..\util\OleXml.h"
#include "..\util\taskutils.h"
#include "..\util\odacutils.h"
#include "..\util\formlogin\formlogin.h"
#include "..\util\appver.h"
#include "..\util\CommandLine.h"
#include <ExtCtrls.hpp>
#include <Buttons.hpp>
#include <AppEvnts.hpp>
#include <Menus.hpp>
#include "DBAccess.hpp"
#include <ActnList.hpp>
#include <Db.hpp>
#include "FormWait.h"
#include "FShowQuery.h"
#include "OdacVcl.hpp"
//#include <inifiles.hpp>

#include <ADODB.hpp>
#include "VirtualTable.hpp"
#include "Halcn6DB.hpp"
#include <Dialogs.hpp>
#include <ImgList.hpp>
#include <Mask.hpp>

#include "parameter.h"
#include "variables.h"
#include "ParameterizedText.h"
#include "Datatype.h"
#include "ThreadSelect.h"

class LvParameter: public Parameter
{
public:
    show();
};

Variables systemVariables;         // Переменные системные
Variables systemFunctions;         // Переменные системные

typedef std::vector<TParamRecord*> QueryVariables;
typedef std::vector<TParamRecord*>::iterator QueryVariablesIterator;

// Структура для сохранения Запроса и параметров к нему
class TQueryItem {
public:
    AnsiString tabname;     // Наименование раздел
    AnsiString taborder;    // Порядковый номер вкладки
    AnsiString queryid;     // id запроса
    AnsiString querytext;   // Текст запроса
    AnsiString querytext2;  // Текст второго запроса (используется в отчетах в MS Word)
    AnsiString queryname;   // Наименование запроса
    AnsiString dbname;      // Индекс базы данных
    AnsiString dbname2;     // Индекс базы данных для второго запроса (используется в отчетах в MS Word)
    AnsiString sortorder;   // Порядок сортировки
    AnsiString spr_task_sql2excel_id;
    AnsiString fieldslist;  // Строка - перечень записей (комментарий к запросу)

    EXPORTMODE DefaultExportType;   // Тип отчета для выгрузки "По умолчанию"
    //int ExportFieldIndex;           // Индекс отчета для выгрузки;

    AnsiString exportparam_id;


    EXPORT_PARAMS_EXCEL param_excel;
    EXPORT_PARAMS_DBASE param_dbase;
    //std::vector<PARAM_EXCEL> param_excel;
    EXPORT_PARAMS_WORD param_word;
    EXPORT_PARAMS_EXECUTE param_execute;

    QueryVariables UserParams;    // Задаваемые параметры к запрос

    bool fExcelFile;      // Флаг Excel в память
    //int fExcelFile;     // Флаг Excel в файл
    bool fWordFile;     // Флаг Word
    bool fDbfFile;       // Флаг Dbf в файл
};


typedef std::vector<TQueryItem*> QueryItemList;

class TTabItem {    // Структура, для хранения структуры вкладок
public:
    AnsiString name;
    QueryItemList queryitem;
};

//---------------------------------------------------------------------------
class TForm1 : public TForm
{
__published:	// IDE-managed Components
    TOraSession *EsaleSession;
	TStatusBar *StatusBar1;
	TSplitter *Splitter1;
	TPanel *Panel2;
	TBitBtn *BitBtn1;
	TApplicationEvents *ApplicationEvents1;
    TPanel *Panel1;
    TOraSession *CCBSession;
    TMainMenu *MainMenu1;
    TMenuItem *N1;
    TMenuItem *N3;
    TListBox *ListBox1;
    TTabControl *TabControl1;
    TMenuItem *N2;
    TMenuItem *CtrlC1;
    TTimer *Timer1;
    TOraSession *DBYYSession;
    TOraSession *DBWORK2Session;
    TListView *ParamsLV;
    TEdit *Edit1;
    TComboBox *ComboBox1;
    TDateTimePicker *DateTimePicker1;
    TActionList *ActionList1;
    TAction *ActionHelp;
    TAction *ActionCopyQuery;
    TAction *ActionDefaultRun;
    TMenuItem *N4;
    TMenuItem *N5;
    TMenuItem *N6;
    TBitBtn *BitBtn2;
    TPopupMenu *PopupMenu1;
    TMenuItem *Excel1;
    TMenuItem *DBF1;
    TSaveDialog *SaveDialog1;
    TAction *ActionShowMainQuery;
    TSpeedButton *SpeedButton1;
    TMenuItem *N7;
    TMenuItem *N8;
    TAction *ActionAboutApp;
    TOraQuery *CheckLockQuery;
    TImageList *ImageList1;
    TPanel *Panel3;
    TMenuItem *MSWord1;
    TMenuItem *Excel2;
    TAction *ActionExportExcelFile;
    TAction *ActionExportWordFile;
    TAction *ActionExportDbfFile;
    TAction *ActionShowSecondaryQuery;
    TMaskEdit *MaskEdit1;
    TEdit *NumEdit1;
    TAction *ActionExportExcelBlank;
    TMenuItem *ActionExportExcelMemory1;
    TMenuItem *N9;
    TAction *ActionAsProcedure;
    TAction *ActionApplictionExit;
    TAction *ActionShowEnvironment;
	void __fastcall FormCreate(TObject *Sender);
	void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
    void __fastcall ListBox1DrawItem(TWinControl *Control, int Index,
    TRect &Rect, TOwnerDrawState State);
    void __fastcall PageControl1DrawTab(TCustomTabControl *Control,
    int TabIndex, const TRect &Rect, bool Active);
    void __fastcall TabControl1Change(TObject *Sender);
    void __fastcall FormResize(TObject *Sender);
    void __fastcall Timer1Timer(TObject *Sender);
    void __fastcall DinamicControlExit(TObject *Sender);
    void __fastcall DinamicControlOnKeyDown(TObject *Sender, WORD &Key, TShiftState Shift);
    void __fastcall ParamsLVClick(TObject *Sender);
    void __fastcall ParamsLVAdvancedCustomDraw(TCustomListView *Sender,
    const TRect &ARect, TCustomDrawStage Stage, bool &DefaultDraw);
    void __fastcall ActionShowHelpExecute(TObject *Sender);
    void __fastcall ActionCopyQueryExecute(TObject *Sender);
    void __fastcall FormShow(TObject *Sender);
    void __fastcall EsaleSessionAfterDisconnect(TObject *Sender);
    void __fastcall BitBtn2Click(TObject *Sender);
    void __fastcall ActionShowMainQueryExecute(TObject *Sender);
    void __fastcall ActionAboutAppExecute(TObject *Sender);
    void __fastcall PopupMenu1Popup(TObject *Sender);
    void __fastcall ActionExportExcelFileExecute(TObject *Sender);
    void __fastcall ActionExportWordFileExecute(TObject *Sender);
    void __fastcall ActionExportDbfFileExecute(TObject *Sender);
    void __fastcall ListBox1MouseDown(TObject *Sender, TMouseButton Button,
          TShiftState Shift, int X, int Y);
    void __fastcall ListBox1Click(TObject *Sender);
    void __fastcall ActionShowSecondaryQueryExecute(TObject *Sender);
    void __fastcall NumEdit1Change(TObject *Sender);
    void __fastcall NumEdit1KeyPress(TObject *Sender, char &Key);
    void __fastcall ActionDefaultRunExecute(TObject *Sender);
    void __fastcall ActionExportExcelBlankExecute(TObject *Sender);
    void __fastcall ActionAsProcedureExecute(TObject *Sender);
    void __fastcall ActionApplictionExitExecute(TObject *Sender);
    void __fastcall ActionShowEnvironmentExecute(TObject *Sender);
private:	// User declarations

    //void showListEditor(const TParamRecord& param);


    void __fastcall OnEditParam();
    int __fastcall LoadQueryList();
    bool __fastcall Auth();         // Авторизация
    bool __fastcall PrepareForm();  // Подготовка
    void PrepareQuery();        // Поготовка
    void PrepareTabs();         // Создание вкладок
    void FillFieldsLB();        // Заполнение списка запросов в соответствии с выбраной вкладкой
    void FillParametersLV();    // Заполнение списка параметров
    void ParseUserParamsStr(AnsiString ParamStr, TQueryItem* queryitem);  // Разбор xml в std::vector<TParamRecord>
    void ParseExportParamsStr(AnsiString ParamStr, TQueryItem* queryitem);  // Разбор xml в std::vector<TParamRecord>
    String GetValueFromSQL(String SQLText, String dbindex);

    String GetSQL(const String& SQLText, QueryVariables* queryParams = NULL) const;   // Составление строки запроса с подстановкой значений и т.д. для выполнения
    bool TestParameters(const QueryVariables* queryParams);

    String __fastcall GetValue(String value);
   // AnsiString GetDefinedValue(AnsiString value);       // Оставлено для совместимости. В последущем полностью заменить GetValue
    void __fastcall DoExport(THREADOPTIONS* threadopt);
    bool __fastcall CheckLock(int dbindex);
    void __fastcall Run(EXPORTMODE ExportMode, int Tag = 0);

    bool CheckCondition(AnsiString condition);
    void InitEnvVariables(); // Инициализации переменных среды

    int __fastcall DataSetToQueryList(TOraQuery* oraquery, std::vector<TQueryItem>& query_list, std::vector<TTabItem>& tab_list);
    TColor __fastcall ColorByIndex(int index);     // Возвращает цвет по индексу


    TOdacUtilLog* OdacLog;
    std::vector<TOraSession*> m_sessions;    // Вектор строк, которые нужно убрать из строки запроса (truncate, insert, update...)

    ThreadSelect *ts;   // Поток формирования отчета
    bool bAdmin;            // Флаг, указывающий на то, что авторизовавшийся пользователь - администратор
    AnsiString AppPath;     // Путь к исполняемому файлу программы
    AnsiString Username;    // Имя пользователя, авторизованного в программе
    double TotalTime;       // Таймер для посчета времени работы потока формирования отчета
    //double TerminateTime;   // Таймер для прерывания потока формирования отчета
    TObject *CurrentDinamicControl;    // Элемент управления для редактирования значения в списке параметров отображаемый в настоящее время
    TQueryItem* CurrentQueryItem; 	    // Текущий выбранный запрос

    std::vector<String> dangerWords;    // Вектор строк, которые нужно убрать из строки запроса (truncate, insert, update...)
    std::vector<TTabItem> TabList;       // Вектор разделов
    std::vector<TQueryItem> QueryList;   // Вектор строк заросов


    //std::vector<String> m_env_func;     // Вектор строк, которые нужно убрать из строки запроса (truncate, insert, update...)

    std::vector<TColor> m_vTabColor;  // Вектор цветов вкладок
    unsigned int m_ColorIndex;


    AnsiString NumEdit_TextOld;     // Переменные для функционирования NumEdit1
    int NumEdit_SelStartOld;        // Переменные для функционирования NumEdit1
    bool NumEdit_bUseSign;          // Переменные для функционирования NumEdit1
    bool NumEdit_bUseDot;           // Переменные для функционирования NumEdit1

    std::map<String, TObject*> paramControls;











    //Variables customVariables;         // Переменные настраиваемые

    void __fastcall AddSystemVariable(const String& name, const String& value);
    void __fastcall AddCustomVariable(const String& name, const String& value);


public:
	__fastcall TForm1(TComponent* Owner);
    __fastcall ~TForm1();






    void __fastcall OnThread(int Status, AnsiString Message = "");
    void __fastcall OnThreadChangeStatus(int Status);
    void __fastcall OnThreadError(int Status);
    void __fastcall OnThreadSuccess(EXPORTMODE ExportMode, std::vector<String> vResultFiles);
    void __fastcall OnThreadSync(int Status);

};


//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif
