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
#include "..\util\MSXMLWorks.h"
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
//#include "ThreadSelect.h"


typedef std::map<String, String> EnvVariables;

// Режим экспорта данных
typedef enum _EXPORTMODE {
    EM_UNDEFINITE = 0,
    EM_PROCEDURE,   // Выполнить как процедуру
    EM_EXCEL_BLANK,     // Экспорт в пустой файл MS Excel
    EM_EXCEL_TEMPLATE,  // Экспорт в шаблон MS Excel
    EM_DBASE4_FILE,     // Экспорт в DBF
    EM_WORD_TEMPLATE    // Экспорт в шаблон MS Word
} EXPORTMODE;

// Структура элемента List в параметрах пользователя
class TParamlistItem {
public:
    AnsiString value;       // Фактическое значение
    AnsiString label;       // Отображаемое значение
    AnsiString visible;     // Безусловный флаг видимости
    AnsiString visibleif;   // Условие, при котором элемент отображается
    bool visibleflg;        // Текущее состояние видимости с учетом visible и visibleif
};

// Структура для хранения параметров запроса
class TParamRecord
{
public:
    AnsiString type;    // Тип
    AnsiString name;    // Внутреннее имя парамера
    AnsiString value;   // Внутреннее? значение параметра
    AnsiString value_src;   // Внутреннее (исходное) значение параметра
    AnsiString label;   // Отображаемое имя парамера
    AnsiString display; // Отображаемое значение параметра
    AnsiString format;  // Формат вывода данных
    AnsiString dbindex; // Индекс базы данных для загрузки списка значений (если в xml src )
    AnsiString src;     // Индекс базы данных для загрузки списка значений (если в xml src )
    AnsiString visible;         // Флаг
    bool deleteifflg;   // Флаг удалять блок если value параметра равен значени deleteifval
    AnsiString deleteifvalue;  // Флаг удалять блок если value параметра равен значени deleteifval
    //std::vector <TParamlistItem> variables;   // Список возможных значений
    std::vector <TParamlistItem> listitem;   // Список значений (для list и variables)
    AnsiString visibleif;   // Зависимость
    AnsiString disableif;   // Зависимость
    AnsiString parent;      // Имя родительского параметра (пока не доработано)

    bool visibleflg;    // вычисляемый параметр
    AnsiString mask;    // Маска ввода
};

/*
class TParamRecordCtrl : public TParamRecord
{
public:
    void Show(TParamRecordCtrl *paramRecord);
    //void SetType(String Type);
    TObject* Control;
}; */

// Структура для хранения параметров поля (столбца) DBASE
typedef struct {    // Для описания структуры dbf-файла
    String type;    // Тип fieldtype is a single character [C,D,F,L,M,N]
    String name;    // Имя поля (до 10 символов).
    int length;         // Длина поля
    int decimals;       // Длина десятичной части
    // Character 	 1-255
    // Date	  8
    // Logical	  1
    // Memo	  10
    // Numeric	1-30
    // Decimals is 0 for non-numeric, number of decimals for numeric.
} DBASEFIELD;

// Структура для хранения параметров поля (столбца) MS Excel
typedef struct {    // Для описания формата ячеек в Excel
    AnsiString format;      // Формат ячейки в Excel
    AnsiString name;        // Имя поля
    //int title_rows;       // Высота заголовка в строках
    int width;              // Ширина столбца
    int bwraptext;          // Флаг перенос по словам
} EXCELFIELD;

// Структура для хранения параметров экспорта в MS Excel
typedef struct {
    AnsiString id;
    AnsiString label;
    //bool fDefault;
    AnsiString template_name;       // Имя файла шаблона Excel
    AnsiString title_label;         // Строка - выводимая в качестве заголовка в отчете Excel (перенести в отдельную структуру)
    int title_height;               // Высота заголовка в строках  (перенести в отдельную структуру)
    std::vector<EXCELFIELD> Fields;     // Список полей для экспорта в файл MS Excel
    AnsiString table_range_name;        // Имя диапазона табличной части (при выводе в шаблон)
    bool fUnbounded;                    // Флаг того, что диапазон table_range_name будет увеличен, в соответствии с количеством записей в источнике данных
} EXPORT_PARAMS_EXCEL;

// Структура для хранения параметров экспорта в MS Word
typedef struct {
    AnsiString id;
    AnsiString label;
    //bool fDefault;
    AnsiString template_name;   // Имя файла шаблона MS Word
    int page_per_doc;           // Количество страниц на документ MS Word
    AnsiString filter_main_field;      // Имя поля из основного запроса для сравнения со значением поля word_filter_sec_field
    AnsiString filter_sec_field;       // Имя поля из вспомогательного запроса (см. word_filter_main_field)
    AnsiString filter_infix_sec_field; // Имя поля из вспомогательного запроса, значение которого будет присоединено к имени результирующего файла
} EXPORT_PARAMS_WORD;

// Структура для хранения параметров режима экспорта - Выполнить
typedef struct {
    AnsiString id;
} EXPORT_PARAMS_EXECUTE;

// Структура для хранения параметров экспорта в DBF
typedef struct {    // Для описания формата ячеек в Excel
    AnsiString id;
    AnsiString label;
    //bool fDefault;
    bool fAllowUnassignedFields;
    std::vector<DBASEFIELD> Fields;    // Список полей для экспрта в файл DBF
} EXPORT_PARAMS_DBASE;


typedef std::vector<TParamRecord> QueryVariables;

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

    EnvVariables Variables;         // Переменные

    QueryVariables UserParams;    // Задаваемые параметры к запрос

    bool fExcelFile;      // Флаг Excel в память
    //int fExcelFile;     // Флаг Excel в файл
    bool fWordFile;     // Флаг Word
    bool fDbfFile;       // Флаг Dbf в файл
};

/*
class TQueryItemCtrl : public TQueryItem
{
public:
    TObject* Control;
};   */

#include "ThreadSelect.h"

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
private:	// User declarations
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

    String __fastcall GetValue(String value);
    AnsiString GetDefinedValue(AnsiString value);       // Оставлено для совместимости. В последущем полностью заменить GetValue
    void __fastcall DoExport(THREADOPTIONS* threadopt);
    bool __fastcall CheckLock(int dbindex);
    void __fastcall Run(EXPORTMODE ExportMode, int Tag = 0);

    bool CheckCondition(AnsiString condition);
    void InitEnvVariables(); // Инициализации переменных среды

    String ReplaceVariables(EnvVariables &variables, const String& Text);


    int __fastcall DataSetToQueryList(TOraQuery* oraquery, std::vector<TQueryItem>& query_list, std::vector<TTabItem>& tab_list);
    TColor __fastcall ColorByIndex(int index);     // Возвращает цвет по индексу

    void __fastcall AddEnvVariable(const String& name, const String& value);
    void __fastcall AddSystemVariable(const String& name, const String& value);


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

    std::vector<String> DangerWords;    // Вектор строк, которые нужно убрать из строки запроса (truncate, insert, update...)
    std::vector<TTabItem> TabList;       // Вектор разделов
    std::vector<TQueryItem> QueryList;   // Вектор строк заросов


    std::vector<String> m_env_func;     // Вектор строк, которые нужно убрать из строки запроса (truncate, insert, update...)

    std::vector<TColor> m_vTabColor;  // Вектор цветов вкладок
    unsigned int m_ColorIndex;


    AnsiString NumEdit_TextOld;     // Переменные для функционирования NumEdit1
    int NumEdit_SelStartOld;        // Переменные для функционирования NumEdit1
    bool NumEdit_bUseSign;          // Переменные для функционирования NumEdit1
    bool NumEdit_bUseDot;           // Переменные для функционирования NumEdit1

    std::map<String, TObject*> paramControls;

    EnvVariables envVariables;      // Переменные среды

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
