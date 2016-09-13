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

// ����� �������� ������
typedef enum _EXPORTMODE {
    EM_UNDEFINITE = 0,
    EM_PROCEDURE,   // ��������� ��� ���������
    EM_EXCEL_BLANK,     // ������� � ������ ���� MS Excel
    EM_EXCEL_TEMPLATE,  // ������� � ������ MS Excel
    EM_DBASE4_FILE,     // ������� � DBF
    EM_WORD_TEMPLATE    // ������� � ������ MS Word
} EXPORTMODE;

// ��������� �������� List � ���������� ������������
class TParamlistItem {
public:
    AnsiString value;       // ����������� ��������
    AnsiString label;       // ������������ ��������
    AnsiString visible;     // ����������� ���� ���������
    AnsiString visibleif;   // �������, ��� ������� ������� ������������
    bool visibleflg;        // ������� ��������� ��������� � ������ visible � visibleif
};

// ��������� ��� �������� ���������� �������
class TParamRecord
{
public:
    AnsiString type;    // ���
    AnsiString name;    // ���������� ��� ��������
    AnsiString value;   // ����������? �������� ���������
    AnsiString value_src;   // ���������� (��������) �������� ���������
    AnsiString label;   // ������������ ��� ��������
    AnsiString display; // ������������ �������� ���������
    AnsiString format;  // ������ ������ ������
    AnsiString dbindex; // ������ ���� ������ ��� �������� ������ �������� (���� � xml src )
    AnsiString src;     // ������ ���� ������ ��� �������� ������ �������� (���� � xml src )
    AnsiString visible;         // ����
    bool deleteifflg;   // ���� ������� ���� ���� value ��������� ����� ������� deleteifval
    AnsiString deleteifvalue;  // ���� ������� ���� ���� value ��������� ����� ������� deleteifval
    //std::vector <TParamlistItem> variables;   // ������ ��������� ��������
    std::vector <TParamlistItem> listitem;   // ������ �������� (��� list � variables)
    AnsiString visibleif;   // �����������
    AnsiString disableif;   // �����������
    AnsiString parent;      // ��� ������������� ��������� (���� �� ����������)

    bool visibleflg;    // ����������� ��������
    AnsiString mask;    // ����� �����
};

/*
class TParamRecordCtrl : public TParamRecord
{
public:
    void Show(TParamRecordCtrl *paramRecord);
    //void SetType(String Type);
    TObject* Control;
}; */

// ��������� ��� �������� ���������� ���� (�������) DBASE
typedef struct {    // ��� �������� ��������� dbf-�����
    String type;    // ��� fieldtype is a single character [C,D,F,L,M,N]
    String name;    // ��� ���� (�� 10 ��������).
    int length;         // ����� ����
    int decimals;       // ����� ���������� �����
    // Character 	 1-255
    // Date	  8
    // Logical	  1
    // Memo	  10
    // Numeric	1-30
    // Decimals is 0 for non-numeric, number of decimals for numeric.
} DBASEFIELD;

// ��������� ��� �������� ���������� ���� (�������) MS Excel
typedef struct {    // ��� �������� ������� ����� � Excel
    AnsiString format;      // ������ ������ � Excel
    AnsiString name;        // ��� ����
    //int title_rows;       // ������ ��������� � �������
    int width;              // ������ �������
    int bwraptext;          // ���� ������� �� ������
} EXCELFIELD;

// ��������� ��� �������� ���������� �������� � MS Excel
typedef struct {
    AnsiString id;
    AnsiString label;
    //bool fDefault;
    AnsiString template_name;       // ��� ����� ������� Excel
    AnsiString title_label;         // ������ - ��������� � �������� ��������� � ������ Excel (��������� � ��������� ���������)
    int title_height;               // ������ ��������� � �������  (��������� � ��������� ���������)
    std::vector<EXCELFIELD> Fields;     // ������ ����� ��� �������� � ���� MS Excel
    AnsiString table_range_name;        // ��� ��������� ��������� ����� (��� ������ � ������)
    bool fUnbounded;                    // ���� ����, ��� �������� table_range_name ����� ��������, � ������������ � ����������� ������� � ��������� ������
} EXPORT_PARAMS_EXCEL;

// ��������� ��� �������� ���������� �������� � MS Word
typedef struct {
    AnsiString id;
    AnsiString label;
    //bool fDefault;
    AnsiString template_name;   // ��� ����� ������� MS Word
    int page_per_doc;           // ���������� ������� �� �������� MS Word
    AnsiString filter_main_field;      // ��� ���� �� ��������� ������� ��� ��������� �� ��������� ���� word_filter_sec_field
    AnsiString filter_sec_field;       // ��� ���� �� ���������������� ������� (��. word_filter_main_field)
    AnsiString filter_infix_sec_field; // ��� ���� �� ���������������� �������, �������� �������� ����� ������������ � ����� ��������������� �����
} EXPORT_PARAMS_WORD;

// ��������� ��� �������� ���������� ������ �������� - ���������
typedef struct {
    AnsiString id;
} EXPORT_PARAMS_EXECUTE;

// ��������� ��� �������� ���������� �������� � DBF
typedef struct {    // ��� �������� ������� ����� � Excel
    AnsiString id;
    AnsiString label;
    //bool fDefault;
    bool fAllowUnassignedFields;
    std::vector<DBASEFIELD> Fields;    // ������ ����� ��� ������� � ���� DBF
} EXPORT_PARAMS_DBASE;


typedef std::vector<TParamRecord> QueryVariables;

// ��������� ��� ���������� ������� � ���������� � ����
class TQueryItem {
public:
    AnsiString tabname;     // ������������ ������
    AnsiString taborder;    // ���������� ����� �������
    AnsiString queryid;     // id �������
    AnsiString querytext;   // ����� �������
    AnsiString querytext2;  // ����� ������� ������� (������������ � ������� � MS Word)
    AnsiString queryname;   // ������������ �������
    AnsiString dbname;      // ������ ���� ������
    AnsiString dbname2;     // ������ ���� ������ ��� ������� ������� (������������ � ������� � MS Word)
    AnsiString sortorder;   // ������� ����������
    AnsiString spr_task_sql2excel_id;
    AnsiString fieldslist;  // ������ - �������� ������� (����������� � �������)

    EXPORTMODE DefaultExportType;   // ��� ������ ��� �������� "�� ���������"
    //int ExportFieldIndex;           // ������ ������ ��� ��������;

    AnsiString exportparam_id;


    EXPORT_PARAMS_EXCEL param_excel;
    EXPORT_PARAMS_DBASE param_dbase;
    //std::vector<PARAM_EXCEL> param_excel;
    EXPORT_PARAMS_WORD param_word;
    EXPORT_PARAMS_EXECUTE param_execute;

    EnvVariables Variables;         // ����������

    QueryVariables UserParams;    // ���������� ��������� � ������

    bool fExcelFile;      // ���� Excel � ������
    //int fExcelFile;     // ���� Excel � ����
    bool fWordFile;     // ���� Word
    bool fDbfFile;       // ���� Dbf � ����
};

/*
class TQueryItemCtrl : public TQueryItem
{
public:
    TObject* Control;
};   */

#include "ThreadSelect.h"

typedef std::vector<TQueryItem*> QueryItemList;

class TTabItem {    // ���������, ��� �������� ��������� �������
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
    bool __fastcall Auth();         // �����������
    bool __fastcall PrepareForm();  // ����������
    void PrepareQuery();        // ���������
    void PrepareTabs();         // �������� �������
    void FillFieldsLB();        // ���������� ������ �������� � ������������ � �������� ��������
    void FillParametersLV();    // ���������� ������ ����������
    void ParseUserParamsStr(AnsiString ParamStr, TQueryItem* queryitem);  // ������ xml � std::vector<TParamRecord>
    void ParseExportParamsStr(AnsiString ParamStr, TQueryItem* queryitem);  // ������ xml � std::vector<TParamRecord>
    String GetValueFromSQL(String SQLText, String dbindex);

    String GetSQL(const String& SQLText, QueryVariables* queryParams = NULL) const;   // ����������� ������ ������� � ������������ �������� � �.�. ��� ����������

    String __fastcall GetValue(String value);
    AnsiString GetDefinedValue(AnsiString value);       // ��������� ��� �������������. � ���������� ��������� �������� GetValue
    void __fastcall DoExport(THREADOPTIONS* threadopt);
    bool __fastcall CheckLock(int dbindex);
    void __fastcall Run(EXPORTMODE ExportMode, int Tag = 0);

    bool CheckCondition(AnsiString condition);
    void InitEnvVariables(); // ������������� ���������� �����

    String ReplaceVariables(EnvVariables &variables, const String& Text);


    int __fastcall DataSetToQueryList(TOraQuery* oraquery, std::vector<TQueryItem>& query_list, std::vector<TTabItem>& tab_list);
    TColor __fastcall ColorByIndex(int index);     // ���������� ���� �� �������

    void __fastcall AddEnvVariable(const String& name, const String& value);
    void __fastcall AddSystemVariable(const String& name, const String& value);


    TOdacUtilLog* OdacLog;
    std::vector<TOraSession*> m_sessions;    // ������ �����, ������� ����� ������ �� ������ ������� (truncate, insert, update...)

    ThreadSelect *ts;   // ����� ������������ ������
    bool bAdmin;            // ����, ����������� �� ��, ��� ���������������� ������������ - �������������
    AnsiString AppPath;     // ���� � ������������ ����� ���������
    AnsiString Username;    // ��� ������������, ��������������� � ���������
    double TotalTime;       // ������ ��� ������� ������� ������ ������ ������������ ������
    //double TerminateTime;   // ������ ��� ���������� ������ ������������ ������
    TObject *CurrentDinamicControl;    // ������� ���������� ��� �������������� �������� � ������ ���������� ������������ � ��������� �����
    TQueryItem* CurrentQueryItem; 	    // ������� ��������� ������

    std::vector<String> DangerWords;    // ������ �����, ������� ����� ������ �� ������ ������� (truncate, insert, update...)
    std::vector<TTabItem> TabList;       // ������ ��������
    std::vector<TQueryItem> QueryList;   // ������ ����� �������


    std::vector<String> m_env_func;     // ������ �����, ������� ����� ������ �� ������ ������� (truncate, insert, update...)

    std::vector<TColor> m_vTabColor;  // ������ ������ �������
    unsigned int m_ColorIndex;


    AnsiString NumEdit_TextOld;     // ���������� ��� ���������������� NumEdit1
    int NumEdit_SelStartOld;        // ���������� ��� ���������������� NumEdit1
    bool NumEdit_bUseSign;          // ���������� ��� ���������������� NumEdit1
    bool NumEdit_bUseDot;           // ���������� ��� ���������������� NumEdit1

    std::map<String, TObject*> paramControls;

    EnvVariables envVariables;      // ���������� �����

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
