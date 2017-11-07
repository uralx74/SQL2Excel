//---------------------------------------------------------------------------

#ifndef ThreadSelectH
#define ThreadSelectH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include "Ora.hpp"
#include "OraDataTypeMap.hpp"
#include "math.h"

//#include "Halcn6DB.hpp"

#include "QueryItem.h"
#include "Parameter.h"
#include "..\util\odacutils.h"
#include "taskutils.h"
#include "DocumentWriter.h"

#include <DB.hpp>
#include "MemDS.hpp"


class TThreadSelectMessage
{
public:
    TThreadSelectMessage(unsigned int status, const AnsiString& message, std::vector<String> files);
    TThreadSelectMessage(unsigned int status, const AnsiString& message);
    unsigned int getStatus() { return _status; };
    AnsiString getMessage() { return _message; };
    std::vector<String> getFiles() { return _files; };

private:
    unsigned int _status;
    //AnsiString _message;
    AnsiString _message;
    std::vector<String> _files;   // ������ ���� ������ - �����������
};


//typedef enum _EXPORTMODE{AS_PROCEDURE = 0, TO_EXCEL_FILE, TO_EXCEL_MEMORY, TO_DBASE4_FILE, AS_PROCEDURE, TO_WORD_FILE, TO_WORD_MEMORY} EXPORTMODE;

typedef enum _TThreadStatus {
    WM_THREAD_USER_CANCEL = 1,
    WM_THREAD_ERROR_QUERY_DONT_SELECT,
    WM_THREAD_ERROR_BD_INCORRERCT,
    WM_THREAD_ERROR_BD_NOT_EXIST,
    WM_THREAD_ERROR_BD_CANT_CONNECT,
    WM_THREAD_ERROR_PARAMS_INCORRECT,
    WM_THREAD_ERROR_NULL_RESULTS,
    WM_THREAD_ERROR_OPEN_QUERY1,
    WM_THREAD_ERROR_OPEN_QUERY2,
    WM_THREAD_ERROR_OPEN_QUERY3,
    WM_THREAD_ERROR_TOO_MORE_RESULTS,
    WM_THREAD_ERROR_IN_PROCESS,         // ����� - ������ ��� ��������� ������
    WM_THREAD_ERROR_IN_PROCESS_ALT,     // ����� - ������ ��� ��������� ������ 

    WM_THREAD_PROCEED_BEGIN_SQL,    // ����� - ������ ������ � ��
    WM_THREAD_PROCEED_BEGIN_FETCH,    // ����� - ������ ���������� ������ �� ��
    WM_THREAD_PROCEED_BEGIN_DOCUMENT, // ����� - ������ �������� ���������
    WM_THREAD_PROCEED_EXCEL,       // ����� - ������������ ������ � Excel, �������������� ������ LPARAM
    WM_THREAD_EXECUTE_DONE,      // ����� ������� ������� (��� ������� � �������� ���������)
    WM_THREAD_EXECUTE_ERROR,      // ����� ������� � ������� (��� ������� � �������� ���������)
    WM_THREAD_COMPLETED_SUCCESSFULLY      // ����� ������� �������
} TThreadStatus;

//class TLogger;  // ����������� ����������

class TQueryItem;

typedef struct {
    AnsiString dstfilename;
    EXPORTMODE exportmode;
    String queryName;   // ������������ ������
    String querytext1;   // ����� ��������� �������
    String querytext2;  // ����� ��������������� �������
    String querytext3;  // ����� ��������������� �������

    TQueryItem* queryitem;   // ��� �������� �������� ���������� ����������?
    void* ParentFormHandle;
    TOraSession* TemplateOraSession1;
    TOraSession* TemplateOraSession2;
    TOraSession* TemplateOraSession3;

} TThreadOptions;


//---------------------------------------------------------------------------
class TThreadSelect : public TThread
{
public:
    //__fastcall ThreadSelect(bool CreateSuspended, THREADOPTIONS* threadopt, void (*f)(const String&, int));
    __fastcall TThreadSelect(bool CreateSuspended, TThreadOptions* threadopt);
    __fastcall ~TThreadSelect();
    TOraSession* __fastcall CreateOraSession(TOraSession* TemplateOraSession);

    void __fastcall SyncThreadChangeStatus();
    /* ������������� - ��������� ������� ���������� ������� */
    //void __fastcall SyncDebug();
    //String debug_message;

private:
    HWND ParentFormHandle;
    TDocumentWriter documentWriter;

    void __fastcall DoExportToWordTemplate();   //
    void __fastcall DoExportToExcel(bool toTemplate = false); // ���������� ������ Excel
    void __fastcall DoExportToDbf();



    EXPORTMODE ExportMode;
    AnsiString DstFileName;             // ��� �����-����������
    TOraSession* _oraSession1;
    TOraSession* _oraSession2; // ������ ������ ��� ������� ������� (������������ � ������� � MS Word)
    TOraSession* _oraSession3; // ������ ������ ��� ������� ������� (������������ � ������� � MS Word)
    TOraQuery* _oraQuery1;
    TOraQuery* _oraQuery2;   // ������ ������ (������������ � ������� � MS Word)
    TOraQuery* _oraQuery3;   // ������ ������ (������������ � ������� � MS Word)

    //TDataSet* JoinDataset(TDataSet* dsLeft, TDataSet* dsRight, String LeftKey, String RigthKey);

    static unsigned int _threadIndex;   // ������� �������
    unsigned int _threadId;  // ������������� ������
    TThreadStatus _threadStatus;
    AnsiString _threadMessage;

    std::vector<String> _resultFiles;   // ������ ���� ������ - �����������
    AnsiString _queryText1;
    AnsiString _queryText2;
    AnsiString _queryText3;
    AnsiString _reportName;

    TExcelExportParams param_excel;
    TWordExportParams param_word;
    TDbaseExportParams param_dbase;    // 2017-09-08


    std::vector<TParamRecord*> UserParams;    // ���������� ��������� � �������


    void SetThreadOpt(TThreadOptions* threadopt);
    void __fastcall Execute();
    void __fastcall setStatus(_TThreadStatus status, const AnsiString& message = "");

    void (*f)(const String&, int);

};
//---------------------------------------------------------------------------

#endif
