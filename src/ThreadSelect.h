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
#include "..\util\MSExcelWorks.h"
#include "..\util\MSWordWorks.h"
#include "taskutils.h"
#include "DocumentWriter.h"


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
    WM_THREAD_ERROR_OPEN_QUERY,
    WM_THREAD_ERROR_OPEN_QUERY2,
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

/*
// ����� �������� ������
typedef enum _EXPORTMODE {
    EM_UNDEFINITE = 0,
    EM_PROCEDURE,   // ��������� ��� ���������
    EM_EXCEL_BLANK,     // ������� � ������ ���� MS Excel
    EM_EXCEL_TEMPLATE,  // ������� � ������ MS Excel
    EM_DBASE4_FILE,     // ������� � DBF
    EM_WORD_TEMPLATE    // ������� � ������ MS Word
} EXPORTMODE;
*/

class TQueryItem;

typedef struct {
    AnsiString dstfilename;
    EXPORTMODE exportmode;
    String queryName;   // ������������ ������
    String querytext;   // ����� ��������� �������
    String querytext2;  // ����� ��������������� �������

    TQueryItem* queryitem;   // ��� �������� �������� ���������� ����������?
    void* ParentFormHandle;
    TOraSession* TemplateOraSession;
    TOraSession* TemplateOraSession2;

} THREADOPTIONS;


//---------------------------------------------------------------------------
class TThreadSelect : public TThread
{
public:
    //__fastcall ThreadSelect(bool CreateSuspended, THREADOPTIONS* threadopt, void (*f)(const String&, int));
    __fastcall TThreadSelect(bool CreateSuspended, THREADOPTIONS* threadopt);
    __fastcall ~TThreadSelect();
    TOraSession* __fastcall CreateOraSession(TOraSession* TemplateOraSession);

    void __fastcall SyncThreadChangeStatus();
    /* ������������� - ��������� ������� ���������� ������� */
    //void __fastcall SyncDebug();
    //String debug_message;

private:
    HWND ParentFormHandle;
    //AnsiString AppPath;     // ���� � ����������
    TDocumentWriter documentWriter;

    void __fastcall DoExportToWordTemplate();   //
    void __fastcall DoExportToExcel(); // ���������� ������ Excel



    EXPORTMODE ExportMode;
    AnsiString DstFileName;             // ��� �����-����������
    TOraSession* ThreadOraSession;
    TOraSession* ThreadOraSession2; // ������ ������ ��� ������� ������� (������������ � ������� � MS Word)
    TOraQuery* OraQueryMain;
    TOraQuery* OraQuerySecondary;   // ������ ������ (������������ � ������� � MS Word)
    static unsigned int _threadIndex;   // ������� �������
    unsigned int _threadId;  // ������������� ������
    TThreadStatus _threadStatus;
    AnsiString _threadMessage;

    std::vector<String> _resultFiles;   // ������ ���� ������ - �����������
    AnsiString _mainQueryText;
    AnsiString _secondaryQueryText;
    AnsiString _reportName;

    TExcelExportParams param_excel;
    //EXPORT_PARAMS_EXCEL param_excel;

    EXPORT_PARAMS_WORD param_word;      //
    EXPORT_PARAMS_DBASE param_dbase;

    std::vector<TParamRecord*> UserParams;    // ���������� ��������� � �������


    void SetThreadOpt(THREADOPTIONS* threadopt);
    void __fastcall Execute();
    //void __fastcall ExportToExcel(TOraQuery *OraQuery); // ���������� ������ Excel
    void __fastcall ExportToExcelTemplate(TOraQuery *QueryTable, TOraQuery *QueryFields);
    void __fastcall ExportToDBF(TOraQuery *OraQuery);   // ���������� DBF-�����
    //void __fastcall ExportToWordTemplate(TOraQuery *QueryMerge, TOraQuery *QueryFormFields);  // ���������� ������ Word �� ���� �������
    void __fastcall setStatus(_TThreadStatus status, const AnsiString& message = "");

    void (*f)(const String&, int);

    //Variant ExportToExcelTable1(TOraQuery* QTable, Variant Worksheet, bool bUnbounded = true);   // �������� !!!!!!!!!!!

};
//---------------------------------------------------------------------------

#endif
