//---------------------------------------------------------------------------

#ifndef ThreadSelectH
#define ThreadSelectH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include "Ora.hpp"
#include "OraDataTypeMap.hpp"

#include "Halcn6DB.hpp"

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
    WM_THREAD_PROCEED_DONE,      // ����� ������� �������
    WM_THREAD_EXECUTE_DONE,      // ����� ������� ������� (��� ������� � �������� ���������)
    WM_THREAD_EXECUTE_ERROR      // ����� ������� � ������� (��� ������� � �������� ���������)
} TThreadStatus;

//class TLogger;  // ����������� ����������

typedef struct {
    AnsiString dstfilename;
    EXPORTMODE exportmode;
    String querytext;   // ����� ��������� �������
    String querytext2;  // ����� ��������������� �������

    TQueryItem* queryitem;   // ��� �������� �������� ���������� ����������?
    void* ParentFormHandle;
    TOraSession* TemplateOraSession;
    TOraSession* TemplateOraSession2;

} THREADOPTIONS;

//---------------------------------------------------------------------------
class ThreadSelect : public TThread
{
public:
    //enum EXPORTMODE {TO_EXCEL= 1, TO_DBASE4};
    __fastcall ThreadSelect(bool CreateSuspended);
    __fastcall ~ThreadSelect();
    TOraSession* __fastcall CreateOraSession(TOraSession* TemplateOraSession);
    void SetThreadOpt(THREADOPTIONS* threadopt);
    void __fastcall SyncThreadDone();
    void __fastcall SyncThreadChangeStatus();

private:
    HWND ParentFormHandle;
    AnsiString AppPath;     // ���� � ����������

    EXPORTMODE ExportMode;
    AnsiString DstFileName;             // ��� �����-����������
    TOraSession* ThreadOraSession;
    TOraSession* ThreadOraSession2; // ������ ������ ��� ������� ������� (������������ � ������� � MS Word)
    TOraQuery* OraQueryMain;
    TOraQuery* OraQuerySecondary;   // ������ ������ (������������ � ������� � MS Word)


    AnsiString sQueryText;
    AnsiString sQueryText2;
    AnsiString sReportName;

    EXPORT_PARAMS_EXCEL param_excel;
    EXPORT_PARAMS_WORD param_word;      //
    EXPORT_PARAMS_DBASE param_dbase;

    std::vector<TParamRecord*> UserParams;    // ���������� ��������� � �������

    TThreadStatus ThreadStatus;
    AnsiString ThreadMessage;

    std::vector<String> vResultFiles;   // ������ ���� ������ - �����������

protected:
    void __fastcall Execute();
    void __fastcall ExportToExcel(TOraQuery *OraQuery); // ���������� ������ Excel
    void __fastcall ExportToExcelTemplate(TOraQuery *QueryTable, TOraQuery *QueryFields);
    void __fastcall ExportToDBF(TOraQuery *OraQuery);   // ���������� DBF-�����
    void __fastcall ExportToWordTemplate(TOraQuery *QueryMerge, TOraQuery *QueryFormFields);  // ���������� ������ Word �� ���� �������



    //Variant ExportToExcelTable1(TOraQuery* QTable, Variant Worksheet, bool bUnbounded = true);   // �������� !!!!!!!!!!!

};
//---------------------------------------------------------------------------

#endif
