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
    WM_THREAD_ERROR_IN_PROCESS,         // Поток - ошибка при обработке данных
    WM_THREAD_ERROR_IN_PROCESS_ALT,     // Поток - ошибка при обработке данных 

    WM_THREAD_PROCEED_BEGIN_SQL,    // Поток - начата работа с БД
    WM_THREAD_PROCEED_BEGIN_FETCH,    // Поток - начато извлечение данных из БД
    WM_THREAD_PROCEED_BEGIN_DOCUMENT, // Поток - начато создание документа
    WM_THREAD_PROCEED_EXCEL,       // Поток - продолжается работа с Excel, обрабатывается запись LPARAM
    WM_THREAD_PROCEED_DONE,      // Поток окончен успешно
    WM_THREAD_EXECUTE_DONE,      // Поток окончен успешно (при запуске в качестве Процедуры)
    WM_THREAD_EXECUTE_ERROR      // Поток окончен с ошибкой (при запуске в качестве Процедуры)
} TThreadStatus;

//class TLogger;  // опережающее объявление

typedef struct {
    AnsiString dstfilename;
    EXPORTMODE exportmode;
    String querytext;   // Текст основного запроса
    String querytext2;  // Текст дополнительного запроса

    TQueryItem* queryitem;   // ЭТО ВОЗМОЖНО ЗАМЕНИТЬ ОТДЕЛЬНЫМИ ЗНАЧЕНИЯМИ?
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
    AnsiString AppPath;     // Путь к приложению

    EXPORTMODE ExportMode;
    AnsiString DstFileName;             // Имя файла-назначения
    TOraSession* ThreadOraSession;
    TOraSession* ThreadOraSession2; // Вторая сессия для второго запроса (используется в отчетах в MS Word)
    TOraQuery* OraQueryMain;
    TOraQuery* OraQuerySecondary;   // Второй запрос (используется в отчетах в MS Word)


    AnsiString sQueryText;
    AnsiString sQueryText2;
    AnsiString sReportName;

    EXPORT_PARAMS_EXCEL param_excel;
    EXPORT_PARAMS_WORD param_word;      //
    EXPORT_PARAMS_DBASE param_dbase;

    std::vector<TParamRecord*> UserParams;    // Задаваемые параметры к запросу

    TThreadStatus ThreadStatus;
    AnsiString ThreadMessage;

    std::vector<String> vResultFiles;   // Список имен файлов - результатов

protected:
    void __fastcall Execute();
    void __fastcall ExportToExcel(TOraQuery *OraQuery); // Заполнение отчета Excel
    void __fastcall ExportToExcelTemplate(TOraQuery *QueryTable, TOraQuery *QueryFields);
    void __fastcall ExportToDBF(TOraQuery *OraQuery);   // Заполнение DBF-файла
    void __fastcall ExportToWordTemplate(TOraQuery *QueryMerge, TOraQuery *QueryFormFields);  // Заполнение отчета Word на базе шаблона



    //Variant ExportToExcelTable1(TOraQuery* QTable, Variant Worksheet, bool bUnbounded = true);   // Временно !!!!!!!!!!!

};
//---------------------------------------------------------------------------

#endif
