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
    std::vector<String> _files;   // Список имен файлов - результатов
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
    WM_THREAD_ERROR_IN_PROCESS,         // Поток - ошибка при обработке данных
    WM_THREAD_ERROR_IN_PROCESS_ALT,     // Поток - ошибка при обработке данных 

    WM_THREAD_PROCEED_BEGIN_SQL,    // Поток - начата работа с БД
    WM_THREAD_PROCEED_BEGIN_FETCH,    // Поток - начато извлечение данных из БД
    WM_THREAD_PROCEED_BEGIN_DOCUMENT, // Поток - начато создание документа
    WM_THREAD_PROCEED_EXCEL,       // Поток - продолжается работа с Excel, обрабатывается запись LPARAM
    WM_THREAD_EXECUTE_DONE,      // Поток окончен успешно (при запуске в качестве Процедуры)
    WM_THREAD_EXECUTE_ERROR,      // Поток окончен с ошибкой (при запуске в качестве Процедуры)
    WM_THREAD_COMPLETED_SUCCESSFULLY      // Поток окончен успешно
} TThreadStatus;

//class TLogger;  // опережающее объявление

/*
// Режим экспорта данных
typedef enum _EXPORTMODE {
    EM_UNDEFINITE = 0,
    EM_PROCEDURE,   // Выполнить как процедуру
    EM_EXCEL_BLANK,     // Экспорт в пустой файл MS Excel
    EM_EXCEL_TEMPLATE,  // Экспорт в шаблон MS Excel
    EM_DBASE4_FILE,     // Экспорт в DBF
    EM_WORD_TEMPLATE    // Экспорт в шаблон MS Word
} EXPORTMODE;
*/

class TQueryItem;

typedef struct {
    AnsiString dstfilename;
    EXPORTMODE exportmode;
    String queryName;   // Наименование запроа
    String querytext;   // Текст основного запроса
    String querytext2;  // Текст дополнительного запроса

    TQueryItem* queryitem;   // ЭТО ВОЗМОЖНО ЗАМЕНИТЬ ОТДЕЛЬНЫМИ ЗНАЧЕНИЯМИ?
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
    /* Синхронизация - изменение статуса выполнения запроса */
    //void __fastcall SyncDebug();
    //String debug_message;

private:
    HWND ParentFormHandle;
    //AnsiString AppPath;     // Путь к приложению
    TDocumentWriter documentWriter;

    void __fastcall DoExportToWordTemplate();   //
    void __fastcall DoExportToExcel(); // Заполнение отчета Excel



    EXPORTMODE ExportMode;
    AnsiString DstFileName;             // Имя файла-назначения
    TOraSession* ThreadOraSession;
    TOraSession* ThreadOraSession2; // Вторая сессия для второго запроса (используется в отчетах в MS Word)
    TOraQuery* OraQueryMain;
    TOraQuery* OraQuerySecondary;   // Второй запрос (используется в отчетах в MS Word)
    static unsigned int _threadIndex;   // Счетчик потоков
    unsigned int _threadId;  // Идентификатор потока
    TThreadStatus _threadStatus;
    AnsiString _threadMessage;

    std::vector<String> _resultFiles;   // Список имен файлов - результатов
    AnsiString _mainQueryText;
    AnsiString _secondaryQueryText;
    AnsiString _reportName;

    TExcelExportParams param_excel;
    //EXPORT_PARAMS_EXCEL param_excel;

    EXPORT_PARAMS_WORD param_word;      //
    EXPORT_PARAMS_DBASE param_dbase;

    std::vector<TParamRecord*> UserParams;    // Задаваемые параметры к запросу


    void SetThreadOpt(THREADOPTIONS* threadopt);
    void __fastcall Execute();
    //void __fastcall ExportToExcel(TOraQuery *OraQuery); // Заполнение отчета Excel
    void __fastcall ExportToExcelTemplate(TOraQuery *QueryTable, TOraQuery *QueryFields);
    void __fastcall ExportToDBF(TOraQuery *OraQuery);   // Заполнение DBF-файла
    //void __fastcall ExportToWordTemplate(TOraQuery *QueryMerge, TOraQuery *QueryFormFields);  // Заполнение отчета Word на базе шаблона
    void __fastcall setStatus(_TThreadStatus status, const AnsiString& message = "");

    void (*f)(const String&, int);

    //Variant ExportToExcelTable1(TOraQuery* QTable, Variant Worksheet, bool bUnbounded = true);   // Временно !!!!!!!!!!!

};
//---------------------------------------------------------------------------

#endif
