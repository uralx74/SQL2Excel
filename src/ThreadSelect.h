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
    WM_THREAD_ERROR_OPEN_QUERY1,
    WM_THREAD_ERROR_OPEN_QUERY2,
    WM_THREAD_ERROR_OPEN_QUERY3,
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

class TQueryItem;

typedef struct {
    AnsiString dstfilename;
    EXPORTMODE exportmode;
    String queryName;   // Наименование запроа
    String querytext1;   // Текст основного запроса
    String querytext2;  // Текст дополнительного запроса
    String querytext3;  // Текст дополнительного запроса

    TQueryItem* queryitem;   // ЭТО ВОЗМОЖНО ЗАМЕНИТЬ ОТДЕЛЬНЫМИ ЗНАЧЕНИЯМИ?
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
    /* Синхронизация - изменение статуса выполнения запроса */
    //void __fastcall SyncDebug();
    //String debug_message;

private:
    HWND ParentFormHandle;
    TDocumentWriter documentWriter;

    void __fastcall DoExportToWordTemplate();   //
    void __fastcall DoExportToExcel(bool toTemplate = false); // Заполнение отчета Excel
    void __fastcall DoExportToDbf();



    EXPORTMODE ExportMode;
    AnsiString DstFileName;             // Имя файла-назначения
    TOraSession* _oraSession1;
    TOraSession* _oraSession2; // Вторая сессия для второго запроса (используется в отчетах в MS Word)
    TOraSession* _oraSession3; // Вторая сессия для второго запроса (используется в отчетах в MS Word)
    TOraQuery* _oraQuery1;
    TOraQuery* _oraQuery2;   // Второй запрос (используется в отчетах в MS Word)
    TOraQuery* _oraQuery3;   // Второй запрос (используется в отчетах в MS Word)

    //TDataSet* JoinDataset(TDataSet* dsLeft, TDataSet* dsRight, String LeftKey, String RigthKey);

    static unsigned int _threadIndex;   // Счетчик потоков
    unsigned int _threadId;  // Идентификатор потока
    TThreadStatus _threadStatus;
    AnsiString _threadMessage;

    std::vector<String> _resultFiles;   // Список имен файлов - результатов
    AnsiString _queryText1;
    AnsiString _queryText2;
    AnsiString _queryText3;
    AnsiString _reportName;

    TExcelExportParams param_excel;
    TWordExportParams param_word;
    TDbaseExportParams param_dbase;    // 2017-09-08


    std::vector<TParamRecord*> UserParams;    // Задаваемые параметры к запросу


    void SetThreadOpt(TThreadOptions* threadopt);
    void __fastcall Execute();
    void __fastcall setStatus(_TThreadStatus status, const AnsiString& message = "");

    void (*f)(const String&, int);

};
//---------------------------------------------------------------------------

#endif
