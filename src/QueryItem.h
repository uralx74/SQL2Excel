/*
 Класс TQueryItem

 */

#ifndef QueryItemH
#define QueryItemH

#include <vector.h>
#include "Datatype.h"
#include "Parameter.h"

// Режим экспорта данных
typedef enum _EXPORTMODE {
    EM_UNDEFINITE = 0,
    EM_PROCEDURE,   // Выполнить как процедуру
    EM_EXCEL_BLANK,     // Экспорт в пустой файл MS Excel
    EM_EXCEL_TEMPLATE,  // Экспорт в шаблон MS Excel
    EM_DBASE4_FILE,     // Экспорт в DBF
    EM_WORD_TEMPLATE    // Экспорт в шаблон MS Word
} EXPORTMODE;

typedef std::vector<TParamRecord*> QueryVariables;
typedef std::vector<TParamRecord*>::iterator QueryVariablesIterator;

// Структура для сохранения Запроса и параметров к нему
class TQueryItem
{
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

    AnsiString exportparam_id;


    EXPORT_PARAMS_EXCEL param_excel;
    EXPORT_PARAMS_DBASE param_dbase;
    EXPORT_PARAMS_WORD param_word;
    EXPORT_PARAMS_EXECUTE param_execute;

    QueryVariables UserParams;    // Задаваемые параметры к запрос

    bool fExcelFile;      // Флаг Excel в память
    bool fWordFile;     // Флаг Word
    bool fDbfFile;       // Флаг Dbf в файл
};

//---------------------------------------------------------------------------
#endif
