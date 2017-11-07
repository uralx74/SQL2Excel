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
    String tabname;     // Наименование раздел
    String taborder;    // Порядковый номер вкладки
    String queryid;     // id запроса
    String querytext1;   // Текст запроса
    String querytext2;  // Текст второго запроса (используется в отчетах в MS Word)
    String querytext3;  // 
    String queryname;   // Наименование запроса
    String dbname1;      // Индекс базы данных
    String dbname2;     // Индекс базы данных для второго запроса (используется в отчетах в MS Word)
    String dbname3;      // Индекс базы данных
    String sortorder;   // Порядок сортировки
    //AnsiString spr_task_sql2excel_id;
    String fieldslist;  // Строка - перечень записей (комментарий к запросу)

    EXPORTMODE DefaultExportType;   // Тип отчета для выгрузки "По умолчанию"

    String exportparam_id;


    TExcelExportParams param_excel;
    TDbaseExportParams param_dbase;
    TWordExportParams param_word;

    //EXPORT_PARAMS_WORD param_word;
    EXPORT_PARAMS_EXECUTE param_execute;

    QueryVariables UserParams;    // Задаваемые параметры к запрос

    bool fExcelFile;      // Флаг Excel в память
    bool fWordFile;     // Флаг Word
    bool fDbfFile;       // Флаг Dbf в файл
};

//---------------------------------------------------------------------------
#endif
