#ifndef DatatypeH
#define DatatypeH

#include <vector>
#include "DocumentWriter.h"

// Структура для хранения параметров поля (столбца) DBASE
/*typedef struct {    // Для описания структуры dbf-файла
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
} DBASEFIELD;*/

// Структура для хранения параметров поля (столбца) MS Excel
/*typedef struct {    // Для описания формата ячеек в Excel
    AnsiString format;      // Формат ячейки в Excel
    AnsiString name;        // Имя поля
    //int title_rows;       // Высота заголовка в строках
    int width;              // Ширина столбца
    int bwraptext;          // Флаг перенос по словам
} EXCELFIELD;  */

// Структура для хранения параметров экспорта в MS Excel
typedef struct  {
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

#endif // DatatypeH
