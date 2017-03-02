#ifndef DatatypeH
#define DatatypeH

#include <vector>
#include "DocumentWriter.h"

// ��������� ��� �������� ���������� ���� (�������) DBASE
/*typedef struct {    // ��� �������� ��������� dbf-�����
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
} DBASEFIELD;*/

// ��������� ��� �������� ���������� ���� (�������) MS Excel
/*typedef struct {    // ��� �������� ������� ����� � Excel
    AnsiString format;      // ������ ������ � Excel
    AnsiString name;        // ��� ����
    //int title_rows;       // ������ ��������� � �������
    int width;              // ������ �������
    int bwraptext;          // ���� ������� �� ������
} EXCELFIELD;  */

// ��������� ��� �������� ���������� �������� � MS Excel
typedef struct  {
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

#endif // DatatypeH
