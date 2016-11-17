/*
 ����� TQueryItem

 */

#ifndef QueryItemH
#define QueryItemH

#include <vector.h>
#include "Datatype.h"
#include "Parameter.h"

// ����� �������� ������
typedef enum _EXPORTMODE {
    EM_UNDEFINITE = 0,
    EM_PROCEDURE,   // ��������� ��� ���������
    EM_EXCEL_BLANK,     // ������� � ������ ���� MS Excel
    EM_EXCEL_TEMPLATE,  // ������� � ������ MS Excel
    EM_DBASE4_FILE,     // ������� � DBF
    EM_WORD_TEMPLATE    // ������� � ������ MS Word
} EXPORTMODE;

typedef std::vector<TParamRecord*> QueryVariables;
typedef std::vector<TParamRecord*>::iterator QueryVariablesIterator;

// ��������� ��� ���������� ������� � ���������� � ����
class TQueryItem
{
public:
    AnsiString tabname;     // ������������ ������
    AnsiString taborder;    // ���������� ����� �������
    AnsiString queryid;     // id �������
    AnsiString querytext;   // ����� �������
    AnsiString querytext2;  // ����� ������� ������� (������������ � ������� � MS Word)
    AnsiString queryname;   // ������������ �������
    AnsiString dbname;      // ������ ���� ������
    AnsiString dbname2;     // ������ ���� ������ ��� ������� ������� (������������ � ������� � MS Word)
    AnsiString sortorder;   // ������� ����������
    AnsiString spr_task_sql2excel_id;
    AnsiString fieldslist;  // ������ - �������� ������� (����������� � �������)

    EXPORTMODE DefaultExportType;   // ��� ������ ��� �������� "�� ���������"

    AnsiString exportparam_id;


    EXPORT_PARAMS_EXCEL param_excel;
    EXPORT_PARAMS_DBASE param_dbase;
    EXPORT_PARAMS_WORD param_word;
    EXPORT_PARAMS_EXECUTE param_execute;

    QueryVariables UserParams;    // ���������� ��������� � ������

    bool fExcelFile;      // ���� Excel � ������
    bool fWordFile;     // ���� Word
    bool fDbfFile;       // ���� Dbf � ����
};

//---------------------------------------------------------------------------
#endif
