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
    String tabname;     // ������������ ������
    String taborder;    // ���������� ����� �������
    String queryid;     // id �������
    String querytext1;   // ����� �������
    String querytext2;  // ����� ������� ������� (������������ � ������� � MS Word)
    String querytext3;  // 
    String queryname;   // ������������ �������
    String dbname1;      // ������ ���� ������
    String dbname2;     // ������ ���� ������ ��� ������� ������� (������������ � ������� � MS Word)
    String dbname3;      // ������ ���� ������
    String sortorder;   // ������� ����������
    //AnsiString spr_task_sql2excel_id;
    String fieldslist;  // ������ - �������� ������� (����������� � �������)

    EXPORTMODE DefaultExportType;   // ��� ������ ��� �������� "�� ���������"

    String exportparam_id;


    TExcelExportParams param_excel;
    TDbaseExportParams param_dbase;
    TWordExportParams param_word;

    //EXPORT_PARAMS_WORD param_word;
    EXPORT_PARAMS_EXECUTE param_execute;

    QueryVariables UserParams;    // ���������� ��������� � ������

    bool fExcelFile;      // ���� Excel � ������
    bool fWordFile;     // ���� Word
    bool fDbfFile;       // ���� Dbf � ����
};

//---------------------------------------------------------------------------
#endif
