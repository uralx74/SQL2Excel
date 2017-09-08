//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
#pragma package(smart_init)

#include "ThreadSelect.h"
#include "FMain.h"



using namespace vartools;

unsigned int TThreadSelect::_threadIndex = 0;

/**/
//__fastcall ThreadSelect::ThreadSelect(bool CreateSuspended, THREADOPTIONS* threadopt, void (*f)(const String&, int))
__fastcall TThreadSelect::TThreadSelect(bool CreateSuspended, THREADOPTIONS* threadopt)
    : TThread(CreateSuspended),
    _threadMessage("")
{
    FreeOnTerminate = true;
    Suspended = true;
    //WParamResultMessage = 0;
    //LParamResultMessage = 0;
    //AppPath = ExtractFilePath(Application->ExeName);

    SetThreadOpt(threadopt);
    _threadIndex++;

    randomize();
    _threadId = random(9999999999);

    //documentWriter = new TDocumentWriter();

}

/**/
__fastcall TThreadSelect::~TThreadSelect()
{
    if (ThreadOraSession != NULL)
    {
        ThreadOraSession->Disconnect();
        //ThreadOraSession->Close();
        delete ThreadOraSession;
    }

    if (ThreadOraSession2 != NULL && ThreadOraSession != ThreadOraSession2)
    {
        ThreadOraSession2->Disconnect();
        delete ThreadOraSession2;
    }

    ThreadOraSession = NULL;
    ThreadOraSession2 = NULL;

    _resultFiles.clear();

    //QueryParams.free();

    _threadIndex--;
}

/* ��������� ���������� ��� ���������� ������� � ���������� ������ */
void TThreadSelect::SetThreadOpt(THREADOPTIONS* threadopt)
{
    //m_th_opt = *threadopt;
    this->ParentFormHandle = threadopt->ParentFormHandle;   // Handle ������� �����
    this->_reportName = threadopt->queryName;
    this->_mainQueryText = threadopt->querytext;            // ����� �������
    this->_secondaryQueryText = threadopt->querytext2;      // ����� �������
    this->DstFileName = threadopt->dstfilename;         // ��� ��������������� �����
    this->ExportMode = threadopt->exportmode;           // ����� �������� _EXPORTMODE


    ThreadOraSession = CreateOraSession(threadopt->TemplateOraSession);
    ThreadOraSession2 = NULL;
    if (threadopt->TemplateOraSession2 != NULL)
    {
        if (threadopt->TemplateOraSession != threadopt->TemplateOraSession2)        // ���� ���������� � ������ ��
        {
            ThreadOraSession2 = CreateOraSession(threadopt->TemplateOraSession2);   // �� ������� ����� ����������
        }
        else                                                                        // ����� �������� ��������� �� ������ ����������
        {
            ThreadOraSession2 = ThreadOraSession;
        }
    }

    // !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    // !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    // !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    // !!!!!!!!!!!!!!!������� �������� ���������� �� �������!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    // !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    // !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!


    switch (threadopt->exportmode) {
    case EM_EXCEL_BLANK:
    {
        // ������������ ������
        this->param_excel.title_label = threadopt->queryitem->param_excel.title_label != ""? threadopt->queryitem->param_excel.title_label : threadopt->queryitem->queryname;  // ������������ ������
        this->param_excel.title_height = threadopt->queryitem->param_excel.title_height;      // ������ ��������� ������� MS Excel
        this->param_excel.Fields = threadopt->queryitem->param_excel.Fields;  // ������ ����� ��� �������� � MS Excel

        //int k = threadopt->queryitem->param_excel.Fields.size();

        this->UserParams = threadopt->queryitem->UserParams;
        break;
    }
    case EM_EXCEL_TEMPLATE:
    {
        this->param_excel = threadopt->queryitem->param_excel;        //
        /*this->param_excel.templateFilename = threadopt->queryitem->param_excel.templateFilename;        // ������������ ������ � ������
        this->param_excel.table_range_name = threadopt->queryitem->param_excel.table_range_name;   2017-09-08*/
        this->param_excel.fUnbounded = threadopt->queryitem->param_excel.fUnbounded;        // ������������ ������ � ������
        break;
    }
    case EM_DBASE4_FILE:
    {
        this->param_dbase = threadopt->queryitem->param_dbase;
        //this->param_dbase.Fields = threadopt->queryitem->param_dbase.Fields;  // ������ ����� ��� �������� � DBF
        //this->param_dbase.fDisableUnassignedFields = threadopt->queryitem->param_dbase.fDisableUnassignedFields;
        this->param_dbase.resultFilename = threadopt->dstfilename; // 2017-09-08
        break;
    }
    case EM_PROCEDURE:
    {
        break;
    }
    case EM_WORD_TEMPLATE:
    {
        this->param_word = threadopt->queryitem->param_word;
        /*this->param_word.pagePerDocument = threadopt->queryitem->param_word.pagePerDocument;              // ���������� ������� �� �������� MS Word
        this->param_word.template_name = threadopt->queryitem->param_word.template_name;        // ������������ ������ � ������
        this->param_word.filter_main_field= threadopt->queryitem->param_word.filter_main_field;
        this->param_word.filter_sec_field = threadopt->queryitem->param_word.filter_sec_field;            */
        //this->param_word.filter_infix_sec_field = threadopt->queryitem->param_word.filter_infix_sec_field;
        break;
    }
    }
}

//---------------------------------------------------------------------------
//
TOraSession* __fastcall TThreadSelect::CreateOraSession(TOraSession* TemplateOraSession)
{
    TOraSession* OraSession = new TOraSession(NULL);
    //ThreadOraSession->OnError = OraSession1Error;
    OraSession->LoginPrompt = false;
    OraSession->Password = TemplateOraSession->Password;
    OraSession->Username = TemplateOraSession->Username;
    OraSession->Server = TemplateOraSession->Server;
    OraSession->Options = TemplateOraSession->Options;
    OraSession->HomeName = TemplateOraSession->HomeName;
    OraSession->ConnectMode = cmNormal;
    OraSession->Pooling = false;
    OraSession->ThreadSafety = true;


    OraSession->DataTypeMap->Clear();
    // rule for numeric(4,0)
    //ThreadOraSession->DataTypeMap->AddDBTypeRule(oraNumber, 0,      4, 0,     0, ftInteger, true);
    // rule for numeric(10,0)
    //ThreadOraSession->DataTypeMap->AddDBTypeRule(oraNumber, 5, 10, 0,     0, ftInteger, true);
    // rule for numeric(15,0)
    //ThreadOraSession->DataTypeMap->AddDBTypeRule(oraNumber, 11, rlAny, 0,     0, ftLargeint, true);
    // rule for numeric(5,2)
    //ThreadOraSession->DataTypeMap->AddDBTypeRule(oraNumber, 0, 9, 1, rlAny, ftInteger, true);
    // rule for numeric(10,4)
    //ThreadOraSession->DataTypeMap->AddDBTypeRule(oraNumber, 10, rlAny, 1,     4, ftInteger, true);

    OraSession->Connect();

 /*   OraQuery->DataTypeMap->Clear();
    // rule for numeric(4,0)
    OraQuery->DataTypeMap->AddDBTypeRule(oraNumber, 0,      4, 0,     0, ftSmallint, true);
    // rule for numeric(10,0)
    OraQuery->DataTypeMap->AddDBTypeRule(oraNumber, 5,     10, 0,     0, ftInteger, true);
    // rule for numeric(15,0)
    OraQuery->DataTypeMap->AddDBTypeRule(oraNumber, 11, rlAny, 0,     0, ftLargeint, true);
    // rule for numeric(5,2)
    OraQuery->DataTypeMap->AddDBTypeRule(oraNumber, 0,      9, 1, rlAny, ftFloat, true);
    // rule for numeric(10,4)
    OraQuery->DataTypeMap->AddDBTypeRule(oraNumber, 10, rlAny, 1,     4, ftBCD, true);
    // rule for numeric(15,6)
    //OraQuery->DataTypeMap->AddDBTypeRule(oraNumber, 10, rlAny, 5, rlAny, ftFMTBcd, true);
    //OraQuery->DataTypeMap->AddDBTypeRule(oraVarchar2, ftString);
    //OraQuery->DataTypeMap->AddDBTypeRule(oraNVarchar2, ftWideString);

    */

    return OraSession;
}

//---------------------------------------------------------------------------
void __fastcall TThreadSelect::Execute()
{
    if (!ThreadOraSession->Connected)
    {
        setStatus(WM_THREAD_ERROR_BD_CANT_CONNECT, "��������� ������. ���� ������ ����������.");
        this->Terminate();
    }

    if (!this->Terminated)
    {
        setStatus(WM_THREAD_PROCEED_BEGIN_SQL, _reportName);
        Synchronize(SyncThreadChangeStatus);

        if (ExportMode == EM_PROCEDURE)
        {
            // ��������� ��� ���������
            try
            {
                OraQueryMain = new TOraQuery(NULL);
                OraQueryMain->FetchAll = true;
                OraQueryMain->Session = ThreadOraSession;
                OraQueryMain->SQL->Add(_mainQueryText);
                OraQueryMain->Execute();

                setStatus(WM_THREAD_EXECUTE_DONE);
            }
            catch (Exception &e)
            {
                setStatus(WM_THREAD_EXECUTE_ERROR, e.Message);
            }

            try
            {
                delete OraQueryMain;
            }
            catch (...)
            {
            }

            OraQueryMain = NULL;

            Synchronize(SyncThreadChangeStatus);
            return;

        }
    }
            /*����� ������� ���� try ��� ������� ������.
            �������� ��������, ����� ����������� ������ � ��������, �������������� � �������.
            ����� ���������� �������� ��������������� ����� �� ������.
            ������ ��������� "... �������� ������������ �������."*/

    if (!this->Terminated && _mainQueryText != "")    // ���� ����� ������ ������
    {
        // ������� �������� �������� ������
        try
        {
            OraQueryMain = OpenOraQuery(ThreadOraSession, _mainQueryText, false);
        }
        catch (Exception &e)
        {
            setStatus(WM_THREAD_ERROR_OPEN_QUERY, "������ ��� ������� ��������� ������.\n" + e.Message);
            this->Terminate();
        }

        /*if (!this->Terminated && OraQueryMain == NULL)
        {
            setStatus(WM_THREAD_ERROR_OPEN_QUERY);
            this->Terminate();
        }*/
    }

    if (!this->Terminated && _secondaryQueryText != "")     // ���� ����� ������ ������
    {    // ������� ��������� ��������������� ������
        try
        {
            OraQuerySecondary = OpenOraQuery(ThreadOraSession2, _secondaryQueryText, false);
        }
        catch (Exception &e)
        {
            setStatus(WM_THREAD_ERROR_OPEN_QUERY2, "������ ��� ������� ��������� ������.\n" + e.Message);
            this->Terminate();
        }

        /*if (OraQuerySecondary == NULL)
        {
            setStatus(WM_THREAD_ERROR_OPEN_QUERY2);
            this->Terminate();
        } */
    }

    if (!this->Terminated)
    {
        // ���������� ������ �� �������
        setStatus(WM_THREAD_PROCEED_BEGIN_FETCH);
        Synchronize(SyncThreadChangeStatus);
    }

    if (!this->Terminated)
    {
        int RecCount = 0;

        OraQueryMain->FetchAll = true;

	    RecCount = OraQueryMain->RecordCount;

        if (RecCount <= 0) // ���� ������ �� ������ �������
        {
            setStatus(WM_THREAD_ERROR_NULL_RESULTS);
            this->Terminate();
        }
        else if (RecCount > 200000 && RecCount < 1000000) // ���� ������ ������ ����� 200 000 �������
        {
            AnsiString msg = "� ������ �������� ���������� �������� " + IntToStr(RecCount) +" �����.\n������������ ������ ����� ������ ���������� �����.\n������������ �����?";
            if (MessageBoxQuestion(msg) == IDNO) {
                setStatus(WM_THREAD_USER_CANCEL);
                this->Terminate();
            }
        }
        else if (RecCount >= 1000000)
        {
            setStatus(WM_THREAD_ERROR_TOO_MORE_RESULTS);
            this->Terminate();
        }
    }

    // �������� ���������
    if ( !this->Terminated )
    {
        try
        {
            setStatus(WM_THREAD_PROCEED_BEGIN_DOCUMENT);
            Synchronize(SyncThreadChangeStatus);

            switch (ExportMode) {
            case EM_EXCEL_TEMPLATE:
                CoInitialize(NULL);
                DoExportToExcel();
                //ExportToExcelTemplate(OraQueryMain, OraQuerySecondary);
                CoUninitialize();
                break;
            case EM_EXCEL_BLANK:
                CoInitialize(NULL);
                DoExportToExcel();
                CoUninitialize();
                break;
            //case EM_EXCEL_FILE_TEMPLATE:
            //    break;
            //case EM_EXCEL_FILE:
            //case EM_EXCEL_MEMORY:
                //ExportToExcel(OraQueryMain);
                //break;
            case EM_DBASE4_FILE:
                DoExportToDbf();
                break;
            case EM_WORD_TEMPLATE:
                CoInitialize(NULL);
                DoExportToWordTemplate();
                CoUninitialize();
                break;
            }
        }
        catch ( Exception& e )
        {
            CoUninitialize();
            setStatus(WM_THREAD_ERROR_IN_PROCESS, e.Message);
            this->Terminate();
        }
        if ( _threadStatus != WM_THREAD_PROCEED_BEGIN_DOCUMENT )
        {
            this->Terminate();
        }
    }

    // �������� ��������� �������
    if ( !this->Terminated )
    {
        if ( OraQueryMain != NULL )
        {
            try
            {
                OraQueryMain->Close();
                delete OraQueryMain;
                OraQueryMain = NULL;
            }
            catch (...)
            {
            }
        }

        // �������� ���������������� �������
        if ( OraQuerySecondary != NULL )
        {
            try
            {
                OraQuerySecondary->Close();
                delete OraQuerySecondary;
                OraQuerySecondary = NULL;
            }
            catch (...)
            {
            }
        }
        setStatus(WM_THREAD_COMPLETED_SUCCESSFULLY);
    }

    Synchronize(SyncThreadChangeStatus);
}

/* ��������� �������� � ������ ���� Excel
*/
void __fastcall TThreadSelect::DoExportToExcel()
{

    // ��������� ������, �� ��������� ����������, ��������� �������������
    // �������� � ������� ������� ������������� ��������� � ����� "separator",
    Variant data_parameters;
    int param_count = UserParams.size();  // ��������� ������
    int visible_param_count = 0;

    for (int i=0; i <= param_count-1; i++)    // ������������ ���-�� ������������ ����������
    {
        if ( UserParams[i]->isVisible() )
        {
            visible_param_count++;
        }
    }

    if (param_count > 0)     // ������ ���������� ��� ������ � Excel
    {
        data_parameters = CreateVariantArray(visible_param_count, 1);
        //data_parameters = CreateVariantArray(2, 1);


        for (int i=0; i <= param_count-1; i++)
        {
            if ( !UserParams[i]->isVisible() )
            {
                continue;
            }

            if (UserParams[i]->type != "separator")
            {
                data_parameters.PutElement(UserParams[i]->getCaption() + ": " + UserParams[i]->getDisplay(), i+1, 1);
            }
            else
            {
                data_parameters.PutElement("[" + UserParams[i]->getCaption() + "]", i+1, 1);
            }
        }
    }


    Variant sql_text;
    {
        // ������ ��� sql-������
        Variant range_sqltext;
        int PartMaxLength = 4000;  // 8 192  - ������������ ����� ������ � ������ EXCEL
        int n = ceil( (float) _mainQueryText.Length() / PartMaxLength);

        sql_text = CreateVariantArray(n, 1);

        for (int i = 1; i <= n; i++)
        {
            String sQueryPart = _mainQueryText.SubString(((i-1) * PartMaxLength) + 1, PartMaxLength);
            sql_text.PutElement(sQueryPart, i, 1);
        }
    }

    // ��������� ��������� ������
    param_excel.addTableDataSet(OraQueryMain, "table_body", "table_column_");
    param_excel.addTableVtArray(data_parameters, "report_parameters");
    param_excel.addTableVtArray(sql_text, "report_query_text");

    documentWriter.ExportToExcel(&param_excel);
    _resultFiles = documentWriter._result.resultFiles;

}


/* ��������� �������� � ������ Word
   ����������: ������� ��������� � ��������� ������ */
void __fastcall TThreadSelect::DoExportToWordTemplate()
{
    if ( OraQueryMain->RecordCount == 0)
    {
        return;
    }

    TWordExportParams* wordExportParams = &param_word;

    //TWordExportParams wordExportParams;
    //wordExportParams.pagePerDocument = param_word.pagePerDocument;

    /* ������������ ��������� ������ */
    wordExportParams->addFormtextDataSet(OraQuerySecondary);     // ����� ���������� �� �������
    //wordExportParams.addSingleTextDataSet(OraQuerySecondary, "rec_");  // ���������� �� �������
    wordExportParams->addMergeDataSet(OraQueryMain);
    //wordExportParams->templateFilename =  param_word.templateFilename;


    bool bFilterExist = param_word.filter_main_field != "" && param_word.filter_sec_field != "";    // ���� � ���������� ����� ������, �� �������, ��� ���������� ������

    if ( bFilterExist ) // ���� ����� ������
    {
        int i = 1;
        while( !OraQuerySecondary->Eof )
        {
            // ���� ����� ������, ��������� ��� � ��������� �������
            try
            {
                String sFilter = param_word.filter_main_field + "='" + OraQuerySecondary->FieldByName(param_word.filter_sec_field)->AsString + "'";
                OraQueryMain->Filtered = false;
                OraQueryMain->Filter = sFilter;
                OraQueryMain->Filtered = true;
            }
            catch ( Exception &e )
            {
                OraQueryMain->Filtered = false;
                _threadStatus = WM_THREAD_ERROR_IN_PROCESS;
                _threadMessage = "��������� ������������ ���������� ������� � ���������� �������� ��� ���������� � ���������� ��������������.\n" + e.Message;
                break;
            }

            if ( OraQueryMain->RecordCount > 0 )         // ���� ���� ������, �� ��������� ��������
            {
                wordExportParams->resultFilename = ExtractFilePath(DstFileName) + ExtractFileName(DstFileName) + "_" + IntToStr(i++)+ "_[:counter].doc";
                documentWriter.ExportToWordTemplate(wordExportParams);
                _resultFiles = documentWriter._result.resultFiles;
            }
            OraQuerySecondary->Next();
        }
    }
    else
    {   // ���� ������ �� �����, ������ ������ ��� ����
        wordExportParams->resultFilename = ExtractFilePath(DstFileName) + ExtractFileName(DstFileName) + "_[:counter].doc";
        documentWriter.ExportToWordTemplate(wordExportParams);
        _resultFiles = documentWriter._result.resultFiles;
    }
}

/* ��������� �������� � dbf-����
   ����������: ������� ��������� � ��������� ������ */
void __fastcall TThreadSelect::DoExportToDbf()
{
    if ( OraQueryMain->RecordCount == 0)
    {
        return;
    }
    param_dbase.srcDataSet = OraQueryMain;  // �������� ��������� � SetThreadOpt 2017-09-08

    documentWriter.ExportToDbf(&param_dbase);
    _resultFiles = documentWriter._result.resultFiles;

}


TThreadSelectMessage::TThreadSelectMessage(unsigned int status, const AnsiString& message, std::vector<String> files) :
    _status(status),
    _message(message),
    _files(files)
{
}

TThreadSelectMessage::TThreadSelectMessage(unsigned int status, const AnsiString& message) :
    _status(status),
    _message(message)
{
}

void __fastcall TThreadSelect::setStatus(_TThreadStatus status, const AnsiString& message)
{
    _threadStatus = status;
    _threadMessage = message;

}

/* ������������� - ��� ������� */
/*void __fastcall TThreadSelect::SyncDebug()
{
    ShowMessage(debug_message);
}*/

/* ������������� - ��������� ������� ���������� ������� */
void __fastcall TThreadSelect::SyncThreadChangeStatus()
{
    TThreadSelectMessage message(_threadStatus, _threadMessage, _resultFiles);
    Form1->threadListener(_threadId, message);
}


