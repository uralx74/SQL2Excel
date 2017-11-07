//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
#pragma package(smart_init)

#include "ThreadSelect.h"
#include "FMain.h"



using namespace vartools;
using namespace datasettools;

unsigned int TThreadSelect::_threadIndex = 0;

/**/
__fastcall TThreadSelect::TThreadSelect(bool CreateSuspended, TThreadOptions* threadopt)
    : TThread(CreateSuspended),
    _threadMessage("")
{
    FreeOnTerminate = true;
    Suspended = true;

    SetThreadOpt(threadopt);
    _threadIndex++;

    randomize();
    _threadId = random(9999999999);

}

/**/
__fastcall TThreadSelect::~TThreadSelect()
{
    if (_oraSession1 != NULL)
    {
        _oraSession1->Disconnect();
        //ThreadOraSession->Close();
        delete _oraSession1;
    }

    if (_oraSession2 != NULL && _oraSession1 != _oraSession2)
    {
        _oraSession2->Disconnect();
        delete _oraSession2;
    }

    _oraSession1 = NULL;
    _oraSession2 = NULL;
    _oraSession3 = NULL;

    _resultFiles.clear();

    _threadIndex--;
}

/* ��������� ���������� ��� ���������� ������� � ���������� ������ */
void TThreadSelect::SetThreadOpt(TThreadOptions* threadopt)
{
    //m_th_opt = *threadopt;
    this->ParentFormHandle = threadopt->ParentFormHandle;   // Handle ������� �����
    this->_reportName = threadopt->queryName;

    this->_queryText1 = threadopt->querytext1;            // ����� �������
    this->_queryText2 = threadopt->querytext2;      // ����� �������
    this->_queryText3 = threadopt->querytext3;      // ����� �������


    this->DstFileName = threadopt->dstfilename;         // ��� ��������������� �����
    this->ExportMode = threadopt->exportmode;           // ����� �������� _EXPORTMODE


    _oraSession1 = CreateOraSession(threadopt->TemplateOraSession1);
    _oraSession2 = NULL;
    _oraSession3 = NULL;

    if (threadopt->TemplateOraSession2 != NULL)
    {
        if (threadopt->TemplateOraSession1 != threadopt->TemplateOraSession2)        // ���� ���������� � ������ ��
        {
            _oraSession2 = CreateOraSession(threadopt->TemplateOraSession2);        // �� ������� ����� ����������
        }
        else                                                                        // ����� �������� ��������� �� ������ ����������
        {
            _oraSession2 = _oraSession1;
        }
    }

    if (threadopt->TemplateOraSession3 != NULL)
    {
        if (threadopt->TemplateOraSession1 != threadopt->TemplateOraSession3)        // ���� ���������� � ������ ��
        {
            _oraSession3 = CreateOraSession(threadopt->TemplateOraSession3);        // �� ������� ����� ����������
        }
        else                                                                        // ����� �������� ��������� �� ������ ����������
        {
            _oraSession3 = _oraSession1;
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
        this->param_excel = threadopt->queryitem->param_excel;
        // ���� ������������ ��������� �� ������ � ����������, �� ����� ��� �� ������������ �������
        this->param_excel.title_label = threadopt->queryitem->param_excel.title_label != ""? threadopt->queryitem->param_excel.title_label : threadopt->queryitem->queryname;  // ������������ ������
        this->UserParams = threadopt->queryitem->UserParams;    // ���������, �������� �������������
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
        this->param_dbase.resultFilename = threadopt->dstfilename; // ��� �����-����������
        break;
    }
    case EM_PROCEDURE:
    {
        break;
    }
    case EM_WORD_TEMPLATE:
    {
        this->param_word = threadopt->queryitem->param_word;
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
    if ( !_oraSession1->Connected )
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
                _oraQuery1 = new TOraQuery(NULL);
                _oraQuery1->FetchAll = true;
                _oraQuery1->Session = _oraSession1;
                _oraQuery1->SQL->Add(_queryText1);
                _oraQuery1->Execute();

                setStatus(WM_THREAD_EXECUTE_DONE);
            }
            catch (Exception &e)
            {
                setStatus(WM_THREAD_EXECUTE_ERROR, e.Message);
            }

            try
            {
                delete _oraQuery1;
            }
            catch (...)
            {
            }

            _oraQuery1 = NULL;

            Synchronize(SyncThreadChangeStatus);
            return;

        }
    }
            /*����� ������� ���� try ��� ������� ������.
            �������� ��������, ����� ����������� ������ � ��������, �������������� � �������.
            ����� ���������� �������� ��������������� ����� �� ������.
            ������ ��������� "... �������� ������������ �������."*/

    // ��������� ������ ������
    if (!this->Terminated && _queryText1 != "")    // ���� ����� ������ ������
    {
        // ������� �������� �������� ������
        try
        {
            _oraQuery1 = OpenOraQuery(_oraSession1, _queryText1, false);
        }
        catch (Exception &e)
        {
            setStatus(WM_THREAD_ERROR_OPEN_QUERY1, "������ ��� ������� ��������� ������ 1.\n" + e.Message);
            this->Terminate();
        }

        /*if (!this->Terminated && OraQueryMain == NULL)
        {
            setStatus(WM_THREAD_ERROR_OPEN_QUERY);
            this->Terminate();
        }*/
    }

    // ��������� ������ ������
    if (!this->Terminated && _queryText2 != "")     // ���� ����� ������ ������
    {    // ������� ��������� ��������������� ������
        try
        {
            _oraQuery2 = OpenOraQuery(_oraSession2, _queryText2, false);
        }
        catch (Exception &e)
        {
            setStatus(WM_THREAD_ERROR_OPEN_QUERY2, "������ ��� ������� ��������� ������ 2.\n" + e.Message);
            this->Terminate();
        }

        /*if (OraQuerySecondary == NULL)
        {
            setStatus(WM_THREAD_ERROR_OPEN_QUERY2);
            this->Terminate();
        } */
    }

    // ��������� ������ ������
    if (!this->Terminated && _queryText3 != "")     // ���� ����� ������ ������
    {    // ������� ��������� ��������������� ������
        try
        {
            _oraQuery3 = OpenOraQuery(_oraSession3, _queryText3, false);
        }
        catch (Exception &e)
        {
            setStatus(WM_THREAD_ERROR_OPEN_QUERY3, "������ ��� ������� ��������� ������ 3.\n" + e.Message);
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

        _oraQuery1->FetchAll = true;

	    RecCount = _oraQuery1->RecordCount;

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

            switch (ExportMode)
            {
            case EM_EXCEL_TEMPLATE:
                CoInitialize(NULL);
                //DoExportToExcelTemplate();
                DoExportToExcel(true);
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
        if ( _oraQuery1 != NULL )
        {
            try
            {
                _oraQuery1->Close();
                delete _oraQuery1;
                _oraQuery1 = NULL;
            }
            catch (...)
            {
            }
        }

        // �������� ���������������� �������
        if ( _oraQuery2 != NULL )
        {
            try
            {
                _oraQuery2->Close();
                delete _oraQuery2;
                _oraQuery2 = NULL;
            }
            catch (...)
            {
            }
        }
        setStatus(WM_THREAD_COMPLETED_SUCCESSFULLY);
    }

    Synchronize(SyncThreadChangeStatus);
}




 /*procedure TDataModule1.CopyField(Field: TField);
  var 
    NewField: TField; 
  begin 
    case Field.DataType of 
      ftString: NewField := TStringField.Create(Self); 
      ftDateTime: NewField := TDataTimeField.Create(Self); 
      // for each DataType there are an option on this case... 
    end; 
    NewField.FieldName := Field.FieldName;
    NewField.Lookup := Field.Lookup;
    // there are too much code here to copy all properties I need...
  end;
  */



/* ��������� �������� � ������ ���� Excel
*/
void __fastcall TThreadSelect::DoExportToExcel(bool toTemplate)
{

    TDataSet* mainDs;
    TDataSet* slaveDs = _oraQuery3;

    // ���� ���������� ��������� ��� �������
    if (param_excel.link_field_left != "" && param_excel.link_field_right != "")
    {
        TDataSet* ds = JoinDataset(_oraQuery1, _oraQuery2, param_excel.link_field_left, param_excel.link_field_right);
        // ////////////////////////////////////// ����� � ���� ������� ������� �������� ���� �� right � ����������� � ������� ��������
        // ////////////////////////////////////// ���������������� ����� ���� ����� ���������� ����� ����� ������� ����������

        mainDs = ds;
        // //////////////////////////////////////  ������� ds ����� �������������
    }
    else
    {
        mainDs = _oraQuery1;
    }

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

        for (int i=0; i <= param_count-1; i++)
        {
            if ( !UserParams[i]->isVisible() )
            {
                continue;
            }

            if (UserParams[i]->type != "separator")
            {
                data_parameters.PutElement(Variant(UserParams[i]->getCaption() + ": " + UserParams[i]->getDisplay()), i+1, 1);
            }
            else
            {
                data_parameters.PutElement(Variant("[" + UserParams[i]->getCaption() + "]"), i+1, 1);
            }
        }
    }


    Variant sql_text;
    {
        // ������ ��� sql-������
        Variant range_sqltext;
        int PartMaxLength = 4000;  // 8 192  - ������������ ����� ������ � ������ EXCEL
        int n = ceil( (float) _queryText1.Length() / PartMaxLength);

        sql_text = CreateVariantArray(n, 1);

        for (int i = 1; i <= n; i++)
        {
            String sQueryPart = _queryText1.SubString(((i-1) * PartMaxLength) + 1, PartMaxLength);
            sql_text.PutElement(Variant(sQueryPart), i, 1);
        }
    }

    // ��������� ��������� ������ � ������ �������
    // ���������!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! 2017-09-26
    if ( toTemplate != true )
    {                           // � ������ ��������
        param_excel.addTableDataSet(mainDs, "table_body", "table_column_");
        if ( param_count > 0 )
        {
            param_excel.addTableVtArray(data_parameters, "report_parameters");
        }
        param_excel.addTableVtArray(sql_text, "report_query_text");

        documentWriter.ExportToExcel(&param_excel);
    }
    else                        // � ������� ������
    {
        param_excel.addSingleDataSet(slaveDs, "");
        param_excel.addTableDataSet(mainDs, param_excel.table_range_name, "");
        documentWriter.ExportToExcelTemplate(&param_excel);
    }

    _resultFiles = documentWriter._result.resultFiles;

}

/* ��������� �������� � ������ Word
   ����������: ������� ��������� � ��������� ������ */
void __fastcall TThreadSelect::DoExportToWordTemplate()
{
    if ( _oraQuery1->RecordCount == 0)
    {
        return;
    }

    TWordExportParams* wordExportParams = &param_word;

    //TWordExportParams wordExportParams;
    //wordExportParams.pagePerDocument = param_word.pagePerDocument;

    /* ������������ ��������� ������ */
    if (_oraQuery1 != NULL)
    {
        wordExportParams->addSingleTextDataSet(_oraQuery1);
        //wordExportParams->addTableDataSet(_oraQuery1);        // ���� ������ ������� 2017-11-07 �������� - �������� � ��������� ������ ����������� �������
        wordExportParams->addFormtextDataSet(_oraQuery1);     // ����� ���������� �� �������
        wordExportParams->addMergeDataSet(_oraQuery1);
        wordExportParams->addSingleImageDataSet(_oraQuery1, "img_");     // ����� ���������� �� �������
    }

    if (_oraQuery2 != NULL)
    {
        wordExportParams->addSingleTextDataSet(_oraQuery2);
        //wordExportParams->addTableDataSet(_oraQuery2);        // ���� ������ ������� 2017-11-07 �������� - �������� � ��������� ������ ����������� �������
        wordExportParams->addFormtextDataSet(_oraQuery2);     // ����� ���������� �� �������
        wordExportParams->addMergeDataSet(_oraQuery2);
        wordExportParams->addSingleImageDataSet(_oraQuery2, "img_");     // ����� ���������� �� �������
    }
    if (_oraQuery3 != NULL)
    {
        wordExportParams->addSingleTextDataSet(_oraQuery3);
        //wordExportParams->addTableDataSet(_oraQuery3);        // ���� ������ ������� 2017-11-07 �������� - �������� � ��������� ������ ����������� �������
        wordExportParams->addFormtextDataSet(_oraQuery3);     // ����� ���������� �� �������
        wordExportParams->addMergeDataSet(_oraQuery3);
        wordExportParams->addSingleImageDataSet(_oraQuery3, "img_");     // ����� ���������� �� �������
    }

    //wordExportParams.addSingleTextDataSet(OraQuerySecondary, "rec_");  // ���������� �� �������
    //wordExportParams->templateFilename =  param_word.templateFilename;


    


    bool bFilterExist = param_word.filter_main_field != "" && param_word.filter_sec_field != "";    // ���� � ���������� ����� ������, �� �������, ��� ���������� ������

    if ( bFilterExist ) // ���� ����� ������
    {
         /* 2017-11-07 ���������� ������� dataset
        int i = 1;
        while( !_oraQuery3->Eof )
        {
            // ���� ����� ������, ��������� ��� � ��������� �������
            try
            {
                String sFilter = param_word.filter_main_field + "='" + _oraQuery3->FieldByName(param_word.filter_sec_field)->AsString + "'";
                _oraQuery1->Filtered = false;
                _oraQuery1->Filter = sFilter;
                _oraQuery1->Filtered = true;
            }
            catch ( Exception &e )
            {
                _oraQuery1->Filtered = false;
                _threadStatus = WM_THREAD_ERROR_IN_PROCESS;
                _threadMessage = "��������� ������������ ���������� ������� � ���������� �������� ��� ���������� � ���������� ��������������.\n" + e.Message;
                break;
            }

            if ( _oraQuery1->RecordCount > 0 )         // ���� ���� ������, �� ��������� ��������
            {
                wordExportParams->resultFilename = ExtractFilePath(DstFileName) + ExtractFileName(DstFileName) + "_" + IntToStr(i++)+ "_[:counter].doc";
                documentWriter.ExportToWordTemplate(wordExportParams);
                _resultFiles = documentWriter._result.resultFiles;
            }
            _oraQuery3->Next();
        }
        */
    }
    else
    {   // ���� ������ �� �����, ������ ������ ��� ����
        //wordExportParams->resultFilename = ExtractFilePath(DstFileName) + ExtractFileName(DstFileName) + "_[:counter].doc";
        wordExportParams->resultFilename = ExtractFilePath(DstFileName) + ExtractFileName(DstFileName);// + ".docx";
        documentWriter.ExportToWordTemplate(wordExportParams);
        _resultFiles = documentWriter._result.resultFiles;
    }
}

/* ��������� �������� � dbf-����
   ����������: ������� ��������� � ��������� ������ */
void __fastcall TThreadSelect::DoExportToDbf()
{
    if ( _oraQuery1->RecordCount == 0)
    {
        return;
    }
    param_dbase.srcDataSet = _oraQuery1;  // �������� ��������� � SetThreadOpt 2017-09-08

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


