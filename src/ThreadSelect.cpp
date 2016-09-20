//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
#pragma package(smart_init)

//#include "ThreadSelect.h"
#include "FMain.h"  // ������ ����� �������� ��������� ������ � ���� �����


//---------------------------------------------------------------------------
__fastcall ThreadSelect::ThreadSelect(bool CreateSuspended)
    : TThread(CreateSuspended)
{
    FreeOnTerminate = true;
    Suspended = true;
    //WParamResultMessage = 0;
    //LParamResultMessage = 0;
    AppPath = ExtractFilePath(Application->ExeName);
    ThreadMessage = "";
}

//---------------------------------------------------------------------------
__fastcall ThreadSelect::~ThreadSelect()
{
    if (ThreadOraSession != NULL) {
        ThreadOraSession->Disconnect();
        //ThreadOraSession->Close();
        delete ThreadOraSession;
    }

    if (ThreadOraSession2 != NULL && ThreadOraSession != ThreadOraSession2) {
        ThreadOraSession2->Disconnect();
        delete ThreadOraSession2;
    }

    ThreadOraSession = NULL;
    ThreadOraSession2 = NULL;

    vResultFiles.clear();
    //QueryParams.free();
}

//---------------------------------------------------------------------------
//
TOraSession* __fastcall ThreadSelect::CreateOraSession(TOraSession* TemplateOraSession)
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
// ��������� ���������� ��� ���������� ������� � ���������� ������
void ThreadSelect::SetThreadOpt(THREADOPTIONS* threadopt)
{
    //m_th_opt = *threadopt;
    this->ParentFormHandle = threadopt->ParentFormHandle;   // Handle ������� �����
    this->sQueryText = threadopt->querytext;        // ����� �������
    this->sQueryText2 = threadopt->querytext2;      // ����� �������
    this->DstFileName = threadopt->dstfilename;     // ��� ��������������� �����
    this->ExportMode = threadopt->exportmode;       // ����� �������� _EXPORTMODE


    ThreadOraSession = CreateOraSession(threadopt->TemplateOraSession);
    ThreadOraSession2 = NULL;
    if (threadopt->TemplateOraSession2 != NULL) {
        if (threadopt->TemplateOraSession != threadopt->TemplateOraSession2)        // ���� ���������� � ������ ��
            ThreadOraSession2 = CreateOraSession(threadopt->TemplateOraSession2);   // �� ������� ����� ����������
        else                                                                        // ����� �������� ��������� �� ������ ����������
            ThreadOraSession2 = ThreadOraSession;
    }

    // !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    // !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    // !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    // !!!!!!!!!!!!!!!������� �������� ���������� �� �������!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    // !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    // !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!


    switch (threadopt->exportmode) {
    case EM_EXCEL_BLANK:
        // ������������ ������
        this->param_excel.title_label = threadopt->queryitem->param_excel.title_label != ""? threadopt->queryitem->param_excel.title_label : threadopt->queryitem->queryname;  // ������������ ������
        this->param_excel.title_height = threadopt->queryitem->param_excel.title_height;      // ������ ��������� ������� MS Excel
        this->param_excel.Fields = threadopt->queryitem->param_excel.Fields;  // ������ ����� ��� �������� � MS Excel
        this->UserParams = threadopt->queryitem->UserParams;
    case EM_EXCEL_TEMPLATE:
        this->param_excel.template_name = threadopt->queryitem->param_excel.template_name;        // ������������ ������ � ������
        this->param_excel.table_range_name = threadopt->queryitem->param_excel.table_range_name;
        this->param_excel.fUnbounded = threadopt->queryitem->param_excel.fUnbounded;        // ������������ ������ � ������
        //this->param_word.filter_main_field= threadopt->queryitem->param_word.filter_main_field;
        //this->param_word.filter_sec_field = threadopt->queryitem->param_word.filter_sec_field;
        //this->param_word.filter_infix_sec_field = threadopt->queryitem->param_word.filter_infix_sec_field;
        break;
    case EM_DBASE4_FILE:
        this->param_dbase.Fields = threadopt->queryitem->param_dbase.Fields;  // ������ ����� ��� �������� � DBF
        this->param_dbase.fAllowUnassignedFields = threadopt->queryitem->param_dbase.fAllowUnassignedFields;
        break;
    case EM_PROCEDURE:
        break;
    case EM_WORD_TEMPLATE:
        this->param_word.page_per_doc = threadopt->queryitem->param_word.page_per_doc;              // ���������� ������� �� �������� MS Word
        this->param_word.template_name = threadopt->queryitem->param_word.template_name;        // ������������ ������ � ������
        this->param_word.filter_main_field= threadopt->queryitem->param_word.filter_main_field;
        this->param_word.filter_sec_field = threadopt->queryitem->param_word.filter_sec_field;
        this->param_word.filter_infix_sec_field = threadopt->queryitem->param_word.filter_infix_sec_field;
        break;
    }
}

//---------------------------------------------------------------------------
void __fastcall ThreadSelect::Execute()
{
    if (!ThreadOraSession->Connected) {
        ThreadStatus = WM_THREAD_ERROR_BD_CANT_CONNECT;
        ThreadMessage = "��������� ������. ���� ������ ����������.";
        this->Terminate();
    }

    if (!this->Terminated) {
        ThreadStatus = WM_THREAD_PROCEED_BEGIN_SQL;
        Synchronize(SyncThreadChangeStatus);

        if (ExportMode == EM_PROCEDURE) {
            // ��������� ��� ���������
            try {
                OraQueryMain = new TOraQuery(NULL);
                OraQueryMain->FetchAll = true;
                OraQueryMain->Session = ThreadOraSession;
                OraQueryMain->SQL->Add(sQueryText);
                OraQueryMain->Execute();

                ThreadStatus = WM_THREAD_EXECUTE_DONE;
            } catch (Exception &e) {
                ThreadMessage = e.Message;
                ThreadStatus = WM_THREAD_EXECUTE_ERROR;
            }
            try {
                delete OraQueryMain;
            } catch (...) {
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

    if (!this->Terminated && sQueryText != "") {   // ���� ����� ������ ������
        // ������� �������� �������� ������
        try {
            OraQueryMain = OpenOraQuery(ThreadOraSession, sQueryText, false);
        } catch (Exception &e) {
            ThreadMessage = "������ ��� ������� ��������� ������.\n" + e.Message;
        }

        if (!this->Terminated && OraQueryMain == NULL) {
            ThreadStatus = WM_THREAD_ERROR_OPEN_QUERY;
            this->Terminate();
        }
    }

    if (!this->Terminated && sQueryText2 != "") {    // ���� ����� ������ ������
        // ������� �������� ��������������� ������
        try {
            OraQuerySecondary = OpenOraQuery(ThreadOraSession2, sQueryText2, false);
        } catch (Exception &e) {
            ThreadMessage = "������ ��� ������� ��������� ������.\n" + e.Message;
        }
        if (OraQuerySecondary == NULL) {
            ThreadStatus = WM_THREAD_ERROR_OPEN_QUERY2;
            this->Terminate();
        }
    }

    if (!this->Terminated) {
        // ���������� ������ �� �������
        ThreadStatus = WM_THREAD_PROCEED_BEGIN_FETCH;
        Synchronize(SyncThreadChangeStatus);
    }

    if (!this->Terminated) {
        int RecCount = 0;

        OraQueryMain->FetchAll = true;

	    RecCount = OraQueryMain->RecordCount;

        if (RecCount <= 0) // ���� ������ �� ������ �������
        {
            ThreadStatus = WM_THREAD_ERROR_NULL_RESULTS;
            this->Terminate();
        }
        else if (RecCount > 200000 && RecCount < 1000000) // ���� ������ ������ ����� 200 000 �������
        {
            AnsiString msg = "� ������ �������� ���������� �������� " + IntToStr(RecCount) +" �����.\n������������ ������ ����� ������ ���������� �����.\n������������ �����?";
            if (MessageBoxQuestion(msg) == IDNO) {
                ThreadStatus = WM_THREAD_USER_CANCEL;
                this->Terminate();
            }
        }
        else if (RecCount >= 1000000)
        {
            ThreadStatus = WM_THREAD_ERROR_TOO_MORE_RESULTS;
            this->Terminate();
        }
    }

    // �������� ���������
    if (!this->Terminated)
    {
        try {
            ThreadStatus = WM_THREAD_PROCEED_BEGIN_DOCUMENT;
            Synchronize(SyncThreadChangeStatus);

            switch (ExportMode) {
            case EM_EXCEL_TEMPLATE:
                ExportToExcelTemplate(OraQueryMain, OraQuerySecondary);
                break;
            case EM_EXCEL_BLANK:
                ExportToExcel(OraQueryMain);
                break;
            //case EM_EXCEL_FILE_TEMPLATE:
            //    break;
            //case EM_EXCEL_FILE:
            //case EM_EXCEL_MEMORY:
                //ExportToExcel(OraQueryMain);
                //break;
            case EM_DBASE4_FILE:
                ExportToDBF(OraQueryMain);
                break;
            case EM_WORD_TEMPLATE:
                ExportToWordTemplate(OraQueryMain, OraQuerySecondary);
                break;
            }
        } catch (Exception& e) {
            ThreadMessage = e.Message;
            ThreadStatus = WM_THREAD_ERROR_IN_PROCESS;
            this->Terminate();
        }
        if (ThreadStatus != WM_THREAD_PROCEED_BEGIN_DOCUMENT) {
            this->Terminate();
        }
    }

    // �������� ��������� �������
    if (!this->Terminated) {
        if (OraQueryMain != NULL) {
            try {
                OraQueryMain->Close();
                delete OraQueryMain;
                OraQueryMain = NULL;
            } catch (...) {
            }
        }

        // �������� ���������������� �������
        if (OraQuerySecondary != NULL) {
            try {
                OraQuerySecondary->Close();
                delete OraQuerySecondary;
                OraQuerySecondary = NULL;
            } catch (...) {
            }
        }
        ThreadStatus = WM_THREAD_PROCEED_DONE;
    }
    if (ThreadStatus == WM_THREAD_PROCEED_DONE) {
        Synchronize(SyncThreadDone);
    } else {
        Synchronize(SyncThreadChangeStatus);
    }
}

//---------------------------------------------------------------------------
// ������������� - ��������� ������� ���������� �������
void __fastcall ThreadSelect::SyncThreadChangeStatus()
{
    Form1->OnThread(ThreadStatus, ThreadMessage);
}

//---------------------------------------------------------------------------
// ������������� - ����������� ������� ��������� �������, �������� ������ ������ - �����������
void __fastcall ThreadSelect::SyncThreadDone()
{
    Form1->OnThread(ThreadStatus, ThreadMessage);
    Form1->OnThreadSuccess(ExportMode, vResultFiles);
}

//---------------------------------------------------------------------------
// ���������� ������� MS Word
// QueryMerge - �������� ������, ������������ � �������� ��������� ������ ��� �������
// QueryFormFields - ��������������� ������, ������������ � �������� ��������� ������
//   ��� ������ ����� FormFields � ������� MS Word. ����� ���� NULL.
void __fastcall ThreadSelect::ExportToWordTemplate(TOraQuery *QueryMerge, TOraQuery *QueryFormFields)
{
    CoInitialize(NULL);

    String TemplateFullName = AppPath + param_word.template_name; // ���������� ���� � �����-�������
    String SavePath = ExtractFilePath(DstFileName);         // ���� ��� ���������� �����������
    String ResultFileNamePrefix = ExtractFileName(DstFileName);     // ������� ����� �����-����������

    std::vector<String> vFormFields;    // ������ � ������� ������ - �����������

    int FieldCount = QueryMerge->FieldCount;

    MSWordWorks msword;
    Variant Document;   // ������

    try {
        msword.OpenWord();
    } catch (Exception &e) {
        ThreadMessage = "��������� ������� ��������� ���������� Microsoft Word."
            "\n����������, ���������� � ���������� ��������������.\n" + e.Message;
        throw Exception(ThreadMessage);
    }

    try {
        //msword.SetVisible(true);
        Document =  msword.OpenDocument(TemplateFullName, false);
        //msword.SetDisplayAlerts(true);
    } catch (Exception &e) {
        msword.CloseApplication();
        VarClear(Document);
        ThreadMessage = "��������� ������� ������ " + TemplateFullName +
            "\n����������, ���������� � ���������� ��������������.\n" + e.Message;
        throw Exception(ThreadMessage);
    }

    vFormFields = msword.GetFormFields(Document);
    int FormFieldsCount = vFormFields.size();

    // ??? ����� �� ��������� QueryFormFields->RecordCount ?
    bool bFilterExist = param_word.filter_main_field != "" && param_word.filter_sec_field != "";    // ���� � ���������� ����� ������, �� �������, ��� ���������� ������


    if (QueryFormFields == NULL) {
        // ���� ����� ���� ������, �� ������ ������ �������
        // ������� ��������� Word � ��������
        if (QueryMerge->RecordCount > 0) {
            vResultFiles = msword.ExportToWordFields(QueryMerge, Document, SavePath, ResultFileNamePrefix + "_", param_word.page_per_doc);
        }
    } else {
        // ���� ������ ��� �������, ��:
        // 1. ���� ����� ������ � ����� ������ ������ ��������� �������
        // 2. ����������� �������� � FormFields-���� � �������
        // 3. ������ �������
        int n_doc = 0;  // ���������� ����� ��������� ������� (������������ � ����� ������ �����������)
        int nPadLength = IntToStr(QueryFormFields->RecordCount).Length();
        for( ; !QueryFormFields->Eof; QueryFormFields->Next()) { // ���� �� ����� FormFields

            // ���� ����� ������, ��������� ��� � ��������� �������
            if (bFilterExist) { // ���� �� ��������������� ������� ������ 1 ������, �� ��������� ������
                try {
                    String sFilter = param_word.filter_main_field + "='" + QueryFormFields->FieldByName(param_word.filter_sec_field)->AsString + "'";
                    QueryMerge->Filtered = false;
                    QueryMerge->Filter = sFilter;
                    QueryMerge->Filtered = true;
                } catch (Exception &e) {
                    QueryMerge->Filtered = false;
                    ThreadStatus = WM_THREAD_ERROR_IN_PROCESS;
                    ThreadMessage = "��������� ������������ ���������� ������� � ���������� �������� ��� ���������� � ���������� ��������������.\n" + e.Message;
                    break;
                }

            if (QueryMerge->RecordCount == 0)   // ���� ��� �������, �� ��������� ��� �����
                continue;
            }

            if (VarIsEmpty(Document))           // ���� ������ �� ������, ��������� ��� (��������� �� ������ ���� �����)
                Document = msword.OpenDocument(TemplateFullName, false);

            // ����������� �������� FormFields ����� � ������� MS Word
            for (int i = 0; i < FormFieldsCount ; i++)  { 		// ���������� FormFields, ����������� ��������������� �������� �� QueryFormFields
                String FormFieldName = vFormFields[i];
                try {
                    if ( FormFieldName.Pos("[IMG]") == 1 ) { // ����� ��������, ��������� ��������� [IMG WIDTH=150 HEIGHT=200]
                        String FieldName = FormFieldName.SubString(6, FormFieldName.Length()-5);
                        String ImgPath = "";
                        TField* Field = QueryFormFields->Fields->FindField(FieldName);
                        if (Field) {
                            ImgPath = AppPath + QueryFormFields->Fields->FieldByName(FieldName)->AsString;


                            // ��������� ����������� � ����
                            if (FileExists(ImgPath))
                            {
                                msword.SetPictureToField(Document, FieldName, ImgPath);
                            //msword.SetPictureToField(Document, FieldName, ImgPath, 80, 80);
                            } else {
                                // ���� �� ������ ��� throw
                                msword.SetTextToFieldF(Document, FieldName, "���� ����������� �� ������! (" + ImgPath + ")");
                                //
                                // !!!!!!!!!!!!!!!!!!!!
                                // ����� ���� ����� throw ???

                            }
                        }
                    } else {
                        TField* Field = QueryFormFields->Fields->FindField(FormFieldName);
                        if (Field) {
                            msword.SetTextToFieldF(Document, FormFieldName, Field->AsString);
                        }
                    }
                } catch (Exception &e) {
                    ThreadStatus = WM_THREAD_ERROR_IN_PROCESS_ALT;
                    ThreadMessage = "�������� ������ ��� ������ ����� FormFields � ������� """ + TemplateFullName + """"
                        ", ���� """ + FormFieldName + """."
                        "\n���������� � ���������� ��������������."
                        "\n" + e.Message;
                    break;
                }
            }

            // ��������� ������ � ����� �����
            AnsiString sFileNameInfix;
            if (param_word.filter_infix_sec_field != "") {
                try {
                    // �������� ������ �� ���� QueryFormFields
                    sFileNameInfix = Trim(QueryFormFields->FieldByName(param_word.filter_infix_sec_field)->AsString);
                } catch (...) {
                    // ���� ��������� ������, �� ���������� ���������� ����� ���������� ������� n_doc
                }
                if (sFileNameInfix == "") {     //
                    sFileNameInfix = StrPadL(IntToStr(n_doc++), nPadLength, "0");
                }

            } else {
                // ���������� � �������� ������� ���������� ����� ���������� �������
                sFileNameInfix = StrPadL(IntToStr(n_doc++), nPadLength, "0");
            }

            // ������� ��������� Word � ��������
            if (QueryMerge->RecordCount > 0) {
                std::vector<AnsiString> vNew;

                try {
                    vNew = msword.ExportToWordFields(QueryMerge, Document, SavePath, ResultFileNamePrefix + "_" + sFileNameInfix + "_", param_word.page_per_doc);
                } catch (Exception &e) {
                    ThreadStatus = WM_THREAD_ERROR_IN_PROCESS;
                    ThreadMessage = "� �������� ������� ��������� � ���������� ������ ��������� ������."
                        "\n���������� � ���������� ��������������."
                        "\n" + e.Message;
                    break;
                }

                vResultFiles.insert(vResultFiles.end(), vNew.begin(), vNew.end());
                vNew.clear();
            }

            msword.CloseDocument(Document);
            VarClear(Document);

            if (!bFilterExist)  // ���� ������ �� ����������, ����� ������� �� �����
                break;
        }
    }

    if (!VarIsEmpty(Document)) {     // ���� ������ ������
        msword.CloseDocument(Document);
        VarClear(Document);
    }
    msword.CloseApplication();

    CoUninitialize();
}

//---------------------------------------------------------------------------
// ������������ ������ MS EXCEL
void __fastcall ThreadSelect::ExportToExcel(TOraQuery *OraQuery)
{
    CoInitialize(NULL);

    bool fDone = false;

    // ���������� ���������� �������
    OraQuery->Last();
	int RecCount = OraQuery->RecordCount;

    // ���������� ���������� �����
    int FieldCount = OraQuery->FieldCount;

    Variant data_body;
    Variant data_head;
    DATAFORMAT df_body;
    df_body.reserve(FieldCount);

    int ExcelFieldCount = param_excel.Fields.size();

    try {     // ����������� ������ �����, ������������ ����� �������, ����������� ���� ������
        //data_body = CreateVariantArray(RecCount, FieldCount);  // ������� ������ ��� �������
        //data_head = CreateVariantArray(1, FieldCount);  // ����� �������

        if (ExcelFieldCount >= FieldCount) {  // ���������� ���� ���� ���� � ExcelFields
            data_body = CreateVariantArray(RecCount, ExcelFieldCount);     // ������� ������ ��� �������
            data_head = CreateVariantArray(1, ExcelFieldCount);            // ����� �������

            for (unsigned int j = 0; j < ExcelFieldCount; j++) {
                data_head.PutElement(param_excel.Fields[j].name, 1, j+1);
                df_body.push_back(param_excel.Fields[j].format);
            }
        } else
        {
            data_body = CreateVariantArray(RecCount, FieldCount);  // ������� ������ ��� �������
            data_head = CreateVariantArray(1, FieldCount);         // ����� �������

            // ��������� ����� �������
            for (int j = 1; j <= FieldCount; j++ ) { 		// ���������� ��� ����

                TField* field = OraQuery->Fields->FieldByNumber(j);
                // ������ ������ �������� � ������� Excel
                AnsiString sCellFormat;

                data_head.PutElement(field->DisplayName, 1, j);
                switch (field->DataType) {  // ����� ������������ � ��������� (�������� ������� � ��.)
                case ftString:
                    sCellFormat = "@";
                    break;
                case ftTime:
                    sCellFormat = "��:��:��";
                    break;
                case ftDate:
                    sCellFormat = "��.��.����";
                    break;
                case ftDateTime:
                    sCellFormat = "��.��.����";
                    break;
                case ftCurrency: case ftFloat:
                    sCellFormat = "0.00";
                    break;
                case ftSmallint: case ftInteger: case ftLargeint:
                    sCellFormat = "0";
                    break;
                default:
                    sCellFormat = "@";
                }
                df_body.push_back(sCellFormat);
	        }
        }
    } catch (Exception &e) {
        VarClear(data_head);
        VarClear(data_body);
        CoUninitialize();
        ThreadMessage = e.Message;
        throw Exception(ThreadMessage);
        //fDone = true;
    }

    MSExcelWorks msexcel;
    Variant Workbook;
    Variant Worksheet1;

    if (!fDone) {           // ���������� ������� ������
        AnsiString s = "";
        OraQuery->First();	// ��������� � ������ ������ (�� ������ ������)
        VarArrayLock(data_body);
        int i = 1;          // ���������� ����� �������
	    while (!OraQuery->Eof)  {
		    for (int j = 1; j <= FieldCount; j++ ) {
        	    s = OraQuery->Fields->FieldByNumber(j)->AsString;
                data_body.PutElement(s, i, j);
            }
            OraQuery->Next();  // ��������� � ��������� ������
            i++;
	    }
        VarArrayUnlock(data_body);
       try {
            msexcel.OpenApplication();
            Workbook = msexcel.OpenDocument();
        } catch (Exception &e) {
            VarClear(data_head);
            VarClear(data_body);
            CoUninitialize();
            ThreadMessage = e.Message;
            throw Exception(ThreadMessage);
        }
        Worksheet1 = msexcel.GetSheet(Workbook, 1);
    }


    if (!fDone && !VarIsEmpty(Worksheet1)) { // ��������� �������� Excel
    //if (!VarIsEmpty(Worksheet1)) { // ��������� �������� Excel
        TDateTime DateTime = TDateTime::CurrentDateTime();

        CELLFORMAT cf_body;
        CELLFORMAT cf_head;
        CELLFORMAT cf_title;
        CELLFORMAT cf_createtime;
        CELLFORMAT cf_sql;

        cf_body.BorderStyle = CELLFORMAT::xlContinuous;
        cf_head.BorderStyle = CELLFORMAT::xlContinuous;
        cf_head.FontStyle = cf_head.FontStyle << CELLFORMAT::fsBold;

        cf_head.bWrapText = false;

        cf_title.FontStyle = cf_title.FontStyle << CELLFORMAT::fsBold;
        cf_createtime.bSetFontColor = true;
        cf_createtime.FontColor = clRed;
        cf_sql.bWrapText = false;

        // ���������� ������ ������
        //std::vector<MSExcelWorks::CELLFORMAT> formats;
        //formats = msexcel.GetDataFormat(ArrayDataBody, 1);
        //std::vector<AnsiString> DataFormat;
        //DataFormat = msexcel.GetDataFormat(ArrayDataBody, 1);
        //for (int i=0; i  < QueryParams.size(); i++) {
            //QueryParams[i].
        //}


        // ��������� ������, �� ��������� ����������, ��������� �������������
        // �������� � ������� ������� ������������� ��������� � ����� "separator",
        Variant data_parameters;
        int param_count = UserParams.size();  // ��������� ������
        int visible_param_count = 0;
        for (int i=0; i <= param_count-1; i++) {    // ������������ ���-�� ������������ ����������
            if ( UserParams[i]->isVisible() )
                visible_param_count++;
        }
        if (param_count > 0) {    // ������ ���������� ��� ������ � Excel
            data_parameters = CreateVariantArray(visible_param_count, 1);
            for (int i=0; i <= param_count-1; i++) {
                if ( !UserParams[i]->isVisible() ) {
                    continue;
                }
                if (UserParams[i]->type != "separator") {
                    data_parameters.PutElement(UserParams[i]->getCaption() + ": " + UserParams[i]->getDisplay(), i+1, 1);
                }
                else
                {
                    data_parameters.PutElement("[" + UserParams[i]->getCaption() + "]", i+1, 1);
                }
            }
        }

        // ����� ������ �� ���� Excel
        Variant range_title = msexcel.WriteToCell(Worksheet1, param_excel.title_label , 1, 1);
        Variant range_createtime = msexcel.WriteToCell(Worksheet1, "�� ��������� ��: " + DateTime.DateTimeString(), 2, 1);
        Variant range_parameters;
        if (param_count > 0)
            range_parameters = msexcel.WriteTable(Worksheet1, data_parameters, 3, 1);

        Variant range_tablehead = msexcel.WriteTable(Worksheet1, data_head, 3 + visible_param_count, 1);
        Variant range_tablebody = msexcel.WriteTable(Worksheet1, data_body, 4 + visible_param_count, 1, &df_body);

        msexcel.SetRangeFormat(range_tablehead, cf_head);
        msexcel.SetRangeFormat(range_tablebody, cf_body);
        msexcel.SetRangeFormat(range_title, cf_title);
        msexcel.SetRangeFormat(range_createtime, cf_createtime);
        if (param_count > 0)
            msexcel.SetRangeFormat(range_parameters, cf_createtime);



        Variant range_all = msexcel.GetRangeFromRange(range_tablehead, 1, 1, msexcel.GetRangeRowsCount(range_tablebody)+1, msexcel.GetRangeColumnsCount(range_tablebody));


        if (this->param_excel.title_height > 0) {
            msexcel.SetRowHeight(range_tablehead, this->param_excel.title_height);    // ������ ������ ��������� �������
        }
        msexcel.SetAutoFilter(range_all);   // �������� ����������
        msexcel.SetColumnsAutofit(range_all);  // ������ ����� �� �����������


        // ��������� ��������� (������ � ��)
        for (int i=0; i < ExcelFieldCount; i++) {  // ���������� ���� ���� ���� � ExcelFields
            //CELLFORMAT cf_cell;
            //cf_cell.bSetFontColor = true;
            //cf_cell.FontColor = clGreen;

            if (param_excel.Fields[i].bwraptext >= 0) {
                CELLFORMAT cf_cell;
                cf_cell.bWrapText = param_excel.Fields[i].bwraptext;
                msexcel.SetRangeFormat(range_tablehead, cf_cell, 1, i+1);
            }

            if (param_excel.Fields[i].width >= 0) {
                msexcel.SetColumnWidth(range_tablehead, i+1, param_excel.Fields[i].width);    // ������ ������ ��������� �������
                //msexcel.SetColumnWidth(range_tablehead, 1);    // ������ ������ ��������� �������
            }
        }

        //msexcel.SetRowsAutofit(range_tablehead);


        // ������� ����� sql-������� �� ������ ����
        Variant Worksheet2 = msexcel.GetSheet(Workbook, 2);


        Variant range_sqltext;
        int PartMaxLength = 4000;  // 8 192  - ������������ ����� ������ � ������ EXCEL
        int n = ceil( (float) sQueryText.Length() / PartMaxLength);
        for (int i = 1; i <= n; i++) {
            AnsiString sQueryPart = sQueryText.SubString(((i-1) * PartMaxLength) + 1, PartMaxLength);
            range_sqltext = msexcel.WriteToCell(Worksheet2, sQueryPart, i, 1);
            msexcel.SetRangeFormat(range_sqltext, cf_sql);
        }

        df_body.clear();

        if (DstFileName == "") {
            msexcel.SetVisible(Workbook);
        } else {
            msexcel.SaveDocument(Workbook, DstFileName);
            VarClear(Workbook);
            VarClear(Worksheet1);
            VarClear(Worksheet2);
            msexcel.CloseApplication();
            vResultFiles.push_back(DstFileName);
        }


        //if (ExportMode == EM_EXCEL_FILE) {
        //    msexcel.SaveAsDocument(Workbook, DstFileName);
        //    msexcel.CloseExcel();
        //} else {
            //Workbook.OlePropertySet("Name", "blabla");
        //    msexcel.SetVisibleExcel(true, true);
        //}
    }

    // ������������ ������
    VarClear(data_head);
    data_head = NULL;

    VarClear(data_body);
    data_body = NULL;

    CoUninitialize();

    if (fDone) {
        throw Exception("����������.");
    }

}

//---------------------------------------------------------------------------
// ���������� Excel ����� � �������������� ������� xlt
void __fastcall ThreadSelect::ExportToExcelTemplate(TOraQuery *QueryTable, TOraQuery *QueryFields)
{
    CoInitialize(NULL);

    String TemplateFullName = AppPath + param_excel.template_name; // ���������� ���� � �����-�������

    // ��������� ������ MS Excel
    MSExcelWorks msexcel;
    Variant Workbook;
    Variant Worksheet;

    try {
        msexcel.OpenApplication();
        Workbook = msexcel.OpenDocument(TemplateFullName);
        Worksheet = msexcel.GetSheet(Workbook, 1);
    } catch (Exception &e) {
        try {
            msexcel.CloseApplication();
        } catch (...) {}
        CoUninitialize();
        ThreadMessage = "������ ��� �������� �����-������� " + TemplateFullName + ".\n���������� � ���������� ��������������.";
        throw Exception(ThreadMessage);
    }

    // ������� ������ ������ �����
    try {
        if (QueryFields != NULL)
            msexcel.ExportToExcelFields(QueryFields, Worksheet);
    } catch (Exception &e) {
        msexcel.CloseApplication();
        CoUninitialize();
        ThreadMessage = e.Message;
        throw e;
    }

    // ����� ��������� ��������� �����
    try {
        if (QueryTable != NULL && param_excel.table_range_name != "") // ������ ���� ������ ��� ��������� �������� �����
            msexcel.ExportToExcelTable(QueryTable, Worksheet, param_excel.table_range_name, param_excel.fUnbounded);
    } catch (Exception &e) {
        try {
            msexcel.CloseApplication();
        } catch (...) {}
        CoUninitialize();
        ThreadMessage = e.Message;
        throw e;
    }

    if (DstFileName == "") {        // ������ ��������� ��������, ���� ��� �����-���������� �� ������
        msexcel.SetVisible(Workbook);
    } else {                        // ����� ��������� � ����
        try {
            msexcel.SaveDocument(Workbook, DstFileName);
            msexcel.CloseApplication();
            vResultFiles.push_back(DstFileName);
        } catch (Exception &e) {
            try {
                msexcel.CloseApplication();
            } catch (...) {}
            ThreadMessage = "������ ��� ���������� ���������� � ���� " + DstFileName + ".\n" + e.Message;
            throw Exception(ThreadMessage);
        }
    }

    // � ���������� ������� ���������� �������� � MS Word
    // ����������� ���� ������ QueryFields � QueryTable

    CoUninitialize();
}

//---------------------------------------------------------------------------
// ���������� DBF-�����
// ���������� ��� ������� � ������������� ���������� TDbf
void __fastcall ThreadSelect::ExportToDBF(TOraQuery *OraQuery)
{
    TStringList* ListFields;
    int n = this->param_dbase.Fields.size();
    if (n > 0) {   // ��������� ������ ����� ��� �������� � DBF ("���;���;�����;����� ������� �����")
        ListFields = new TStringList();
        for (int i = 0; i < n; i++) {
            ListFields->Add(param_dbase.Fields[i].name + ";" + param_dbase.Fields[i].type + ";"+ param_dbase.Fields[i].length + ";" + param_dbase.Fields[i].decimals);
        }
    } else {
        ThreadMessage = "�� ����� ������ ����� � ���������� ��������."
            "\n����������, ���������� � ���������� ��������������.";
        throw Exception(ThreadMessage);
    }

    // ��� ������� ������, � ����� � ���, ��� ��������� ���� ����� ���������� �������
    if (ListFields->Count > OraQuery->FieldCount && !param_dbase.fAllowUnassignedFields) {
        ThreadMessage = "���������� ��������� ����� ��������� ���������� ����� � ��������� ������."
            "\n����������, ���������� � ���������� ��������������.";
        throw Exception(ThreadMessage);
    }

    TCreateHalcyonDataSet* CreateDbfDS;
    CreateDbfDS = new TCreateHalcyonDataSet(NULL);

    THalcyonDataSet* HalcyonDataSet1;
    HalcyonDataSet1 = new THalcyonDataSet(NULL);

    HalcyonDataSet1->TableName = ExtractFileName(DstFileName);
    HalcyonDataSet1->DatabaseName = ExtractFilePath(DstFileName);

    CreateDbfDS->DBFTable = HalcyonDataSet1;
    CreateDbfDS->DBFType = DBaseIV;
    CreateDbfDS->AutoOverwrite = true;

    CreateDbfDS->CreateFields->Clear();
    CreateDbfDS->CreateFields->AddStrings(ListFields);
    CreateDbfDS->Execute();
    delete ListFields;

    HalcyonDataSet1->DisableControls();
    HalcyonDataSet1->Open();
    HalcyonDataSet1->CheckBrowseMode();
    HalcyonDataSet1->UpdateCursorPos();

    try {
	    while (!OraQuery->Eof)  {
            HalcyonDataSet1->Append();
		    for (int j = 1; j <= OraQuery->FieldCount; j++ ) {
                //HalcyonDataSet1->Edit();
                //AssignRecord(Table1, HalcyonDataSet1, True);
                HalcyonDataSet1->Fields->FieldByNumber(j)->Value = OraQuery->Fields->FieldByNumber(j)->Value;
            }
            OraQuery->Next();  // ��������� � ��������� ������
	    }
        HalcyonDataSet1->Post();
        HalcyonDataSet1->Close();

        vResultFiles.push_back(DstFileName);

    } catch(Exception &e) {
        HalcyonDataSet1->Close();


        delete HalcyonDataSet1;
        delete CreateDbfDS;

        ThreadMessage = e.Message;
        throw e;
    }

    delete HalcyonDataSet1;
    delete CreateDbfDS;
}

