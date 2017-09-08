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

/* Установка параметров для выполнения запроса и подготовки отчета */
void TThreadSelect::SetThreadOpt(THREADOPTIONS* threadopt)
{
    //m_th_opt = *threadopt;
    this->ParentFormHandle = threadopt->ParentFormHandle;   // Handle главной формы
    this->_reportName = threadopt->queryName;
    this->_mainQueryText = threadopt->querytext;            // Текст запроса
    this->_secondaryQueryText = threadopt->querytext2;      // Текст запроса
    this->DstFileName = threadopt->dstfilename;         // Имя результирующего файла
    this->ExportMode = threadopt->exportmode;           // Режим экспорта _EXPORTMODE


    ThreadOraSession = CreateOraSession(threadopt->TemplateOraSession);
    ThreadOraSession2 = NULL;
    if (threadopt->TemplateOraSession2 != NULL)
    {
        if (threadopt->TemplateOraSession != threadopt->TemplateOraSession2)        // Если соединения к разным БД
        {
            ThreadOraSession2 = CreateOraSession(threadopt->TemplateOraSession2);   // то создаем новое соединение
        }
        else                                                                        // иначе копируем указатель на первое соединение
        {
            ThreadOraSession2 = ThreadOraSession;
        }
    }

    // !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    // !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    // !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    // !!!!!!!!!!!!!!!сделать загрузку параметров из вектора!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    // !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    // !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!


    switch (threadopt->exportmode) {
    case EM_EXCEL_BLANK:
    {
        // Наименование отчета
        this->param_excel.title_label = threadopt->queryitem->param_excel.title_label != ""? threadopt->queryitem->param_excel.title_label : threadopt->queryitem->queryname;  // Наименование отчета
        this->param_excel.title_height = threadopt->queryitem->param_excel.title_height;      // Высота заголовка таблицы MS Excel
        this->param_excel.Fields = threadopt->queryitem->param_excel.Fields;  // Вектор полей для экспорта в MS Excel

        //int k = threadopt->queryitem->param_excel.Fields.size();

        this->UserParams = threadopt->queryitem->UserParams;
        break;
    }
    case EM_EXCEL_TEMPLATE:
    {
        this->param_excel = threadopt->queryitem->param_excel;        //
        /*this->param_excel.templateFilename = threadopt->queryitem->param_excel.templateFilename;        // Тестирование печать в шаблон
        this->param_excel.table_range_name = threadopt->queryitem->param_excel.table_range_name;   2017-09-08*/
        this->param_excel.fUnbounded = threadopt->queryitem->param_excel.fUnbounded;        // Тестирование печать в шаблон
        break;
    }
    case EM_DBASE4_FILE:
    {
        this->param_dbase = threadopt->queryitem->param_dbase;
        //this->param_dbase.Fields = threadopt->queryitem->param_dbase.Fields;  // Вектор полей для экспорта в DBF
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
        /*this->param_word.pagePerDocument = threadopt->queryitem->param_word.pagePerDocument;              // Количество страниц на документ MS Word
        this->param_word.template_name = threadopt->queryitem->param_word.template_name;        // Тестирование печать в шаблон
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
        setStatus(WM_THREAD_ERROR_BD_CANT_CONNECT, "Произошла ошибка. База данных недоступна.");
        this->Terminate();
    }

    if (!this->Terminated)
    {
        setStatus(WM_THREAD_PROCEED_BEGIN_SQL, _reportName);
        Synchronize(SyncThreadChangeStatus);

        if (ExportMode == EM_PROCEDURE)
        {
            // Выполнить как процедуру
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
            /*Здесь сделать блок try для анализа ошибки.
            Возможна ситуация, когда отсутствует доступ к таблицам, использующимся в запросе.
            Тогда необходимо выводить соответствующий текст об ошибке.
            Сейчас выводится "... Проверте правильность запроса."*/

    if (!this->Terminated && _mainQueryText != "")    // Если задан первый запрос
    {
        // Пробуем выполниь основной запрос
        try
        {
            OraQueryMain = OpenOraQuery(ThreadOraSession, _mainQueryText, false);
        }
        catch (Exception &e)
        {
            setStatus(WM_THREAD_ERROR_OPEN_QUERY, "Ошибка при попытке выполнить запрос.\n" + e.Message);
            this->Terminate();
        }

        /*if (!this->Terminated && OraQueryMain == NULL)
        {
            setStatus(WM_THREAD_ERROR_OPEN_QUERY);
            this->Terminate();
        }*/
    }

    if (!this->Terminated && _secondaryQueryText != "")     // Если задан второй запрос
    {    // Пробуем выполнить вспомогательный запрос
        try
        {
            OraQuerySecondary = OpenOraQuery(ThreadOraSession2, _secondaryQueryText, false);
        }
        catch (Exception &e)
        {
            setStatus(WM_THREAD_ERROR_OPEN_QUERY2, "Ошибка при попытке выполнить запрос.\n" + e.Message);
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
        // Извлечение данных из запроса
        setStatus(WM_THREAD_PROCEED_BEGIN_FETCH);
        Synchronize(SyncThreadChangeStatus);
    }

    if (!this->Terminated)
    {
        int RecCount = 0;

        OraQueryMain->FetchAll = true;

	    RecCount = OraQueryMain->RecordCount;

        if (RecCount <= 0) // Если запрос не вернул записей
        {
            setStatus(WM_THREAD_ERROR_NULL_RESULTS);
            this->Terminate();
        }
        else if (RecCount > 200000 && RecCount < 1000000) // Если запрос вернул более 200 000 записей
        {
            AnsiString msg = "С учетом заданных параметров получено " + IntToStr(RecCount) +" строк.\nФормирование отчета может занять длительное время.\nСформировать отчет?";
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

    // Создание документа
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

    // Закрытие основного запроса
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

        // Закрытие вспомогательного запроса
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

/* Процедура экспорта в чистый файл Excel
*/
void __fastcall TThreadSelect::DoExportToExcel()
{

    // Заполняем массив, со значеними параметров, заданными пользователем
    // Возможно в будущем сделать распознавание параметра с типом "separator",
    Variant data_parameters;
    int param_count = UserParams.size();  // Параметры отчета
    int visible_param_count = 0;

    for (int i=0; i <= param_count-1; i++)    // Подсчитываем кол-во отображаемых параметров
    {
        if ( UserParams[i]->isVisible() )
        {
            visible_param_count++;
        }
    }

    if (param_count > 0)     // Список параметров для вывода в Excel
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
        // Массив для sql-текста
        Variant range_sqltext;
        int PartMaxLength = 4000;  // 8 192  - максимальная длина строки в ячейке EXCEL
        int n = ceil( (float) _mainQueryText.Length() / PartMaxLength);

        sql_text = CreateVariantArray(n, 1);

        for (int i = 1; i <= n; i++)
        {
            String sQueryPart = _mainQueryText.SubString(((i-1) * PartMaxLength) + 1, PartMaxLength);
            sql_text.PutElement(sQueryPart, i, 1);
        }
    }

    // Добавляем источники данных
    param_excel.addTableDataSet(OraQueryMain, "table_body", "table_column_");
    param_excel.addTableVtArray(data_parameters, "report_parameters");
    param_excel.addTableVtArray(sql_text, "report_query_text");

    documentWriter.ExportToExcel(&param_excel);
    _resultFiles = documentWriter._result.resultFiles;

}


/* Процедура экспорта в шаблон Word
   Примечание: Следует перенести в отдельный модуль */
void __fastcall TThreadSelect::DoExportToWordTemplate()
{
    if ( OraQueryMain->RecordCount == 0)
    {
        return;
    }

    TWordExportParams* wordExportParams = &param_word;

    //TWordExportParams wordExportParams;
    //wordExportParams.pagePerDocument = param_word.pagePerDocument;

    /* Присоединяем источники данных */
    wordExportParams->addFormtextDataSet(OraQuerySecondary);     // Общая информация по участку
    //wordExportParams.addSingleTextDataSet(OraQuerySecondary, "rec_");  // Информация по реестру
    wordExportParams->addMergeDataSet(OraQueryMain);
    //wordExportParams->templateFilename =  param_word.templateFilename;


    bool bFilterExist = param_word.filter_main_field != "" && param_word.filter_sec_field != "";    // Если в параметрах задан фильтр, то считаем, что установлен фильтр

    if ( bFilterExist ) // Если задан фильтр
    {
        int i = 1;
        while( !OraQuerySecondary->Eof )
        {
            // Если задан фильтр, применяем его к основному запросу
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
                _threadMessage = "Проверьте корректность параметров фильтра в параметрах экспорта или обратитесь к системному администратору.\n" + e.Message;
                break;
            }

            if ( OraQueryMain->RecordCount > 0 )         // Если есть записи, то формируем документ
            {
                wordExportParams->resultFilename = ExtractFilePath(DstFileName) + ExtractFileName(DstFileName) + "_" + IntToStr(i++)+ "_[:counter].doc";
                documentWriter.ExportToWordTemplate(wordExportParams);
                _resultFiles = documentWriter._result.resultFiles;
            }
            OraQuerySecondary->Next();
        }
    }
    else
    {   // если фильтр не задан, делаем экпорт как есть
        wordExportParams->resultFilename = ExtractFilePath(DstFileName) + ExtractFileName(DstFileName) + "_[:counter].doc";
        documentWriter.ExportToWordTemplate(wordExportParams);
        _resultFiles = documentWriter._result.resultFiles;
    }
}

/* Процедура экспорта в dbf-файл
   Примечание: Следует перенести в отдельный модуль */
void __fastcall TThreadSelect::DoExportToDbf()
{
    if ( OraQueryMain->RecordCount == 0)
    {
        return;
    }
    param_dbase.srcDataSet = OraQueryMain;  // Возможно перенести в SetThreadOpt 2017-09-08

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

/* Синхронизация - для отладки */
/*void __fastcall TThreadSelect::SyncDebug()
{
    ShowMessage(debug_message);
}*/

/* Синхронизация - изменение статуса выполнения запроса */
void __fastcall TThreadSelect::SyncThreadChangeStatus()
{
    TThreadSelectMessage message(_threadStatus, _threadMessage, _resultFiles);
    Form1->threadListener(_threadId, message);
}


