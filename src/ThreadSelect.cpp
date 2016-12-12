//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
#pragma package(smart_init)

#include "ThreadSelect.h"
#include "FMain.h"

#include "dbf.hpp"
#include "Dbf_Lang.hpp"



unsigned int ThreadSelect::_threadIndex = 0;

/**/
//__fastcall ThreadSelect::ThreadSelect(bool CreateSuspended, THREADOPTIONS* threadopt, void (*f)(const String&, int))
__fastcall ThreadSelect::ThreadSelect(bool CreateSuspended, THREADOPTIONS* threadopt)
    : TThread(CreateSuspended),
    _threadMessage("")
{
    FreeOnTerminate = true;
    Suspended = true;
    //WParamResultMessage = 0;
    //LParamResultMessage = 0;
    AppPath = ExtractFilePath(Application->ExeName);

    SetThreadOpt(threadopt);
    _threadIndex++;

    randomize();
    _threadId = random(9999999999);
}

/**/
__fastcall ThreadSelect::~ThreadSelect()
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
void ThreadSelect::SetThreadOpt(THREADOPTIONS* threadopt)
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
        // Наименование отчета
        this->param_excel.title_label = threadopt->queryitem->param_excel.title_label != ""? threadopt->queryitem->param_excel.title_label : threadopt->queryitem->queryname;  // Наименование отчета
        this->param_excel.title_height = threadopt->queryitem->param_excel.title_height;      // Высота заголовка таблицы MS Excel
        this->param_excel.Fields = threadopt->queryitem->param_excel.Fields;  // Вектор полей для экспорта в MS Excel
        this->UserParams = threadopt->queryitem->UserParams;
    case EM_EXCEL_TEMPLATE:
        this->param_excel.template_name = threadopt->queryitem->param_excel.template_name;        // Тестирование печать в шаблон
        this->param_excel.table_range_name = threadopt->queryitem->param_excel.table_range_name;
        this->param_excel.fUnbounded = threadopt->queryitem->param_excel.fUnbounded;        // Тестирование печать в шаблон
        //this->param_word.filter_main_field= threadopt->queryitem->param_word.filter_main_field;
        //this->param_word.filter_sec_field = threadopt->queryitem->param_word.filter_sec_field;
        //this->param_word.filter_infix_sec_field = threadopt->queryitem->param_word.filter_infix_sec_field;
        break;
    case EM_DBASE4_FILE:
        this->param_dbase.Fields = threadopt->queryitem->param_dbase.Fields;  // Вектор полей для экспорта в DBF
        this->param_dbase.fAllowUnassignedFields = threadopt->queryitem->param_dbase.fAllowUnassignedFields;
        break;
    case EM_PROCEDURE:
        break;
    case EM_WORD_TEMPLATE:
        this->param_word.page_per_doc = threadopt->queryitem->param_word.page_per_doc;              // Количество страниц на документ MS Word
        this->param_word.template_name = threadopt->queryitem->param_word.template_name;        // Тестирование печать в шаблон
        this->param_word.filter_main_field= threadopt->queryitem->param_word.filter_main_field;
        this->param_word.filter_sec_field = threadopt->queryitem->param_word.filter_sec_field;
        this->param_word.filter_infix_sec_field = threadopt->queryitem->param_word.filter_infix_sec_field;
        break;
    }
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
void __fastcall ThreadSelect::Execute()
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
    {    // Пробуем выполниь вспомогательный запрос
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
        }
        catch ( Exception& e )
        {
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

void __fastcall ThreadSelect::setStatus(_TThreadStatus status, const AnsiString& message)
{
    _threadStatus = status;
    _threadMessage = message;

}

/* Синхронизация - изменение статуса выполнения запроса */
void __fastcall ThreadSelect::SyncThreadChangeStatus()
{
    TThreadSelectMessage message(_threadStatus, _threadMessage, _resultFiles);
    Form1->threadListener(_threadId, message);
}

/*
 Заполнение шаблона MS Word
 QueryMerge - основной запрос, используется в качестве источника данных при слиянии
 QueryFormFields - вспомогательный запрос, используется в качестве источника данных
 при замене полей FormFields в шаблоне MS Word. Может быть NULL.
 */
void __fastcall ThreadSelect::ExportToWordTemplate(TOraQuery *QueryMerge, TOraQuery *QueryFormFields)
{
    CoInitialize(NULL);

    String TemplateFullName = AppPath + param_word.template_name; // Абсолютный путь к файлу-шаблону
    String SavePath = ExtractFilePath(DstFileName);         // Путь для сохранения результатов
    String ResultFileNamePrefix = ExtractFileName(DstFileName);     // Префикс имени файла-результата

    std::vector<String> vFormFields;    // Вектор с именами файлов - результатов

    int FieldCount = QueryMerge->FieldCount;

    MSWordWorks msword;
    Variant Document;   // Шаблон

    try {
        msword.OpenWord();
    } catch (Exception &e) {
        _threadMessage = "Неудалось создать экземпляр приложения Microsoft Word."
            "\nПожалуйста, обратитесь к системному администратору.\n" + e.Message;
        throw Exception(_threadMessage);
    }

    try {
        //msword.SetVisible(true);
        Document =  msword.OpenDocument(TemplateFullName, false);
        //msword.SetDisplayAlerts(true);
    } catch (Exception &e) {
        msword.CloseApplication();
        VarClear(Document);
        _threadMessage = "Неудалось открыть шаблон " + TemplateFullName +
            "\nПожалуйста, обратитесь к системному администратору.\n" + e.Message;
        throw Exception(_threadMessage);
    }

    vFormFields = msword.GetFormFields(Document);
    int FormFieldsCount = vFormFields.size();

    // ??? Нужно ли учитывать QueryFormFields->RecordCount ?
    bool bFilterExist = param_word.filter_main_field != "" && param_word.filter_sec_field != "";    // Если в параметрах задан фильтр, то считаем, что установлен фильтр


    if (QueryFormFields == NULL) {
        // Если задан один запрос, то делаем только слияние
        // Слияние документа Word с таблицей
        if (QueryMerge->RecordCount > 0) {
            _resultFiles = msword.ExportToWordFields(QueryMerge, Document, SavePath, ResultFileNamePrefix + "_", param_word.page_per_doc);
        }
    } else {
        // Если задано два запроса, то:
        // 1. если задан фильтр в цикле задаем фильтр основному запросу
        // 2. подставляем значения в FormFields-поля в шаблоне
        // 3. делаем слияние
        int n_doc = 0;  // Порядковый номер процедуры слияния (используется в имени файлов результатов)
        int nPadLength = IntToStr(QueryFormFields->RecordCount).Length();
        for( ; !QueryFormFields->Eof; QueryFormFields->Next()) { // Цикл по полям FormFields

            // Если задан фильтр, применяем его к основному запросу
            if (bFilterExist)  // Если во вспомогательном запросе больше 1 строки, то применяем фильтр
            {
                try
                {
                    String sFilter = param_word.filter_main_field + "='" + QueryFormFields->FieldByName(param_word.filter_sec_field)->AsString + "'";
                    QueryMerge->Filtered = false;
                    QueryMerge->Filter = sFilter;
                    QueryMerge->Filtered = true;
                }
                catch ( Exception &e )
                {
                    QueryMerge->Filtered = false;
                    _threadStatus = WM_THREAD_ERROR_IN_PROCESS;
                    _threadMessage = "Проверьте корректность параметров фильтра в параметрах экспорта или обратитесь к системному администратору.\n" + e.Message;
                    break;
                }

                if (QueryMerge->RecordCount == 0)         // Если нет записей, то следующий шаг цикла
                {
                    continue;
                }
            }

            if (VarIsEmpty(Document))           // Если шаблон не открыт, открываем его (требуется на втором шаге цикла)
            {
                Document = msword.OpenDocument(TemplateFullName, false);
            }

            // Подставляем значения FormFields полей в шаблоне MS Word
            for (int i = 0; i < FormFieldsCount ; i++)   		// Перебираем FormFields, подставляем соответствующие значения из QueryFormFields
            {
                String FormFieldName = vFormFields[i];
                try {
                    if ( FormFieldName.Pos("[IMG]") == 1 ) // ЗДЕСЬ ДОДЕЛАТЬ, учитывать параметры [IMG WIDTH=150 HEIGHT=200]
                    {
                        String FieldName = FormFieldName.SubString(6, FormFieldName.Length()-5);
                        String ImgPath = "";
                        TField* Field = QueryFormFields->Fields->FindField(FieldName);
                        if (Field) {
                            ImgPath = AppPath + QueryFormFields->Fields->FieldByName(FieldName)->AsString;


                            // Вставляем изображение в поле
                            if (FileExists(ImgPath))
                            {
                                msword.SetPictureToField(Document, FieldName, ImgPath);
                                //msword.SetPictureToField(Document, FieldName, ImgPath, 80, 80);
                            } else {
                                // Файл не найден или throw
                                msword.SetTextToFieldF(Document, FieldName, "Файл изображения не найден! (" + ImgPath + ")");
                                //
                                // !!!!!!!!!!!!!!!!!!!!
                                // Может быть лучше throw ???

                            }
                        }
                    } else {
                        TField* Field = QueryFormFields->Fields->FindField(FormFieldName);
                        if (Field) {
                            msword.SetTextToFieldF(Document, FormFieldName, Field->AsString);
                        }
                    }
                } catch (Exception &e) {
                    _threadStatus = WM_THREAD_ERROR_IN_PROCESS_ALT;
                    _threadMessage = "Возникла ошибка при замене полей FormFields в шаблоне """ + TemplateFullName + """"
                        ", поле """ + FormFieldName + """."
                        "\nОбратитесь к системному администратору."
                        "\n" + e.Message;
                    break;
                }
            }

            // Формируем инфикс к имени файла
            AnsiString sFileNameInfix;
            if (param_word.filter_infix_sec_field != "")
            {
                try
                {
                    // Получаем инфикс из поля QueryFormFields
                    sFileNameInfix = Trim(QueryFormFields->FieldByName(param_word.filter_infix_sec_field)->AsString);
                }
                catch (...)
                {
                    // Если произошла ошибка, то используем порядковый номер очередного слияния n_doc
                }
                if (sFileNameInfix == "") {     //
                    sFileNameInfix = StrPadL(IntToStr(n_doc++), nPadLength, "0");
                }

            }
            else
            {
                // Используем в качестве инфикса порядковый номер очередного слияния
                sFileNameInfix = StrPadL(IntToStr(n_doc++), nPadLength, "0");
            }

            // Слияние документа Word с таблицей
            if (QueryMerge->RecordCount > 0)
            {
                std::vector<AnsiString> vNew;

                try
                {
                    vNew = msword.ExportToWordFields(QueryMerge, Document, SavePath, ResultFileNamePrefix + "_" + sFileNameInfix + "_", param_word.page_per_doc);
                }
                catch (Exception &e)
                {
                    _threadStatus = WM_THREAD_ERROR_IN_PROCESS;
                    _threadMessage = "В процессе слияния документа с источником данных произошла ошибка."
                        "\nОбратитесь к системному администратору."
                        "\n" + e.Message;
                    break;
                }

                _resultFiles.insert(_resultFiles.end(), vNew.begin(), vNew.end());
                vNew.clear();
            }

            msword.CloseDocument(Document);
            VarClear(Document);

            if (!bFilterExist)  // Если фильтр не установлен, тогда выходим из цикла
            {
                break;
            }
        }
    }

    if (!VarIsEmpty(Document))      // Если шаблон открыт
    {
        msword.CloseDocument(Document);
        VarClear(Document);
    }
    msword.CloseApplication();

    CoUninitialize();
}

//---------------------------------------------------------------------------
// ФОРМИРОВАНИЕ ОТЧЕТА MS EXCEL
void __fastcall ThreadSelect::ExportToExcel(TOraQuery *OraQuery)
{
    CoInitialize(NULL);

    bool fDone = false;

    // Определяем количество записей
    OraQuery->Last();
	int RecCount = OraQuery->RecordCount;

    // Определяем количество полей
    int FieldCount = OraQuery->FieldCount;

    Variant data_body;
    Variant data_head;
    DATAFORMAT df_body;
    df_body.reserve(FieldCount);

    int ExcelFieldCount = param_excel.Fields.size();

    try {     // Определение списка полей, формирование шапки таблицы, определение типа данных
        //data_body = CreateVariantArray(RecCount, FieldCount);  // Создаем массив для таблицы
        //data_head = CreateVariantArray(1, FieldCount);  // Шапка таблицы

        if (ExcelFieldCount >= FieldCount)   // Заполнение если есть поля в ExcelFields
        {
            data_body = CreateVariantArray(RecCount, ExcelFieldCount);     // Создаем массив для таблицы
            data_head = CreateVariantArray(1, ExcelFieldCount);            // Шапка таблицы

            for (unsigned int j = 0; j < ExcelFieldCount; j++)
            {
                data_head.PutElement(param_excel.Fields[j].name, 1, j+1);
                df_body.push_back(param_excel.Fields[j].format);
            }
        }
        else
        {
            data_body = CreateVariantArray(RecCount, FieldCount);  // Создаем массив для таблицы
            data_head = CreateVariantArray(1, FieldCount);         // Шапка таблицы

            // Формируем шапку таблицы
            for (int j = 1; j <= FieldCount; j++ )  		// Перебираем все поля
            {
                TField* field = OraQuery->Fields->FieldByNumber(j);
                // Задаем формат столбцов в таблице Excel
                AnsiString sCellFormat;

                data_head.PutElement(field->DisplayName, 1, j);
                switch (field->DataType) {  // Нужно тестирование и доработка (добавить форматы и тд.)
                case ftString:
                    sCellFormat = "@";
                    break;
                case ftTime:
                    sCellFormat = "чч:мм:сс";
                    break;
                case ftDate:
                    sCellFormat = "ДД.ММ.ГГГГ";
                    break;
                case ftDateTime:
                    sCellFormat = "ДД.ММ.ГГГГ";
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
    }
    catch (Exception &e)
    {
        VarClear(data_head);
        VarClear(data_body);
        CoUninitialize();
        _threadMessage = e.Message;
        throw Exception(_threadMessage);
        //fDone = true;
    }

    MSExcelWorks msexcel;
    Variant Workbook;
    Variant Worksheet1;

    if (!fDone)           // Заполнение массива данных
    {
        AnsiString s = "";
        OraQuery->First();	// Переходим к первой записи (на всякий случай)
        VarArrayLock(data_body);
        int i = 1;          // Пропускаем шапку таблицы
	    while (!OraQuery->Eof)
        {
		    for (int j = 1; j <= FieldCount; j++ )
            {
        	    s = OraQuery->Fields->FieldByNumber(j)->AsString;
                data_body.PutElement(s, i, j);
            }
            OraQuery->Next();  // Переходим к следующей записи
            i++;
	    }
        VarArrayUnlock(data_body);
       try
       {
            msexcel.OpenApplication();
            Workbook = msexcel.OpenDocument();
        }
        catch (Exception &e)
        {
            VarClear(data_head);
            VarClear(data_body);
            CoUninitialize();
            _threadMessage = e.Message;
            throw Exception(_threadMessage);
        }
        Worksheet1 = msexcel.GetSheet(Workbook, 1);
    }


    if (!fDone && !VarIsEmpty(Worksheet1))  // Заполняем документ Excel
    {
    //if (!VarIsEmpty(Worksheet1)) { // Заполняем документ Excel
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

        // Определяем формат данных
        //std::vector<MSExcelWorks::CELLFORMAT> formats;
        //formats = msexcel.GetDataFormat(ArrayDataBody, 1);
        //std::vector<AnsiString> DataFormat;
        //DataFormat = msexcel.GetDataFormat(ArrayDataBody, 1);
        //for (int i=0; i  < QueryParams.size(); i++) {
            //QueryParams[i].
        //}


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
        if (param_count > 0) {    // Список параметров для вывода в Excel
            data_parameters = CreateVariantArray(visible_param_count, 1);
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

        // Вывод данных на лист Excel
        Variant range_title = msexcel.WriteToCell(Worksheet1, param_excel.title_label , 1, 1);
        Variant range_createtime = msexcel.WriteToCell(Worksheet1, "По состоянию на: " + DateTime.DateTimeString(), 2, 1);
        Variant range_parameters;
        if (param_count > 0)
        {
            range_parameters = msexcel.WriteTable(Worksheet1, data_parameters, 3, 1);
        }

        Variant range_tablehead = msexcel.WriteTable(Worksheet1, data_head, 3 + visible_param_count, 1);
        Variant range_tablebody = msexcel.WriteTable(Worksheet1, data_body, 4 + visible_param_count, 1, &df_body);

        msexcel.SetRangeFormat(range_tablehead, cf_head);
        msexcel.SetRangeFormat(range_tablebody, cf_body);
        msexcel.SetRangeFormat(range_title, cf_title);
        msexcel.SetRangeFormat(range_createtime, cf_createtime);
        if (param_count > 0)
        {
            msexcel.SetRangeFormat(range_parameters, cf_createtime);
        }


        Variant range_all = msexcel.GetRangeFromRange(range_tablehead, 1, 1, msexcel.GetRangeRowsCount(range_tablebody)+1, msexcel.GetRangeColumnsCount(range_tablebody));


        if (this->param_excel.title_height > 0)
        {
            msexcel.SetRowHeight(range_tablehead, this->param_excel.title_height);    // Задаем высоту заголовка таблицы
        }

        msexcel.SetAutoFilter(range_all);   // Включаем автофильтр
        msexcel.SetColumnsAutofit(range_all);  // Ширина ячеек по содержимому


        // Настройка заголовка (размер и тп)
        for (int i=0; i < ExcelFieldCount; i++)   // Заполнение если есть поля в ExcelFields
        {    //CELLFORMAT cf_cell;
            //cf_cell.bSetFontColor = true;
            //cf_cell.FontColor = clGreen;

            if (param_excel.Fields[i].bwraptext >= 0)
            {
                CELLFORMAT cf_cell;
                cf_cell.bWrapText = param_excel.Fields[i].bwraptext;
                msexcel.SetRangeFormat(range_tablehead, cf_cell, 1, i+1);
            }

            if (param_excel.Fields[i].width >= 0)
            {
                msexcel.SetColumnWidth(range_tablehead, i+1, param_excel.Fields[i].width);    // Задаем высоту заголовка таблицы
                //msexcel.SetColumnWidth(range_tablehead, 1);    // Задаем высоту заголовка таблицы
            }
        }

        //msexcel.SetRowsAutofit(range_tablehead);


        // Выводим текст sql-запроса на второй лист
        Variant Worksheet2 = msexcel.GetSheet(Workbook, 2);


        Variant range_sqltext;
        int PartMaxLength = 4000;  // 8 192  - максимальная длина строки в ячейке EXCEL
        int n = ceil( (float) _mainQueryText.Length() / PartMaxLength);
        for (int i = 1; i <= n; i++)
        {
            AnsiString sQueryPart = _mainQueryText.SubString(((i-1) * PartMaxLength) + 1, PartMaxLength);
            range_sqltext = msexcel.WriteToCell(Worksheet2, sQueryPart, i, 1);
            msexcel.SetRangeFormat(range_sqltext, cf_sql);
        }

        df_body.clear();

        if (DstFileName == "")
        {
            msexcel.SetVisible(Workbook);
        }
        else
        {
            msexcel.SaveDocument(Workbook, DstFileName);
            VarClear(Workbook);
            VarClear(Worksheet1);
            VarClear(Worksheet2);
            msexcel.CloseApplication();
            _resultFiles.push_back(DstFileName);
        }


        //if (ExportMode == EM_EXCEL_FILE) {
        //    msexcel.SaveAsDocument(Workbook, DstFileName);
        //    msexcel.CloseExcel();
        //} else {
            //Workbook.OlePropertySet("Name", "blabla");
        //    msexcel.SetVisibleExcel(true, true);
        //}
    }

    // Освобождение памяти
    VarClear(data_head);
    data_head = NULL;

    VarClear(data_body);
    data_body = NULL;

    CoUninitialize();

    if (fDone)
    {
        throw Exception("Прерывание.");
    }

}

//---------------------------------------------------------------------------
// Заполнение Excel файла с использованием шаблона xlt
void __fastcall ThreadSelect::ExportToExcelTemplate(TOraQuery *QueryTable, TOraQuery *QueryFields)
{
    CoInitialize(NULL);

    String TemplateFullName = AppPath + param_excel.template_name; // Абсолютный путь к файлу-шаблону

    // Открываем шаблон MS Excel
    MSExcelWorks msexcel;
    Variant Workbook;
    Variant Worksheet;

    try
    {
        msexcel.OpenApplication();
        Workbook = msexcel.OpenDocument(TemplateFullName);
        Worksheet = msexcel.GetSheet(Workbook, 1);
    }
    catch (Exception &e)
    {
        try
        {
            msexcel.CloseApplication();
        }
        catch (...)
        {
        }
        CoUninitialize();
        _threadMessage = "Ошибка при открытии файла-шаблона " + TemplateFullName + ".\nОбратитесь к системному администратору.";
        throw Exception(_threadMessage);
    }

    // Сначала делаем замену полей
    try
    {
        if (QueryFields != NULL)
        {
            msexcel.ExportToExcelFields(QueryFields, Worksheet);
        }
    }
    catch (Exception &e)
    {
        msexcel.CloseApplication();
        CoUninitialize();
        _threadMessage = e.Message;
        throw Exception(e);
    }

    // Затем вставляем табличную часть
    try
    {
        if (QueryTable != NULL && param_excel.table_range_name != "") // Должно быть задано имя диапазона таблично части
        {
            msexcel.ExportToExcelTable(QueryTable, Worksheet, param_excel.table_range_name, param_excel.fUnbounded);
        }
    }
    catch (Exception &e)
    {
        try {
            msexcel.CloseApplication();
        }
        catch (...)
        {
        }
        CoUninitialize();
        _threadMessage = e.Message;
        throw e;
    }

    if (DstFileName == "")         // Просто открываем документ, если имя файла-результата не задано
    {
        msexcel.SetVisible(Workbook);
    }
    else
    {                        // иначе сохраняем в файл
        try
        {
            msexcel.SaveDocument(Workbook, DstFileName);
            msexcel.CloseApplication();
            _resultFiles.push_back(DstFileName);
        }
        catch (Exception &e)
        {
            try
            {
                msexcel.CloseApplication();
            }
            catch (...)
            {
            }
            _threadMessage = "Ошибка при сохранении результата в файл " + DstFileName + ".\n" + e.Message;
            throw Exception(_threadMessage);
        }
    }

    // В дальнейшем сделать аналогично выгрузке в MS Word
    // обьединение двух таблиц QueryFields и QueryTable

    CoUninitialize();
}

//---------------------------------------------------------------------------
// Заполнение DBF-файла
// Переделать эту функцию с использование компонента TDbf
void __fastcall ThreadSelect::ExportToDBF(TOraQuery *OraQuery)
{
    /*TStringList* ListFields;
    int n = this->param_dbase.Fields.size();
    if (n > 0)    // Формируем список полей для экспорта в DBF ("Имя;Тип;Длина;Длина дробной части")
    {
        ListFields = new TStringList();
        for (int i = 0; i < n; i++)
        {
            ListFields->Add(param_dbase.Fields[i].name + ";" + param_dbase.Fields[i].type + ";"+ param_dbase.Fields[i].length + ";" + param_dbase.Fields[i].decimals);
        }
    }
    else
    {
        _threadMessage = "Не задан список полей в параметрах экспорта."
            "\nПожалуйста, обратитесь к системному администратору.";
        throw Exception(_threadMessage);
    }*/

    // Это условие убрано, в связи с тем, что некоторые поля могут оставаться пустыми
    if (param_dbase.Fields.size() > OraQuery->FieldCount && !param_dbase.fAllowUnassignedFields)
    {
        _threadMessage = "Количество требуемых полей превышает количество полей в источнике данных."
            "\nПожалуйста, обратитесь к системному администратору.";
        throw Exception(_threadMessage);
    }

    if (param_dbase.Fields.size() == 0) {
        _threadMessage = "Не задан список полей в параметрах экспорта."
            "\nПожалуйста, обратитесь к системному администратору.";
        throw Exception(_threadMessage);
    }

    // Создаем dbf-файл назначения
    TDbf* pTable = new TDbf(NULL);

    //pTableDst->TableLevel = 7; // required for AutoInc field
    pTable->TableLevel = 4;
    pTable->LanguageID = DbfLangId_RUS_866;

    pTable->TableName = ExtractFileName(DstFileName);
    pTable->FilePathFull = ExtractFilePath(DstFileName);


    // Создаем определение полей таблицы из параметров
    TDbfFieldDefs* TempFieldDefs = new TDbfFieldDefs(NULL);

    if (TempFieldDefs == NULL) {
        _threadMessage = "Can't create storage.";
        throw Exception(_threadMessage);
    }

    for(std::vector<DBASEFIELD>::iterator it = param_dbase.Fields.begin(); it < param_dbase.Fields.end(); it++ )
    {
        TDbfFieldDef* TempFieldDef = TempFieldDefs->AddFieldDef();
        TempFieldDef->FieldName = it->name;
        //TempFieldDef->Required = true;
        //TempFieldDef->FieldType = Field->type;    // Use FieldType if Field->Type is TFieldType else use NativeFieldType
        TempFieldDef->NativeFieldType = it->type[1];
        TempFieldDef->Size = it->length;
        TempFieldDef->Precision = it->decimals;
    }

    if (TempFieldDefs->Count == 0)
    {
        delete pTable;
        _threadMessage = "Не удалось загрузить описание полей.";
        throw Exception(_threadMessage);
    }

    pTable->CreateTableEx(TempFieldDefs);
    pTable->Exclusive = true;
    try
    {
        pTable->Open();
    }
    catch (Exception &e)
    {
        _threadMessage = e.Message;
    }

    // Запись данных в таблицу
    try
    {
	    while ( !OraQuery->Eof )
        {
            pTable->Append();
		    for (int j = 1; j <= OraQuery->FieldCount; j++ )
            {
                pTable->Fields->FieldByNumber(j)->Value = OraQuery->Fields->FieldByNumber(j)->Value;
            }
            OraQuery->Next();  // Переходим к следующей записи
	    }
        pTable->Post();
        pTable->Close();

        _resultFiles.push_back(DstFileName);

    }
    catch(Exception &e)
    {
        pTable->Close();

        delete TempFieldDefs;
        delete pTable;

        _threadMessage = e.Message;
        throw Exception(e);
    }

    delete TempFieldDefs;
    delete pTable;
}

