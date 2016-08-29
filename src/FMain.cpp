//---------------------------------------------------------------------------

#include <vcl.h>
#include <vector>
#include <vcl\Clipbrd.hpp>
#pragma hdrstop

#include "FMain.h"

using namespace std;

//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "Ora"
#pragma link "DBAccess"
#pragma link "OdacVcl"
#pragma link "DBAccess"
#pragma link "DBAccess"
#pragma link "Halcn6DB"
#pragma link "NumEdit"
#pragma resource "*.dfm"
TForm1 *Form1;
const String TASKNAME = "SQL2EXCEL";

//---------------------------------------------------------------------------
//
__fastcall TForm1::TForm1(TComponent* Owner)
	: TForm(Owner)
{
    // Список цветов вкладок
    m_vTabColor.push_back(RGB(180,255,20));     // green
    m_vTabColor.push_back(RGB(120,230,90));     // green
    m_vTabColor.push_back(RGB(0,190,90));       // green
    m_vTabColor.push_back(RGB(0,190,210));      // blue
    m_vTabColor.push_back(RGB(90,225,255));     // blue
    m_vTabColor.push_back(RGB(100,176,255));    // blue
    m_vTabColor.push_back(RGB(200,145,255));    // violet
    m_vTabColor.push_back(RGB(255,100,220));    // violet
    m_vTabColor.push_back(RGB(255,130,170));    // red light
    m_vTabColor.push_back(RGB(255,100,0));      // red
    m_vTabColor.push_back(RGB(255,180,50));     // orange
    m_vTabColor.push_back(RGB(255,255,0));      // yellow


    // Список функций
    m_env_func.reserve(4);
    m_env_func.push_back("_date(");     // Функция date(,,,,)
    m_env_func.push_back("_sql(");      // Функция sql(Text, DBIndex)
    m_env_func.push_back("_compare(");  // Функция compare(Text, Text)
    m_env_func.push_back("_in(");       // Функция in(Text, set)

    // Список "опасных" слов
    DangerWords.reserve(4);
    DangerWords.push_back("execute");
    DangerWords.push_back("truncate");
    //DangerWords.push_back("commit");
    DangerWords.push_back("drop");
    //DangerWords.push_back("insert");
    //DangerWords.push_back("update");
    //DangerWords.push_back("delete");

    OdacLog = new TOdacUtilLog();

    //threadopt = new THREADOPTIONS;

    AppPath = ExtractFilePath(Application->ExeName);
}

//---------------------------------------------------------------------------
// Возвращает цвет по индексу
TColor __fastcall TForm1::ColorByIndex(int index)
{
    int ColorIndex = index % m_vTabColor.size();
    return m_vTabColor[ColorIndex];
}

//---------------------------------------------------------------------------
//
__fastcall TForm1::~TForm1()
{
    delete OdacLog;
}

//---------------------------------------------------------------------------
//
void __fastcall TForm1::FormCreate(TObject *Sender)
{
    m_sessions.push_back(EsaleSession);
    m_sessions.push_back(CCBSession);
    m_sessions.push_back(DBYYSession);
    m_sessions.push_back(DBWORK2Session);

    if (Auth() && PrepareForm()) {
        OdacLog->Init(EsaleSession, "", Username, TASKNAME, AppFullVersion);
        OdacLog->WriteLog("Start application");    // Запись в Лог-таблицу
        //FormResize(NULL);     // Если this->WindowState = wsMaximized
    } else {
        Application->ShowMainForm = false;
        Application->Terminate();
    }
}

//---------------------------------------------------------------------------
// Подготовка элементов управления
bool __fastcall TForm1::PrepareForm()
{
    int result = LoadQueryList();

    bAdmin = Username.UpperCase() == "ADMIN";
    //miExecute->Visible = bAdmin;

    switch (result) {
    case -2:
		MessageBoxStop("Отсутствуют доступные для текущего пользователя запросы. Программа будет закрыта!");
        //this->Free();
        return false;
    case -1:
	    MessageBoxStop("Не удалось открыть таблицу справочника запросов. Программа будет закрыта!");
        //this->Free();
        return false;
    default:
        PrepareTabs();
        FillFieldsLB();
        FillParametersLV();

        ListBox1->ItemIndex = 0;    // Выделяем самую первую строку в списке запросов
        StatusBar1->Panels->Add();
        StatusBar1->Panels->Items[0]->Text = "Готов";

        TabControl1->DoubleBuffered = true;
        Form1->Caption = "Программа для отчетов " + AppFullVersion + " - " + Username;

        return true;
    }
}

//---------------------------------------------------------------------------
// Авторизация пользователя в программе
bool __fastcall TForm1::Auth()
{
    TOraSession* OraSessionAuth = new TOraSession(NULL);
    OraSessionAuth->AssignConnect(EsaleSession);

    LoginForm = new TLoginForm(Application);
    bool loggedon = LoginForm->Execute(OraSessionAuth);
    LoginForm->Free();

    if (loggedon) {
        Username =  UpperCase(Trim(OraSessionAuth->Username));
        OraSessionAuth->Disconnect();
        delete OraSessionAuth;
        return true;
    } else {
        return false;
    }
}

//---------------------------------------------------------------------------
// Загружает вектор
int __fastcall TForm1::LoadQueryList()
{
    // Выбор запросов в соответствии с ролью пользователя и его именем
    AnsiString Str = "select * from ("
        " select * from ("
        " SELECT spr_task_sql2excel.*, nvl(SYS.DBA_ROLE_PRIVS.GRANTED_ROLE, null) GRANTED_ROLE, row_number() over (partition by SPR_TASK_SQL2EXCEL_ID order by queryname) N FROM spr_task_sql2excel"
        " LEFT join SYS.DBA_ROLE_PRIVS on GRANTEE = '" + Username + "'"
        " and upper(userlist) like '%ROLE=\"' || SYS.DBA_ROLE_PRIVS.GRANTED_ROLE || '\"%'"
        " ) where N=1"
        " )"
        " where fvisible=1 and (upper(userlist) like '%USER=\"" + Username + "\"%' or GRANTED_ROLE is not null)"
        " order by taborder,tabname,sortorder,queryname";

    //AnsiString Str = "SELECT * FROM spr_task_sql2excel where fvisible=1 and upper(userlist) like '%USER=\"" + Username + "\"%' order by taborder,tabname,sortorder,queryname";

    TOraQuery *OraQuery_SprTask = OpenOraQuery(EsaleSession, Str);

    if (OraQuery_SprTask == NULL) {
        delete OraQuery_SprTask;
        return -1;
    }


    int RecCount = OraQuery_SprTask->RecordCount;
    if (RecCount <= 0) {
        delete OraQuery_SprTask;
        return -2;
    }

    DataSetToQueryList(OraQuery_SprTask, QueryList, TabList);

    // Очистка памяти
    OraQuery_SprTask->Close();
    delete OraQuery_SprTask;
	OraQuery_SprTask = NULL;

    return RecCount;
}

//---------------------------------------------------------------------------
// Загружает список запросов в вектор
int __fastcall TForm1::DataSetToQueryList(TOraQuery* oraquery, std::vector<TQueryItem>& query_list, std::vector<TTabItem>& tab_list)
{
    InitEnvVariables();  // Инициализация переменных среды

    int RecCount = oraquery->RecordCount;
    if (RecCount <= 0) {
        return NULL;
    }

    query_list.reserve(RecCount);


	// Формируем массив строк параметров и запросов
   	// Формируем список запросов
    oraquery->First();		// Переходим к первой записи (на всякий случай)
    int i = 0;
    int tabindex = 0;
    AnsiString PrevTabName = "";

   	for( ; !oraquery->Eof; oraquery->Next()) {
        // В дальнейшем переделать на вектор указателей, чтобы не дублировать данные при добавлени в вектор
        TQueryItem query;
        query.fExcelFile = false;  // Флаг Excel в файл
        query.fWordFile = false;   // Флаг Word в файл
        query.fDbfFile = false;    // Флаг Dbf в файл

        query.param_excel.title_height = -1;
        query.taborder  = oraquery->FieldByName("taborder")->AsString;
        query.queryid   = oraquery->FieldByName("SPR_TASK_SQL2EXCEL_ID")->AsString;
        query.querytext = oraquery->FieldByName("sqltext1")->AsString + oraquery->FieldByName("sqltext2")->AsString
            + oraquery->FieldByName("sqltext3")->AsString + oraquery->FieldByName("sqltext4")->AsString;
        query.querytext2 = oraquery->FieldByName("sqltext2_1")->AsString;
        query.queryname = oraquery->FieldByName("queryname")->AsString;
        query.dbname    = oraquery->FieldByName("dbname")->AsString;
        query.dbname2   = oraquery->FieldByName("dbname2")->AsString;
        query.sortorder = oraquery->FieldByName("sortorder")->AsString;
        query.spr_task_sql2excel_id = oraquery->FieldByName("spr_task_sql2excel_id")->AsString;
        query.tabname    = oraquery->FieldByName("tabname")->AsString;

        try {
            ParseUserParamsStr(oraquery->FieldByName("userparams")->AsString, &query);
        } catch (...) {
        }
        ParseExportParamsStr(oraquery->FieldByName("exportparams")->AsString, &query);

        if (oraquery->FieldByName("fieldslist")->IsNull) {      // Строка - комментарий к запросу (перечень выводимых полей)
            if (query.param_excel.Fields.size() > 0) {                 // Если не заполнена строка, то брать из параметра выгрузки в Excel
                query.fieldslist = "";
                vector<EXCELFIELD>::iterator cur;
                for (cur = query.param_excel.Fields.begin(); cur < query.param_excel.Fields.end() - 1; cur++)
                    query.fieldslist += cur->name + " | ";
                query.fieldslist += cur->name;
            }
        } else {
            query.fieldslist = oraquery->FieldByName("fieldslist")->AsString;
        }


        query_list.push_back(query);

        tabindex = tab_list.size()-1;
        if (tabindex == -1 || tab_list[tabindex].name != query.tabname) {
            TTabItem Tab;
            Tab.name = query.tabname;
            tab_list.push_back(Tab);
            tabindex++;
        }
        tab_list[tabindex].queryitem.push_back(&query_list[query_list.size()-1]);

        i++;
    }

    return RecCount;
}

//---------------------------------------------------------------------------
// Возращает значение из SQL-Запроса
String TForm1::GetValueFromSQL(String SQLText, String dbindex)
{
    if (SQLText.Trim() == "")
        return "";

    String result = "";
    TOraQuery *OraQuery = NULL;

    try {

        TOraSession *orasession = m_sessions[StrToInt(dbindex)];
        orasession->Connected = true;

        TOraQuery *OraQuery = new TOraQuery(NULL);
        OraQuery->Session = orasession;
        OraQuery->SQL->Add(SQLText);
        OraQuery->Open();

        result = OraQuery->FieldByName("value")->AsString;

        //result = OraQuery->FieldByIndex(0)->AsString;

        OraQuery->Close();
        delete OraQuery;
    } catch(...) {
        if (OraQuery != NULL)
            delete OraQuery;
    }

    return result;
}

//---------------------------------------------------------------------------
// Разбор xml-текста параметров
void TForm1::ParseUserParamsStr(AnsiString ParamStr, TQueryItem* queryitem)
{
    if (ParamStr == "")
        return;

    // Формирование списка параметров
    MsxmlWorks msxml;

   	// Формируем список параметров
	AnsiString xmlParams;
    std::vector<TParamRecord>* ListParams = &queryitem->UserParams;

    msxml.LoadXMLText(ParamStr);

    if (msxml.GetParseError() != "")
        return;

    Variant RootNode = msxml.GetRootNode();
    Variant node = msxml.GetFirstNode(RootNode);

    while (!node.IsEmpty())
    {
        TParamRecord param;
        param.type = LowerCase(msxml.GetAttributeValue(node, "type"));
        param.name = msxml.GetAttributeValue(node, "name");
        param.label = msxml.GetAttributeValue(node, "label");
        param.format = msxml.GetAttributeValue(node, "format");

        param.src = msxml.GetAttributeValue(node, "src");
        param.dbindex = msxml.GetAttributeValue(node, "dbindex");
        param.visible = Trim(LowerCase(msxml.GetAttributeValue(node, "visible")));
        param.visibleif = Trim(LowerCase(msxml.GetAttributeValue(node, "visibleif")));

        //param.Control = (TObject*) Edit1;


        if (param.name == "uchastok" )
        {
            param.value = param.value;
        }


        //TReplaceFlags replaceflags = TReplaceFlags() << rfReplaceAll << rfIgnoreCase;
        //param.value_src = StringReplace(param.value_src, "--", "_", replaceflags); // Для совместимости
        //param.visibleif= StringReplace(param.visibleif, "--", "_", replaceflags); // Для совместимости

        // Тест!!!!!!!
        param.parent = msxml.GetAttributeValue(node, "parent");
        param.value_src = msxml.GetAttributeValue(node, "value");
        param.value_src = ReplaceVariables(&m_env_var, param.value_src);
        param.value = ReplaceVariables(&queryitem->Variables, param.value_src);
        param.value = GetDefinedValue(param.value);     // Доработать здесь!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

        // visibleif
        if (param.visible == "" && param.visibleif != "") {  // visible имеет приоритет над visibleif
            String condition = ReplaceVariables(&m_env_var, param.visibleif);  // Подстановка предопределенных значений в среде
            condition = ReplaceVariables(&queryitem->Variables, condition);  // Подстановка значений, определенных в QUERYITEM

            if (GetDefinedValue(condition) == "true")
                param.visibleflg = true;
            else
                param.visibleflg = false;
            //param.visibleflg = CheckCondition(condition);
            //param.visibleflg = CheckCondition(param.visibleif);
        } else {
            if (param.visible == "false")   // visible имеет приоритет над visibleif
                param.visibleflg = false;
            else
                param.visibleflg = true;
        }

        // deleteif
        if (!msxml.GetAttribute(node, "deleteif").IsEmpty()) {// Если в xml отсутствует параметр value
            param.deleteifflg = true;
            param.deleteifvalue = msxml.GetAttributeValue(node, "deleteif").UpperCase();
        } else {
            param.deleteifflg = false;
        }

        if (param.type == "variable" )
        {
            param.visibleflg = false;
            if (param.name.Length()>0)
                queryitem->Variables.push_back(ENVITEM("_" + param.name, param.value));
            //queryitem->Variables.push_back(ENVITEM(param.name, param.value));

            //node = msxml.GetNextNode(node);
            //continue;
            //param.display = param.value;
            //param.value = GetValueFromSQL(param.value_src, param.dbindex);
        }
        else if (param.type == "string" )
        {
            param.display = param.value;
            param.mask = msxml.GetAttributeValue(node, "mask");
        }
        if (param.type == "integer" )
        {
            param.display = param.value;
        }
        if (param.type == "float" )
        {
            param.display = param.value;
        }
        else if (param.type == "date")
        {
            if (param.value == "")
                param.value = DateToStr(Now());

            param.display = param.value;

            try {
                // Преобразовываем дату в строку в нужном формате
                if (param.format != "") {
                    AnsiString oldShortDateFormat = ShortDateFormat;
                    AnsiString oldDateSeparator = DateSeparator;
                    ShortDateFormat = "dd.MM.yyyy";
                    DateSeparator = '.';
                    param.value = FormatDateTime(param.format, StrToDate(param.value));
                    AnsiString ShortDateFormat = oldShortDateFormat;
                    AnsiString DateSeparator = oldDateSeparator;
                }
            } catch (...){
            }
        }
        else if (param.type == "list")       // Если тип параметра list
        {
            Variant subnode = msxml.GetFirstNode(node);

            // Если в опциях list-a отсутствует параметр value
            // то value будет равно первому значению из списка
            bool bParamValueExist = !msxml.GetAttribute(node, "value").IsEmpty();


            /*AnsiString sqltext = "";
            try {
                sqltext = msxml.GetAttributeValue(subnode, "src");
            } catch (...) {}  */

            if (param.src != "") {    // Если задан sql-запрос

                int dbindex = 0;
                //AnsiString sdbindex = "";
                //Variant DbIndexAttribute = msxml.GetAttribute(subnode, "dbindex");
                //sdbindex = msxml.GetAttributeValue(subnode, "dbindex");
                if (param.dbindex != "") {
                    try {
                        dbindex = StrToInt(param.dbindex);
                    } catch (...) {}
                } else {
                    try {
                        dbindex = StrToInt(queryitem->dbname);  //------------------------------------------------------------------------
                    } catch (...) {}
                    param.dbindex = IntToStr(dbindex);
                }

                try {
                    //TOraSession *orasession = getSessionByIndex(dbindex);
                    TOraSession *orasession = m_sessions[dbindex];
                    orasession->Connected = true;

                    TOraQuery *OraQuery = new TOraQuery(NULL);
                    OraQuery->Session = orasession;
                    OraQuery->SQL->Add(param.src);
                    OraQuery->Open();

                    TParamlistItem item;

                    while (!OraQuery->Eof) {
                        item.value = OraQuery->FieldByName("value")->AsString;
                        item.label = OraQuery->FieldByName("label")->AsString;
                        param.listitem.push_back(item);
                        OraQuery->Next();
                    }

                    OraQuery->Close();
                    delete OraQuery;
                } catch (...) {}
            } else {            // Если задан список значений

                // Если в списке значений компонента list отсутствует параметр value
                // то проставляем значение value для каждого item-a
                bool bValueAutoInc = msxml.GetAttribute(subnode, "value").IsEmpty();
                int i = 0;
                while (!subnode.IsEmpty()) {
                    TParamlistItem item;
                    item.value = bValueAutoInc? IntToStr(i++) : msxml.GetAttributeValue(subnode, "value");
                    item.label = msxml.GetAttributeValue(subnode, "label");
                    item.visible = Trim(LowerCase(msxml.GetAttributeValue(subnode, "visible")));
                    item.visibleif = Trim(LowerCase(msxml.GetAttributeValue(subnode, "visibleif")));

                    if (item.visible == "" && item.visibleif != "") {  // visible имеет приоритет над visibleif
                        String condition = ReplaceVariables(&m_env_var, item.visibleif);  // Подстановка предопределенных значений в среде
                        condition = ReplaceVariables(&queryitem->Variables, condition);  // Подстановка значений, определенных в QUERYITEM


                        if (GetDefinedValue(condition) == "true")
                            item.visibleflg = true;
                        else
                            item.visibleflg = false;


                        //item.visibleflg = CheckCondition(condition);


                        //item.visibleflg = CheckCondition(item.visibleif);
                        //if (record->visibleif != "" && CheckCondition(record->visibleif) != true) {
                        //AnsiString s = "s";
                    } else {
                        if (item.visible == "false")   // visible имеет приоритет над visibleif
                            item.visibleflg = false;
                        else
                            item.visibleflg = true;
                    }


                    param.listitem.push_back(item);
                    subnode = msxml.GetNextNode(subnode);
                }

                // Задаем значение по умолчанию равным первому из списка
                // но только если параметр param.value не задан явно
                if (!bParamValueExist && param.value == "" && param.listitem.size() > 0) {
                    param.value = param.listitem[0].value;
                }
            }

            // Задаем отображение param.display в соответствии с param.value (выбираем из списка)
            for (int j = 0; j < param.listitem.size(); j++)
            {
                if (param.value == param.listitem[j].value)
                {
                    param.display = param.listitem[j].label;
                    break;
                }
            }

            /*if (param.listitem.size() > ival) {
                param.value = param.listitem[ival].value;
                param.display = param.listitem[ival].label;
            }*/
        }
        ListParams->push_back(param);
        node = msxml.GetNextNode(node);
    }
}

//---------------------------------------------------------------------------
// Разбор xml-текста параметров для экспорта
void TForm1::ParseExportParamsStr(AnsiString ParseStr, TQueryItem* queryitem)
{
    if (ParseStr == "") {
        queryitem->DefaultExportType = EM_EXCEL_BLANK; // Если параметры отсутствуют, то по умолчанию выполнять запрос как SELECT
        return;
    }


    //EXPORTMODE FirstExportMode = EM_UNDEFINITE;
    AnsiString FirstId = "";
    queryitem->DefaultExportType = EM_UNDEFINITE;

    try {
        String attribute;
        MsxmlWorks msxml;
        msxml.LoadXMLText(ParseStr);

        if (msxml.GetParseError() != "")
            return;

        Variant RootNode = msxml.GetRootNode();
        Variant node = msxml.GetFirstNode(RootNode);


        // Ко всем id параметров, заданных пользователем подставляется префикс
        queryitem->exportparam_id = "m_" + msxml.GetAttributeValue(RootNode, "default");

        int unassigned_id = 0;
        while (!node.IsEmpty())
        {
            if (LowerCase(msxml.GetAttributeValue(node, "enable")) == "false") {
                node = msxml.GetNextNode(node);
                continue;
            }


            // Если не задан параметр выгрузки по умолчанию,
            // то используется первый параметр, в порядке загрузке (именованый или с id = "0")
            if (queryitem->exportparam_id == "m_") {
                queryitem->exportparam_id = msxml.GetAttributeValue(node, "id", "0");
            }


            String sNodeName = msxml.GetNodeName(node);


            if (sNodeName == "excel") {  // exportparams - excel
                if (queryitem->fExcelFile)    // Загружаем только первый параметр этого типа
                    break;

                queryitem->param_excel.id = "m_" + msxml.GetAttributeValue(node, "id");
                if (queryitem->param_excel.id == "m_") {
                    queryitem->param_excel.id = IntToStr(unassigned_id++);
                }

                queryitem->param_excel.title_label = msxml.GetAttributeValue(node, "title", queryitem->queryname);
                queryitem->param_excel.title_height = msxml.GetAttributeValue(node, "title-height", -1); // Высота заголовка в строках
                queryitem->param_excel.template_name = msxml.GetAttributeValue(node, "template", "");
                queryitem->param_excel.fUnbounded = msxml.GetAttributeValue(node, "unbounded", false);
                queryitem->param_excel.table_range_name = msxml.GetAttributeValue(node, "table_range", "");

                std::vector<EXCELFIELD>* ListFields = &queryitem->param_excel.Fields;

                Variant subnode = msxml.GetFirstNode(node);
                while (!subnode.IsEmpty())
                {
                    if (msxml.GetNodeName(subnode) == "field") {

                        EXCELFIELD field;
                        field.format = LowerCase(msxml.GetAttributeValue(subnode, "format"));
                        field.name = msxml.GetAttributeValue(subnode, "name");
                        field.width = msxml.GetAttributeValue(subnode, "width", -1);    // Ширина столбца
                        attribute = LowerCase(Trim(msxml.GetAttributeValue(subnode, "wraptext")));  // перенос по словам
                        if (attribute == "false")
                            field.bwraptext = 0;
                        else if (attribute == "true")
                            field.bwraptext = 1;
                        else
                            field.bwraptext = -1;

                        ListFields->push_back(field);
                    }

                    subnode = msxml.GetNextNode(subnode);
                }

                if (queryitem->param_excel.id == queryitem->exportparam_id) {
                    if (queryitem->param_excel.template_name == "") {   // В шаблон если указан шаблон
                        queryitem->DefaultExportType = EM_EXCEL_BLANK;
                    } else {
                        queryitem->DefaultExportType = EM_EXCEL_TEMPLATE;
                    }
                }

            } else if (sNodeName == "dbase4")
            {
                if (queryitem->fDbfFile)    // Загружаем только первый параметр этого типа
                    break;
                queryitem->fDbfFile = true;
                queryitem->param_dbase.fAllowUnassignedFields = msxml.GetAttributeValue(node, "allowunassigned", false);

                queryitem->param_dbase.id = "m_" + msxml.GetAttributeValue(node, "id");
                if (queryitem->param_dbase.id == "m_") {
                    queryitem->param_dbase.id = IntToStr(unassigned_id++);
                }

                // Парсим список полей dbase4
                std::vector<DBASEFIELD>* ListFields = &queryitem->param_dbase.Fields;
                Variant subnode = msxml.GetFirstNode(node);
                while (!subnode.IsEmpty())
                {
                    if (msxml.GetNodeName(subnode) == "field") {
                        DBASEFIELD field;
                        field.type = LowerCase(msxml.GetAttributeValue(subnode, "type"));
                        field.name = LowerCase(msxml.GetAttributeValue(subnode, "name"));
                        field.length = msxml.GetAttributeValue(subnode, "length", 0);
                        field.decimals = msxml.GetAttributeValue(subnode, "decimals", 0);
                        ListFields->push_back(field);
                    }
                    subnode = msxml.GetNextNode(subnode);
                }

                if (queryitem->param_dbase.id == queryitem->exportparam_id) {
                    queryitem->DefaultExportType = EM_DBASE4_FILE;
                }

            } else if (sNodeName == "word") {
               if (queryitem->fWordFile)   // Загружаем только первый параметр этого типа
                    break;
                queryitem->fWordFile = true;

                queryitem->param_word.id = "m_" + msxml.GetAttributeValue(node, "id");
                if (queryitem->param_word.id == "m_") {
                    queryitem->param_word.id = IntToStr(unassigned_id++);
                }

                queryitem->param_word.template_name = msxml.GetAttributeValue(node, "template");
                queryitem->param_word.filter_main_field = msxml.GetAttributeValue(node, "filter_main_field", "");
                queryitem->param_word.filter_sec_field = msxml.GetAttributeValue(node, "filter_sec_field", "");
                queryitem->param_word.filter_infix_sec_field = msxml.GetAttributeValue(node, "filter_infix_sec_field", "");
                queryitem->param_word.page_per_doc = msxml.GetAttributeValue(node, "page_per_doc", 0);

                if (queryitem->param_word.id == queryitem->exportparam_id) {
                    queryitem->DefaultExportType = EM_WORD_TEMPLATE;
                }
            } if (sNodeName == "execute") {
                queryitem->param_execute.id = "m_" + msxml.GetAttributeValue(node, "id");
                if (queryitem->param_execute.id == "m_") {
                    queryitem->param_execute.id = IntToStr(unassigned_id++);
                }

                if (queryitem->param_execute.id == queryitem->exportparam_id) {
                    queryitem->DefaultExportType = EM_PROCEDURE;
                }
            }

            // Вместо этого кода реализовать хранение параметров в векторе
            /*if (FirstId == "") {
                AnsiString sId = "m_" + msxml.GetAttributeValue(node, "id")
                if (sId == "m_")
                    sId = "0";
                FirstId = sId;
            }*/

            node = msxml.GetNextNode(node);
        }

        if (queryitem->DefaultExportType == EM_UNDEFINITE) {
            if (queryitem->param_excel.template_name == "") {   // В шаблон если указан шаблон
                queryitem->DefaultExportType = EM_EXCEL_BLANK;
            } else {
                queryitem->DefaultExportType = EM_EXCEL_TEMPLATE;
            }
        }


    } catch (...) {
    }
}

//---------------------------------------------------------------------------
//
void __fastcall TForm1::Run(EXPORTMODE ExportMode, int Tag)
{
    // Проверяем не заблокированы ли запросы к БД
    if (CheckLock(StrToInt(CurrentQueryItem->dbname))) {
        return;
    }

    // Проверяем не заблокированы ли запросы к БД
    if (CurrentQueryItem->dbname2 !="" && CheckLock(StrToInt(CurrentQueryItem->dbname2))) {
        return;
    }

    THREADOPTIONS* threadopt = new THREADOPTIONS;
    switch (ExportMode) {
        case EM_PROCEDURE: {
        // Предупреждаем, что может произойти необратимое изменение данных
            String msg = "Внимание! Выполнение данного запроса может привести к необратимому изменению данных. Продолжить?";
            if (MessageBoxQuestion(msg) != IDNO) {
                threadopt->exportmode = ExportMode;
                DoExport(threadopt);
            }
        }
        break;
        case EM_EXCEL_BLANK:
        {
            if (Tag == 0) {
                threadopt->exportmode = ExportMode;
                DoExport(threadopt);
            } else {
                // Опции окна сохранения результов
                SaveDialog1->Options.Clear();
                SaveDialog1->Options << ofFileMustExist;
                SaveDialog1->Filter = "MS Excel файлы (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";
                SaveDialog1->FilterIndex = 1;
                SaveDialog1->DefaultExt = "xlsx";

                AnsiString filename;
                if (SaveDialog1->Execute()) {
                    threadopt->dstfilename = SaveDialog1->FileName;
                    threadopt->exportmode = ExportMode;
                    DoExport(threadopt);
                }
            }
        }
        break;
        case EM_EXCEL_TEMPLATE: {
            if (Tag == 0) {
                threadopt->exportmode = ExportMode;
                DoExport(threadopt);
            } else {
                // Опции окна сохранения результов
                SaveDialog1->Options.Clear();
                SaveDialog1->Options << ofFileMustExist;
                SaveDialog1->Filter = "MS Excel файлы (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";
                SaveDialog1->FilterIndex = 1;
                SaveDialog1->DefaultExt = "xlsx";

                AnsiString filename;
                if (SaveDialog1->Execute()) {
                    threadopt->dstfilename = SaveDialog1->FileName;
                    threadopt->exportmode = ExportMode;
                    DoExport(threadopt);
                }
            }
        }
        break;
        case EM_WORD_TEMPLATE: {
            // Проверяем на сущестовование файла-шаблона
            String TemplateFullName = AppPath + CurrentQueryItem->param_word.template_name;
            if(!FileExists(TemplateFullName)) {
                MessageBoxStop("Файл шаблона " + TemplateFullName + " не найден.");
                return;
            }

            // Опции окна сохранения результов
            SaveDialog1->Options.Clear();
            SaveDialog1->Options << ofFileMustExist;
            SaveDialog1->Filter = "MS Word файлы (*.doc)|*.doc|Все файлы (*.*)|*.*";
            SaveDialog1->FilterIndex = 1;
            SaveDialog1->DefaultExt = "";

            if (SaveDialog1->Execute()) {
                threadopt->dstfilename = ChangeFileExt(SaveDialog1->FileName, "");
                threadopt->exportmode = ExportMode;
                DoExport(threadopt);
             }
        }
        break;
        case EM_DBASE4_FILE: {

            SaveDialog1->Options.Clear();
            SaveDialog1->Options << ofFileMustExist;
            SaveDialog1->Filter = "DBase4 файлы (*.dbf)|*.dbf|Все файлы (*.*)|*.*";
            SaveDialog1->FilterIndex = 1;
            SaveDialog1->DefaultExt = "dbf";

            AnsiString filename;
            if (SaveDialog1->Execute()) {
                //filename = SaveDialog1->FileName;
                //DoExport(ThreadSelect::TO_DBASE4_FILE, filename);
                threadopt->dstfilename = SaveDialog1->FileName;
                threadopt->exportmode = ExportMode;
                DoExport(threadopt);
            }
        }
        break;
    }

    delete threadopt;
}


//---------------------------------------------------------------------------
// Выполнение запроса, установленного "По умолчанию"
void __fastcall TForm1::ActionDefaultRunExecute(TObject *Sender)
{
    Run(CurrentQueryItem->DefaultExportType);
}

//---------------------------------------------------------------------------
//
void __fastcall TForm1::ActionAsProcedureExecute(TObject *Sender)
{
    Run(CurrentQueryItem->DefaultExportType);
}


//---------------------------------------------------------------------------
// Экспорт в файл Excel
void __fastcall TForm1::ActionExportExcelFileExecute(TObject *Sender)
{
    if (CurrentQueryItem->param_excel.template_name == "")
        Run(EM_EXCEL_BLANK, 1);
    else
        Run(EM_EXCEL_TEMPLATE, 1);

}

//---------------------------------------------------------------------------
// Экспорт в Excel (в память)
void __fastcall TForm1::ActionExportExcelBlankExecute(TObject *Sender)
{
    if (CurrentQueryItem->param_excel.template_name == "")
        Run(EM_EXCEL_BLANK, 0);
    else
        Run(EM_EXCEL_TEMPLATE, 0);
}

//---------------------------------------------------------------------------
// Экспорт в MS Word
void __fastcall TForm1::ActionExportWordFileExecute(TObject *Sender)
{
    Run(EM_WORD_TEMPLATE);
}

//---------------------------------------------------------------------------
// Экспорт в файл DBASE4
void __fastcall TForm1::ActionExportDbfFileExecute(TObject *Sender)
{
    Run(EM_DBASE4_FILE);
}


//---------------------------------------------------------------------------
//
void __fastcall TForm1::FormClose(TObject *Sender, TCloseAction &Action)
{
    QueryList.clear();
    TabList.clear();
    DangerWords.clear();
}

//---------------------------------------------------------------------------
//
String TForm1::GetSQL(String SQLText)
{
    DinamicControlExit(NULL);
    //ShowMessage(Parameters[SelIndex][0].value.c_str());		// Удаляем побелы

    // Функция получения текста запроса
    // Заменяет параметры в строке на значения
    // Удаляет /** и **/
    // Собирает строку запроса
    // Обрабатываем строку вручную

    TParamRecord *params;

    int nDangerWords = DangerWords.size();

    // Защита от Иньекций  (обработка параметров)
    for (unsigned int j = 0; j < CurrentQueryItem->UserParams.size(); j ++) {   // Убираем из парамтров все лишнее
        /*int found = Parameters[SelIndex][j].value.find("--") +
        Parameters[SelIndex][j].value.find("/*") +
        Parameters[SelIndex][j].value.find("select ")+
        Parameters[SelIndex][j].value.find("'")+
        Parameters[SelIndex][j].value.find(";");
        if (found > -5) {
            Parameters[SelIndex][j].value = "";
            break;
        }*/
        for (int i = 0; i < nDangerWords; i++)
        {
            //int k = Parameters[SelIndex][j].value.Pos(DangerWords[i]);
            //AnsiString s = CurrentQueryItem->Parameters[j].value;
            if (CurrentQueryItem->UserParams[j].value.Pos(DangerWords[i]) != 0) {
                CurrentQueryItem->UserParams[j].value = "";
                break;
            }
        }
    }

    // Готовим запрос (заменяем параметры на значения)
	std::vector<EXPLODESTRING> sqlstring;
    sqlstring = ExplodeByBackslash(SQLText, "/**", "**/", true);

	for (unsigned int i = 0; i < sqlstring.size(); i++) {  // Цикл по параметрам в исходной строке
        EXPLODESTRING *item;
        item = &sqlstring[i];
    	if (item->fBacksleshed) { 			// Замена (подстановка) параметров в строке запроса если строка является  /** параметр **/

            TReplaceFlags replaceflags = TReplaceFlags() << rfReplaceAll << rfIgnoreCase;
            item->text = StringReplace(item->text, "/**", "", replaceflags);
 			item->text = StringReplace(item->text, "**/", "", replaceflags);     // Удаляем **/

            //item->text = item->text.LowerCase();  // 2016-07-06

            //bool bParamFined = false;
            for (unsigned int j = 0; j < CurrentQueryItem->UserParams.size(); j++) {   //Заменяем --Параметр на Значение
                TParamRecord *param;
                param = &CurrentQueryItem->UserParams[j];

                param->name = param->name.Trim();		// Удаляем побелы
                param->value = param->value.Trim();		// Удаляем побелы
                if (param->name != "" && item->text.Pos(":"+param->name)>0) {
                    //bParamFined = true;
                    if (param->deleteifflg == true && param->value.UpperCase() == param->deleteifvalue.UpperCase())
                        item->text = "";
                    else
 			            item->text = StringReplace(item->text, ":"+param->name, param->value, replaceflags);     // Удаляем **/
                }
                /*else
                    item->text = "";   */
            }
              /*if (!bParamFined)   // Удаляем строку с параметром, если отсутствуют подходящие параметры в списке
                item->text = "" + item->text + " ERROR! ONE OF THESE PARAMETERS NOT FOUND!";   */
        }
    }

    AnsiString result = Implode(sqlstring, "");

    // Защита от Иньекций  (обработка запроса в целом)
    for (int i = 0; i < nDangerWords; i++)
        if (result.Pos(DangerWords[i]) != 0) {
            result = "";
            break;
        }

    return result;  // Собираем вектор в строку и возращаем результат
}

//---------------------------------------------------------------------------
// В последующем вставить эту функцию в taskutil.h
String __fastcall TForm1::GetValue(String value)
{
    if (value.Length() < 2 || value[1] != '_' )
        return value;

    //String f_date = '_date('
    //vector<String>::iterator cur;
    //for (cur = m_env_func.begin(); cur <m_env_func.end() - 1; cur++) {


    String Result;
    int n = m_env_func.size();
    for (int i = 0; i < n; i++) {
        if (value.Pos(m_env_func[i]) != 1)
            continue;

        // Получаем строку с параметрами
        std::vector<EXPLODESTRING2> sqlstring;
        sqlstring = ExplodeByBackslash2(value, m_env_func[i], ")");
        std::vector<AnsiString> params;

        // Разбиваем строку с параметрами с разделителем - (,)
        if (sqlstring[1].fBacksleshed) {
            params = Explode(sqlstring[1].text, ",", false);
        }

        int n_params = params.size();
        switch (i) {
            // Функция _date(v1, v2, p1, p2, format)
            // Возвращает дату, рассчитываемую с учетом задаваемой опциями точкой отсчета
            case 0:
            {
                TDateTime ResultDate = Date();

                 // Обработка параметров
                if ( n_params == 5) {
                    String param_day = params[0];   // Кол-во дней
                    String param_month = params[1]; // Кол-во месяцев
                    String param_option_day = params[2];    // Точка отсчета дней
                    String param_option_month = params[3];  // Точка отсчета месяцев
                    String param_format = params[4];
                    //break;

                    // Вычисляем дату
                    // Сначала определим точку отсчета (день и месяц), если заданы специальные опции
                    // Текущий месяц (0), Первый месяц (1), последний месяц (2)
                    if (param_option_month == "1" || param_option_month == "first") {
                        ResultDate = EncodeDate(YearOf(ResultDate), 1, DayOf(ResultDate));
                    } else if (param_option_month == "2" || param_option_month == "last") {
                        ResultDate = EncodeDate(YearOf(ResultDate), 12, DayOf(ResultDate));
                    }

                    // Текущее число (0), Первый день месяца (1), последний день месяца (2)
                    if (param_option_day == "1" || param_option_day == "first") {
                        ResultDate = EncodeDate(YearOf(ResultDate), MonthOf(ResultDate), 1);
                    } else if (param_option_day == "2" || param_option_day == "last") {
                        ResultDate = EncodeDate(YearOf(ResultDate), MonthOf(ResultDate), DaysInAMonth(ResultDate));
                    }

                    // Прибавляем дни и месяцы
                    ResultDate = IncMonth(ResultDate, StrToInt(param_month));
                    ResultDate = ResultDate + StrToInt(param_day);

                    String format = ExplodeByBackslash2(param_format, "'", "'", false)[0].text;  // извлекаем формат из кавычек
                    DateTimeToString(Result, format, ResultDate);
                }

                break;
            }
            case 1: // m_env_func[i] = "_sql("
            {
                Result = GetValueFromSQL(params[0], params[1]);
                break;
            }

            // Функция _compare(val1, val2)
            // Сравненивает два значения на равенство
            case 2:
            {
                if (n_params != 2)
                    Result = "error";//value;
                else
                    Result = params[0] == params[1]? "true" : "false";

                break;

               /*
                 if (condition.Trim() == "")
                    return false;
               vector<AnsiString> t;
                t = Explode(condition, "=", false);
                if (t.size() == 1)


                if (t.size() == 1) {
                t[0] = t[0].LowerCase();
                if (t[0]=="true")
                    return true;
                else
                    return false;
                }
                else if (t.size() != 2) {
                return false;
                } else
                return t[0] == t[1];  */


            }
            // Функция _in(val1, {v1,v2,v3,...})
            // Проверяет вхождения значения во множество
            case 3:
            {
                if (n_params != 2)
                    Result = "error";
                else {
                    Result = "false";
                    String value = params[0];

                    String tmp = ExplodeByBackslash2(params[1], "{", "}", false)[0].text;
                    std::vector<AnsiString> vset;
                    vset = Explode(tmp, ",", false);

                    int n_vsetsize = vset.size();
                    if (n_vsetsize > 0) {
                        for (int j = 0; j < n_vsetsize; j++) {
                            if (value == vset[j])
                                Result = "true";
                        }
                    } else {        // Ошибка
                        Result = "error";
                    }
                }
                break;
            }
        }

    }
    return Result;
}


//---------------------------------------------------------------------------
// Оставлена для совместимости со старым форматом даты.
// В последующем заменить одной функцией GetValue // !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
AnsiString TForm1::GetDefinedValue(AnsiString value)
{
    if (value.Length() < 2 || value[1] != '_')
        return value;

    int aaa = value.Pos("_date(");
    int bbb = value.Pos("_sql(");
    int ccc = value.Pos("_compare(");
    int ddd = value.Pos("_in(");
    if (aaa || bbb || ccc || ddd)
        return GetValue(value);
}

//---------------------------------------------------------------------------
// Инициализация переменных среды
// ВНИМАНИЕ! Эта функция зависит от наличия некоторых таблиц в базе данных
// заданных значением EsaleSession.
void TForm1::InitEnvVariables()
{
    m_env_var.reserve(4);

    // Филиал
    AnsiString filial;      // Филиал  (id_rn)
    AnsiString maingroup;   // Участок
    if (Username.UpperCase() == "ADMIN") {
        filial = "01";
        maingroup = "01";
    } else if (Username.Length() == 7) {
        // Определение кода филиала
        // и кода участка
        // Здесь желательно переделать полностью
        // Возможно следует использовать Роли в БД Oracle
        TOraQuery *OraQuery = new TOraQuery(NULL);
        OraQuery->Session = EsaleSession;
        //OraQuery->SQL->Add("select * from raion where substr(:username,3,2) = substr(uuser,3,2)");
        OraQuery->SQL->Add("select  trim(to_char(nvl2(uch, '01', id),'00')) id, trim(to_char(nvl(uch, '01'),'00')) uch from raion"
            " left join nasel_uch on nasel_uch.id_rn = raion.id where substr(:username,3,2) = substr(uuser,3,2)"
            //" and substr(:username,2,2) <> '10'"
            " order by nasel_uch.porydok");


        OraQuery->ParamByName("username")->AsString = Username;
        OraQuery->Open();

        if (OraQuery->RecordCount == 0) {
            filial = "01";
            maingroup = "01";
        } else {
            filial = OraQuery->FieldByName("id")->AsString;
            maingroup = OraQuery->FieldByName("uch")->AsString;
        }
        OraQuery->Close();
        delete OraQuery;

        //filial = Username.SubString(3, 2);
    }

    // Формируем список полей в одну строку _roles
    String roles = "{";
    TOraQuery *OraQuery = new TOraQuery(NULL);
    OraQuery->Session = EsaleSession;
    OraQuery->SQL->Add("select * from session_roles");
    OraQuery->Open();
    while (!OraQuery->Eof) {
        roles += "'"+OraQuery->FieldByName("role")->AsString+"'";
        OraQuery->Next();
        if (!OraQuery->Eof)
            roles += ",";
    }
    roles +="}";
    OraQuery->Close();
    delete OraQuery;

    m_env_var.push_back(ENVITEM("_filial",filial));
    m_env_var.push_back(ENVITEM("_maingroup",maingroup));
    m_env_var.push_back(ENVITEM("_username", Username.LowerCase()));
    m_env_var.push_back(ENVITEM("_roles", roles.LowerCase()));
}

//---------------------------------------------------------------------------
// Обработчик событий в потоке
void __fastcall TForm1::OnThread(int Status, AnsiString Message)
{
    switch (Status) {
        case WM_THREAD_PROCEED_BEGIN_SQL:
        {
           this->Enabled = false;
           Application->CreateForm(__classid(TForm_Wait), &Form_Wait);
           Form_Wait->Label3->Caption = "Выполнение запроса...";
           TotalTime = 0;
           Timer1->Enabled = true;
           Form_Wait->Show();
           break;
        }
        case WM_THREAD_PROCEED_BEGIN_FETCH:
        {
           Form_Wait->Label3->Caption = "Извлечение данных...";
           break;
        }
        case WM_THREAD_PROCEED_BEGIN_DOCUMENT:
        {
            Form_Wait->Label3->Caption = "Создание документа...";
            break;
        }
        case WM_THREAD_PROCEED_EXCEL:
            break;

        case WM_THREAD_PROCEED_DONE:
        {
            this->Enabled = true;
            Timer1->Enabled = false;
            OdacLog->WriteLog("Report is prepared");    // Запись в Лог-таблицу
            Form_Wait->Release();
            ts = NULL;
            break;
            // Далее выполняется OnThreadSuccess
        }
        case WM_THREAD_USER_CANCEL:      // Отмена пользователем
        {
            this->Enabled = true;
            Timer1->Enabled = false;
            Form_Wait->Release();
            break;
        }
        case WM_THREAD_ERROR_BD_CANT_CONNECT:
        {
            this->Enabled = true;
            Timer1->Enabled = false;
            Form_Wait->Release();
            AnsiString msg = "Не удалось подключиться к базе данных. \nПопробуйте выполнить запрос позднее.";
            MessageBoxInf(msg);
            break;
        }
        case WM_THREAD_ERROR_NULL_RESULTS:
        {
            this->Enabled = true;
            Timer1->Enabled = false;
            Form_Wait->Release();
            AnsiString msg = "С учетом заданных параметров получено 0 строк.\nПопробуйте изменить параметры отбора.";
            MessageBoxInf(msg);
            break;
        }
        case WM_THREAD_ERROR_TOO_MORE_RESULTS:
        {
            this->Enabled = true;
            Timer1->Enabled = false;
            Form_Wait->Release();
            AnsiString msg = "С учетом заданных параметров получено более 1 млн. строк.\nПопробуйте уточнить параметры отбора.";
            MessageBoxInf(msg);
            break;
        }
        case WM_THREAD_ERROR_PARAMS_INCORRECT:
        {
            this->Enabled = true;
            Timer1->Enabled = false;
            Form_Wait->Release();
            AnsiString msg = "Не удалось выполнить запрос.\nПроверьте корректность параметров отбора.";
            MessageBoxStop(msg);
            break;
        }
        case WM_THREAD_ERROR_IN_PROCESS:
        {
            this->Enabled = true;
            Timer1->Enabled = false;
            Form_Wait->Release();
            AnsiString msg = "Возникла ошибка в процессе обработки данных.\n" + Message;
            MessageBoxStop(msg);
            break;
        }
        case WM_THREAD_ERROR_IN_PROCESS_ALT:
        {
            this->Enabled = true;
            Timer1->Enabled = false;
            Form_Wait->Release();
            AnsiString msg = Message;
            MessageBoxStop(msg);
            break;
        }
        case WM_THREAD_ERROR_OPEN_QUERY:
        {
            this->Enabled = true;
            Timer1->Enabled = false;
            Form_Wait->Release();
            AnsiString msg = "Возникла ошибка при открытии основного запроса.\n" + Message;
            MessageBoxStop(msg);
            break;
        }
        case WM_THREAD_ERROR_OPEN_QUERY2:
        {
            this->Enabled = true;
            Timer1->Enabled = false;
            Form_Wait->Release();
            AnsiString msg = "Возникла ошибка при открытии вспомогательного запроса.\n" + Message;
            MessageBoxStop(msg);
            break;
        }
        case WM_THREAD_EXECUTE_DONE:
        {
            this->Enabled = true;
            Timer1->Enabled = false;
            Form_Wait->Release();
            AnsiString msg = "Выполнено успешно.\n";
            MessageBoxInf(msg);
            break;

        }
        case WM_THREAD_EXECUTE_ERROR:
        {
            this->Enabled = true;
            Timer1->Enabled = false;
            Form_Wait->Release();
            AnsiString msg = "Возникла непредвиденная ошибка в процессе выполнения запроса.\n";
            MessageBoxStop(msg);
            break;
        }
    }
}

//---------------------------------------------------------------------------
// Обработчик завершения потока
void __fastcall TForm1::OnThreadSuccess(EXPORTMODE ExportMode, std::vector<String> vResultFiles)
{
    switch (ExportMode) {
    case EM_EXCEL_TEMPLATE:
    case EM_EXCEL_BLANK:
        if (vResultFiles.size() > 0) {
            MessageBoxInf("Результат сохранен в файле " + vResultFiles[0]);
            try {
                ExploreFile(this->Handle, vResultFiles[0]);
            } catch (...) {
            }
        }
        break;
    case EM_DBASE4_FILE:
        //AnsiString filepath = ExtractFilePath(threadopt->filename);
        MessageBoxInf("Результат сохранен в файле " + vResultFiles[0]);
        try {
            ExploreFile(this->Handle, vResultFiles[0]);
        } catch (...) {
        }
        break;
    case EM_WORD_TEMPLATE:
        String s = "";
        int MaxOut = 5;         // ПРОВЕРИТЬ!!!!!!!!!!!!!!!!!!!!
        int n = vResultFiles.size();
        int nOut = n > MaxOut ? MaxOut : n;
        for (int i = 0; i < nOut; i++) {    // Выводим имена только первых MaxOut файлов
            s += "\n" + vResultFiles[i] ;
        }
        if (n > MaxOut)
            s += "\n...";       // Если файлов > MaxOut
            
        AnsiString filepath = ExtractFilePath(vResultFiles[0]);
        MessageBoxInf("Результат сохранен в каталоге " + filepath +
            "\nФайлы (" + IntToStr(n) + " шт.):" + s);
        try {
            ExploreFile(this->Handle, vResultFiles[0]);
        } catch (...) {
        }
        break;
     }
}

//---------------------------------------------------------------------------
//
void __fastcall TForm1::ListBox1DrawItem(TWinControl *Control, int Index,
      TRect &Rect, TOwnerDrawState State)
{
    TColor colorText1;  // Цвет текста первой строки
    TColor colorText1Sel;  // Цвет текста первой строки
    TColor colorText2;  // Цвет текста второй строки
    TColor colorText2Sel;  // Цвет текста второй строки
    TColor colorBkOdd;  // Цвет фона Нечетного элемента
    TColor colorBkEven; // Цвет фона Четного элемента

    // Определение цвета
    colorText1 = RGB(0,0,0);
    colorText1Sel = RGB(255,255,255);
    colorText2 = RGB(80,80,80);
    colorText2Sel = RGB(255,255,255);
    colorBkOdd = RGB(240,240,240);
    colorBkEven = RGB(255,255,255);

    TListBox *pListBox = static_cast <TListBox *> (Control);
    TCanvas *pCanvas = pListBox->Canvas;


    std::string text1;      // Верхняя строка
    std::string text2;      // Нижняя строка
    std::string str = pListBox->Items->Strings[Index].c_str();

    //Разделение строки по символу перевода строки \n
    //AnsiString str = pListBox->Items->Strings[Index];
    // Разбиваем строки
    std::string separator = "\\n";
    int k = str.find(separator);
    if (k > 0) {
        text1 = str.substr(0, k);
        text1 = text1 + "\0";

        int le = str.length()-k;
        text2 = str.substr(k+2, le);
    } else {
        text1 = str;
        text2 = "";
    }

    // РИСОВАНИЕ
    pCanvas->Lock();

    if (Index % 2 == 1) {       //Красим не чётные строки
        pCanvas->Brush->Color = colorBkOdd;
        pCanvas->FillRect(Rect);
    } else {
        pCanvas->Brush->Color = colorBkEven;
        pCanvas->FillRect(Rect);    // Очищаем область (старый фон)
    }

    // if the item is selected...
    if (State.Contains(odSelected)) {
        pCanvas->Font->Color = clHighlightText;
        pCanvas->Brush->Color = clHighlight;
        pCanvas->FillRect(Rect);
    }

    // Вывод текста
    const int offset = 3;       // set this to offset the text

    if (State.Contains(odSelected))     // Цвет для первой строки
        pCanvas->Font->Color = colorText1Sel;    // Цвет шрифта
    else
        pCanvas->Font->Color = colorText1;    // Цвет шрифта


    pCanvas->TextWidth(text1.c_str())+2+2;
    unsigned int th = pCanvas->TextHeight(text1.c_str());

    if (text1 != "") {
        //pCanvas->Font->Size=12;
        pCanvas->Font->Style = pCanvas->Font->Style << fsBold;// << fsUnderline;
        pCanvas->TextOut(
            Rect.Left + offset, Rect.Top+3,
            text1.c_str() );
    }
    if (text2 != "") {
        if (State.Contains(odSelected))     // Цвет для второй строки
            pCanvas->Font->Color = colorText2Sel;    // Цвет шрифта
        else
            pCanvas->Font->Color = colorText2;    // Цвет шрифта

        //pCanvas->Font->Color = RGB(80,80,80);    // Цвет шрифта
        pCanvas->Font->Style = pCanvas->Font->Style >> fsBold;// >> fsUnderline;
        //pCanvas->Font->Size=pCanvas->Font->Size + 12;
        pCanvas->Font->Height=pCanvas->Font->Height - 6;
        //pCanvas->Font->Style = fsNormal;
        pCanvas->TextOut(
            Rect.Left + offset, Rect.Top + th + 3,
            text2.c_str() );
    }

    if (State.Contains(odFocused)) {    // Убираем рамку фокуса
        // remove the focus rect (i.e., XOR it away)
        DrawFocusRect(pCanvas->Handle, &Rect);
    }

    pCanvas->Unlock();
}


//---------------------------------------------------------------------------
// Выводит список параметров в ListView
void TForm1::FillParametersLV()
{
    //MessageBoxStop("Параметры содержат недопустимые значения!\nИзмените значения параметров!");

    if (ListBox1->ItemIndex < 0)
        ListBox1->ItemIndex = 0;

    TQueryItem* qi = TabList[TabControl1->TabIndex].queryitem[ListBox1->ItemIndex];
    if (CurrentQueryItem == qi)
        return;

    // Важное место!!! Установка выбранного запроса
    CurrentQueryItem = qi;

    ParamsLV->Items->BeginUpdate();
    ParamsLV->Items->Clear();
	for (unsigned int i = 0; i < CurrentQueryItem->UserParams.size(); i++) {
        TParamRecord *record = &CurrentQueryItem->UserParams[i];

        if (!record->visibleflg)
            continue;

        /*if (record->visibleif != "" && CheckCondition(record->visibleif) != true) {
            continue;
        }       */

        TListItem *Item = ParamsLV->Items->Add();
        Item->Caption = record->label.c_str();
        Item->SubItems->Add(record->display.c_str());
    }
    ParamsLV->Items->EndUpdate();

    // Задаем вид кнопки - выполняющий запрос "По умолчанию"
    // Здесь, возможно, стоит поработать, так как по умолчанию
    // если выгружается в Excel, то ВСЕГДА выгружается в память
    BitBtn1->Glyph = NULL;
    switch(CurrentQueryItem->DefaultExportType) {
    case EM_PROCEDURE:
        ActionDefaultRun->Caption = ActionAsProcedure->Caption;
        BitBtn2->Enabled = false;
        ImageList1->GetBitmap(0, BitBtn1->Glyph);
        break;
    case EM_EXCEL_BLANK:
        ActionDefaultRun->Caption = ActionExportExcelBlank->Caption;
        BitBtn2->Enabled = true;
        ImageList1->GetBitmap(1, BitBtn1->Glyph);
        break;
    case EM_EXCEL_TEMPLATE:
        ActionDefaultRun->Caption = ActionExportExcelBlank->Caption;
        BitBtn2->Enabled = true;
        ImageList1->GetBitmap(1, BitBtn1->Glyph);
        break;
    case EM_WORD_TEMPLATE:
        ActionDefaultRun->Caption = ActionExportWordFile->Caption;
        BitBtn2->Enabled = true;
        ImageList1->GetBitmap(2, BitBtn1->Glyph);
        break;
    /*case EM_WORD_MEMORY:    // пока не реализовано
        ActionDefaultRun->Caption = ActionExportWordFile->Caption;
        BitBtn2->Enabled = true;
        ImageList1->GetBitmap(2, BitBtn1->Glyph);
        break; */
    case EM_DBASE4_FILE:
        ActionDefaultRun->Caption = ActionExportDbfFile->Caption;
        BitBtn2->Enabled = true;
        ImageList1->GetBitmap(3,BitBtn1->Glyph);
        break;
    }
    BitBtn1->Caption = ActionDefaultRun->Caption;
}

//---------------------------------------------------------------------------
// Вычисление логического выражения  - по умолчанию возвращает false
bool TForm1::CheckCondition(AnsiString condition)
{
    if (condition.Trim() == "")
        return false;

    vector<AnsiString> t;
    t = Explode(condition, "=", false);
    if (t.size() == 1)


    if (t.size() == 1) {
        t[0] = t[0].LowerCase();
        if (t[0]=="true")
            return true;
        else
            return false;
    }
    else if (t.size() != 2) {
        return false;
    } else
        return t[0] == t[1];


/*    String lparam = ReplaceVariables(&m_env_var, t[0]);  // Подстановка предопределенных значений в среде
    lparam = ReplaceVariables(&queryitem->Variables, t[0]);  // Подстановка значений, определенных в QUERYITEM

    String rparam = ReplaceEnvVariables(&m_env_var, t[1]);
    rparam = ReplaceEnvVariables(&queryitem->Variables, t[1]); */

//    return lparam == rparam;
}

//---------------------------------------------------------------------------
// Подстановка переменных, определенных в m_env_var (глобальных)
String TForm1::ReplaceEnvVariables(AnsiString condition)
{
    TReplaceFlags replaceflags = TReplaceFlags() << rfReplaceAll << rfIgnoreCase;

    for (std::vector<ENVITEM>::iterator it = m_env_var.begin() ; it != m_env_var.end(); ++it) {
        condition = StringReplace(condition, it->name, it->value, replaceflags);     // Удаляем **/
    }

    return condition;
}

//---------------------------------------------------------------------------
// Подстановка переменных, определенных в QUERYITEM (локальных)
//AnsiString TForm1::ReplaceQueryVariables(std::vector<TParamRecord>* ListParams)
String TForm1::ReplaceVariables(std::vector<ENVITEM>* Variables, const String& Text)
{
    if (Variables->size() < 1 || Text.Length() < 1)
        return Text;
        
    TReplaceFlags replaceflags = TReplaceFlags() << rfReplaceAll << rfIgnoreCase;
    String Result = Text;
    for (std::vector<ENVITEM>::iterator it = (*Variables).begin() ; it != (*Variables).end(); ++it) {
        Result = StringReplace(Result, it->name, it->value, replaceflags);
    }
    return Result;
}

//---------------------------------------------------------------------------
// Рисование вкладок
void __fastcall TForm1::PageControl1DrawTab(TCustomTabControl *Control,
      int TabIndex, const TRect &Rect, bool Active)
{
    TTabControl *pTabControl = static_cast <TTabControl *> (Control);
    TCanvas *pCanvas = Control->Canvas;

    // Текст, который будем выводить на вкладке
    AnsiString TabCaption = TabControl1->Tabs->Strings[TabIndex];


    // ОПРЕДЕЛЯЕМ ЦВЕТА
    TColor colorBk;     // Цвет фона
    TColor colorText = colorText = RGB(0,0,0);     // Цвет текста

    // Цикл из 8 цветов (в т.ч. светло серый)
    //int m_ColorIndex = TabIndex % m_vTabColor.size();  // Повторять цвета по кругу

    //colorBk = m_vTabColor[m_ColorIndex];
    colorBk = ColorByIndex(TabIndex);

    pCanvas->Brush->Color = colorBk;
    pCanvas->Font->Color = colorText;    // Цвет шрифта


    // СОЗДАЕМ ПОВЕРНУТЫЙ НА 90 ГРАДУСОВ ШРИФТ
    HFONT hfontTimes;         // Font handle
    LOGFONT logfont;          // Logical font structure

    // First, clear all fields.
    memset (&logfont, 0, sizeof (logfont));

    // Создаем вертикальный шрифт
    logfont.lfHeight = pTabControl->Font->Height;   //-13;
    logfont.lfWidth = 0;
    logfont.lfEscapement = 900;         // Вращение
    logfont.lfOrientation = 900;        // Вращение  900
    //Active ? logfont.lfWeight = FW_BOLD: logfont.lfWeight = FW_NORMAL;
    logfont.lfWeight = FW_NORMAL;  //FW_BOLD;
    logfont.lfItalic = FALSE;
    logfont.lfUnderline = FALSE;
    logfont.lfStrikeOut = FALSE;
    logfont.lfCharSet = DEFAULT_CHARSET;
    logfont.lfOutPrecision = OUT_DEFAULT_PRECIS;
    logfont.lfClipPrecision = CLIP_DEFAULT_PRECIS;
    logfont.lfQuality = DEFAULT_QUALITY;
    logfont.lfPitchAndFamily = DEFAULT_PITCH | FF_DONTCARE;
    _tcsncpy (logfont.lfFaceName, TEXT("Arial"), LF_FACESIZE);
    //const char* fontname = pTabControl->Font->Name.c_str();
    //_tcsncpy (logfont.lfFaceName,  fontname, LF_FACESIZE);
    logfont.lfFaceName[LF_FACESIZE-1] = TEXT('\0');  // Ensure null termination
    pCanvas->Font->Handle = CreateFontIndirect (&logfont);

    // РИСОВАНИЕ
    pCanvas->Lock();    // Блокирум канвас перед рисованием
    pCanvas->TextRect(Rect, Rect.Left+10, Rect.Bottom-4, TabCaption);
    pCanvas->Unlock();


}
//---------------------------------------------------------------------------
//
void __fastcall TForm1::TabControl1Change(TObject *Sender)
{
    DinamicControlExit(CurrentDinamicControl);

    FillFieldsLB();
    FillParametersLV();
}

//---------------------------------------------------------------------------
// Выводит список запросов в ListBox 
void TForm1::FillFieldsLB()
{
    Panel3->Color = m_vTabColor[TabControl1->TabIndex % m_vTabColor.size()];

    TTabItem* TabItem = &TabList[TabControl1->TabIndex];
    ListBox1->Items->BeginUpdate();
    ListBox1->Clear();
    for (int i = 0; i < TabItem->queryitem.size(); i++) {
        AnsiString sName = TabItem->queryitem[i]->queryname;   // QueryName
        AnsiString sFields = TabItem->queryitem[i]->fieldslist; // Fields
        ListBox1->Items->Add(sName + "\\n" + sFields);
     }
    ListBox1->Items->EndUpdate();
    if (ListBox1->Items->Count > 0)
        ListBox1->ItemIndex = 0;
}

//---------------------------------------------------------------------------
// Создание вкладок
void TForm1::PrepareTabs()
{
    AnsiString stabs;

    if (TabList.size() > 0)
        stabs = TabList[0].name;

    for (int i = 1; i < TabList.size(); i++) {
        stabs = stabs + "\n" + TabList[i].name  ;
    }

    TabControl1->Tabs->BeginUpdate();
    TabControl1->Tabs->SetText(stabs.c_str());
    TabControl1->Tabs->EndUpdate();
}

//---------------------------------------------------------------------------
// Изменяет ширину элемента TabControl при измении размера формы
void __fastcall TForm1::FormResize(TObject *Sender)
{
    TabControl1->Width = TabControl1->RowCount() * TabControl1->TabHeight;  // Изменяем ширину TabControl1 в зависимости от количества строк в нем
}

//---------------------------------------------------------------------------
// Таймер выполнения запроса
void __fastcall TForm1::Timer1Timer(TObject *Sender)
{
    TotalTime += 0.001 * Timer1 -> Interval;
    AnsiString sec = IntToStr((int) TotalTime % 60);
    AnsiString min = IntToStr((int) TotalTime / 60);
    sec = str_pad(sec.c_str(), 2, "0", STR_PAD_LEFT).c_str();
    StatusBar1->Panels->Items[0]->Text =  min + ":" + sec;

    Application->ProcessMessages();
}

//---------------------------------------------------------------------------
// Прячет динамический элемент управления для редактирования параметра
void __fastcall TForm1::DinamicControlExit(TObject *Sender)
{
    if (Sender != NULL)
    {
        TControl *Control = (TControl*)Sender;
        Control->Visible = false;
        TParamRecord *param = &CurrentQueryItem->UserParams[Control->Tag];
        if (param->type == "list")
        {
            TComboBox *ComboBox = (TComboBox*)Sender;
            param->display = ComboBox->Text;
            if (ComboBox->ItemIndex >=0) {

                //for(param->listitem)
                int n = 0;
                int i = 0;  // Индекс элемента в векторе
                for (i = 0; i < param->listitem.size(); i++) { // пропуск скрытых элементов (visibleflg = false)
                    param->listitem[i];
                    if (!param->listitem[i].visibleflg)
                        continue;
                    if (n == ComboBox->ItemIndex) {
                        break;
                    }
                    n++;
                }
                if (param->listitem[i].value != "")
                    param->value = param->listitem[i].value;
                else
                    param->value = IntToStr(ComboBox->ItemIndex);

            }
            /*// Тест - зависимые параметры - свойство parent
            int n = CurrentQueryItem->UserParams.size();
            for (int i = 0; i < n; i++) {
                TParamRecord *p = &CurrentQueryItem->UserParams[i];
                if (p->parent != "") {
                    if (param->name == p->parent) {
                        p->value = param->value;
                        //p->display = p->listitem;
                        //p->display = param->display;
                        UpdateParametersLV();
                    }
                    //p->parent = p->parent;
                }
            }   */
        }
        else if (param->type == "date")
        {
            TDateTimePicker* DateTimePicker = (TDateTimePicker*)Sender;
            param->display = DateToStr(DateTimePicker->DateTime);
            if (param->format == "")
                param->value = param->display;
            else
                param->value = FormatDateTime(param->format, DateTimePicker->DateTime);
        } else if (param->type == "string") {
            if (param->mask == "") {
                TEdit* EditBox = (TEdit*)Sender;
                param->display = EditBox->Text;
            } else {
                TMaskEdit* EditBox = (TMaskEdit*)Sender;
                param->display = EditBox->Text;
             }
             param->value = param->display;
        } else if (param->type == "integer" || param->type == "float" ) {
            //TNumEdit* EditBox = (TNumEdit*)Sender;
            param->display = NumEdit1->Text;
            param->value = param->display;
        }

        ParamsLV->Items->Item[ParamsLV->Tag]->SubItems->Strings[0] = param->display;
    } else {
        DateTimePicker1->Visible = false;
        Edit1->Visible = false;
        MaskEdit1->Visible = false;
        ComboBox1->Visible = false;
        NumEdit1->Visible = false;
    }
}

//---------------------------------------------------------------------------
// Обрабатывает сообщение KeyDown от динамических полей редактирования параметров
void __fastcall TForm1::DinamicControlOnKeyDown(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
    if (Key == VK_ESCAPE) { // Доделать (неправильно обрабатывается нажатие Esc)
        TControl *Control = (TControl*)Sender;
        Control->Visible = false;
    } else if (Key == VK_RETURN) {
        DinamicControlExit(Sender);
    } 
}


/*
void TParamRecordCtrl::Show(TParamRecordCtrl *paramRecord)
{
    ((TWinControl*) paramRecord->Control)->Visible = true;
}

/*
void TParamRecordCtrl:: SetType(String Type)
{
    this.type = Type;

}*/


//---------------------------------------------------------------------------
// Начало редактирования значения параметра
void __fastcall TForm1::OnEditParam()
{
    if (ParamsLV->Selected == NULL)
        return;

    TRect rect = ParamsLV->Items->Item[0]->DisplayRect(drLabel);

    int top = ParamsLV->Selected->Top;
    int left = ParamsLV->Columns->Items[0]->Width+1;
    int width =  ParamsLV->Columns->Items[1]->Width;
    int height = rect.Height();//.Bottom - rect.Top;

    // Определяем индекс параметра с учетом флага visible
    int LV_itemindex = ParamsLV->Selected->Index;  // Индекс элемента в ParamsLV
    ParamsLV->Tag = LV_itemindex;      // Текущий выделенный элемент в ParamsLV

    int n = 0;
    int paramitem_index = 0;  // Индекс элемента в векторе параметров CurrentQueryItem->Parameters
    for (paramitem_index = 0; paramitem_index < CurrentQueryItem->UserParams.size(); paramitem_index++) {
        TParamRecord *param = &CurrentQueryItem->UserParams[paramitem_index];
        if (!param->visibleflg) // пропуск скрытых элементов (visibleflg = false)
            continue;
        if (n == LV_itemindex) {
            break;
        }
        n++;
    }

    TParamRecord *param;
    param = &CurrentQueryItem->UserParams[paramitem_index];

    //TWinControl *Control;

    Edit1->Visible = false;
    ComboBox1->Visible = false;
    DateTimePicker1->Visible = false;


    if (param->type == "date")
    {
        DateTimePicker1->Parent = ParamsLV;       // Устанавливает границы отображения элемента
        DateTime_SetFormat(DateTimePicker1->Handle, "dd.MM.yyyy");
        DateTimePicker1->Width = width;
        DateTimePicker1->Top = top;
        DateTimePicker1->Left = left;
        DateTimePicker1->Height = height-2;
        DateTimePicker1->Font = ParamsLV->Font;
        DateTimePicker1->Font->Size = 10;
        DateTimePicker1->Tag = paramitem_index;  // Текущий выделенный элемент в векторе

        try {
            DateTimePicker1->Date = StrToDate(param->display);
        } catch (...) {
            DateTimePicker1->Date = Now();
        }

        DateTimePicker1->Visible = true;
        DateTimePicker1->SetFocus();
        //CurrentDinamicControl = DateTimePicker1;
    }
    else if (param->type == "list") {
        ComboBox1->Parent = ParamsLV;
        ComboBox1->Width = width;
        ComboBox1->Top=top;
        ComboBox1->Left=left;
        ComboBox1->Height=height-2;
        ComboBox1->Font = ParamsLV->Font;
        ComboBox1->Font->Size = 10;
        ComboBox1->Tag = paramitem_index;  // Текущий выделенный элемент в векторе

        ComboBox1->Clear();
        int cur_item=0; // Текущий элемент. i-не подходит, так как добавляться могут не все элементы
        for (int i=0; i < param->listitem.size();i++)
        {
            TParamlistItem item = param->listitem[i];


            String condition = ReplaceVariables(&m_env_var, item.visibleif);  // Подстановка предопределенных значений в среде
            condition = ReplaceVariables(&CurrentQueryItem->Variables, condition);  // Подстановка значений, определенных в QUERYITEM

            //if (GetDefinedValue(condition) == "true")
            //  item.visibleflg = true;
            //else
            //  item.visibleflg = false;

            if (item.visibleif != "" && GetDefinedValue(condition) != "true") {
                //bool k = CheckCondition(item.visibleif);
                continue;
            }

            ComboBox1->Items->Add(item.label);
            if (item.label == param->display) {
                ComboBox1->ItemIndex = cur_item;
            }
            cur_item++;
        }
        ComboBox1->Text = param->display;
        ComboBox1->Visible = true;
        ComboBox1->SetFocus();

        //CurrentDinamicControl = ComboBox1;
    } else if (param->type == "string" /* || param->type == "integer" || param->type == "float"*/) {
        if (param->mask == "") {
            //TEdit* EditBox = new TEdit(this);
            Edit1->Parent = ParamsLV;
            Edit1->Width = width;
            Edit1->Top=top;
            Edit1->Left=left;
            Edit1->Height=height-2;
            Edit1->Font = ParamsLV->Font;
            Edit1->Font->Size = 10;
            Edit1->Tag = paramitem_index;
            Edit1->Text = param->display;
            Edit1->Visible = true;
            Edit1->SetFocus();
        } else {
            MaskEdit1->EditMask = param->mask;
            MaskEdit1->Parent = ParamsLV;
            MaskEdit1->Width = width;
            MaskEdit1->Top=top;
            MaskEdit1->Left=left;
            MaskEdit1->Height=height-2;
            MaskEdit1->Font = ParamsLV->Font;
            MaskEdit1->Font->Size = 10;
            MaskEdit1->Tag = paramitem_index;
            MaskEdit1->Text = param->display;
            MaskEdit1->Visible = true;
            MaskEdit1->SetFocus();
        }
    } else if (param->type == "integer" || param->type == "float") {
        NumEdit_bUseSign = true;
        NumEdit_bUseDot = param->type == "float";
        NumEdit1->Parent = ParamsLV;
        NumEdit1->Width = width;
        NumEdit1->Top=top;
        NumEdit1->Left=left;
        NumEdit1->Height=height-2;
        NumEdit1->Font = ParamsLV->Font;
        NumEdit1->Font->Size = 10;
        NumEdit1->Tag = paramitem_index;
        NumEdit1->Text = param->display;
        NumEdit1->Visible = true;
        NumEdit1->SetFocus();
    };
}

//---------------------------------------------------------------------------
// Обрабатывает Click в списке параметров
void __fastcall TForm1::ParamsLVClick(TObject *Sender)
{
    OnEditParam();
}

//---------------------------------------------------------------------------
// Прорисовка списка параметров
void __fastcall TForm1::ParamsLVAdvancedCustomDraw(
      TCustomListView *Sender, const TRect &ARect, TCustomDrawStage Stage,
      bool &DefaultDraw)
{
    DefaultDraw = false;

    TColor colorBkSubItems = RGB(255,255,255);
    TColor colorBkItems = clBtnFace;//RGB(245,245,245);
    TColor colorFontItems = RGB(0,0,0);

    TCanvas *pCanvas = ParamsLV->Canvas;

    TRect Rect;

    pCanvas->Lock();

    if (ParamsLV->Items->Count > 0) {
        TListItem *Item = ParamsLV->Items->Item[ParamsLV->Tag];  // Текущий выделенный элемент в ParamsLV
        if (Edit1->Visible) {
            Rect = Item->DisplayRect(drBounds);
            Edit1->Top = Rect.Top;
        } else if (ComboBox1->Visible) {
            Rect = Item->DisplayRect(drBounds);
            ComboBox1->Top = Rect.Top;
        } else if (DateTimePicker1->Visible) {
            Rect = Item->DisplayRect(drBounds);
            DateTimePicker1->Top = Rect.Top;
        }
    } else {
        pCanvas->Brush->Color = clBtnFace;
        pCanvas->FillRect(ARect);
    }

    for(int i = 0; i < ParamsLV->Items->Count; i++)
    {
        TListItem *Item = ParamsLV->Items->Item[i];

        Rect = Item->DisplayRect(drBounds);

        TRect RectItem;
        TRect RectSubItem;

        pCanvas->Font->Size = 10;

        bool bSeparator = CurrentQueryItem->UserParams[Item->Index].type == "separator";

        if (!bSeparator) {
            // Если не разделитель
            RectItem = Item->DisplayRect(drLabel);
            RectItem = TRect(RectItem.Left, RectItem.Top, RectItem.Right, RectItem.Bottom-1);
            RectSubItem = TRect(RectItem.Right+1, Rect.Top, Rect.Right, Rect.Bottom-1);

            pCanvas->Font->Color = colorFontItems;
            pCanvas->Brush->Color = colorBkItems;
            pCanvas->TextRect(RectItem, RectItem.Left+2, RectItem.Top+2, Item->Caption);

            pCanvas->Font->Color = clNavy;
            pCanvas->Brush->Color = colorBkSubItems;
            pCanvas->TextRect(RectSubItem, RectSubItem.Left+2, RectSubItem.Top+2, Item->SubItems->Strings[0]);
        } else {
            // Если разделитель
            RectItem = Item->DisplayRect(drLabel);
            RectItem = TRect(RectItem.Left, RectItem.Top, Rect.Right, Rect.Bottom-1);

            pCanvas->Font->Color = colorFontItems;
            pCanvas->Brush->Color = colorBkItems;

            pCanvas->Font->Style = pCanvas->Font->Style << fsBold;
            pCanvas->TextRect(RectItem, RectItem.Left+2, RectItem.Top+2, Item->Caption);
            pCanvas->Font->Style = pCanvas->Font->Style >> fsBold;
        }




        //pCanvas->MoveTo(Rect.Left, Rect.Bottom);
        //pCanvas->LineTo(Rect.Right, Rect.Bottom);
    }

    pCanvas->Unlock();

}
//---------------------------------------------------------------------------
// F1
void __fastcall TForm1::ActionShowHelpExecute(TObject *Sender)
{
    //int nrezerv = ReservedParams.size();
    AnsiString str;
/*    str = "Список полей для автозамены на значения по умолчанию:\n";
    for (int i = 0; i < nrezerv; i++) { // Заменяем на предопределенные значения
        str += ReservedParams[i][0] + " = " + ReservedParams[i][1] + "\n";
    }     */
    str += "Образец заполнения параметров:\n"
    "<?xml version=\"1.0\"?>\n"
    "<parameters>\n"
    "<parameter type=\"list\" name=\"param_year\" value=\"--curyear-6\" label=\"Год закрытия\">\n"
    "    <item value=\"2014\" label=\"2014\"/>\n"
    "    <item value=\"2015\" label=\"2015\"/>\n"
    "</parameter>\n"
    "<parameter type=\"list\" name=\"param_month\" value=\"--curmonth\" label=\"Месяц закрытия\">\n"
    "    <item value=\"01\" label=\"январь\"/>\n"
    "    <item value=\"02\" label=\"февраль\"/>\n"
    "    <item value=\"03\" label=\"март\"/>\n"
    "    <item value=\"04\" label=\"апрель\"/>\n"
    "    <item value=\"05\" label=\"май\"/>\n"
    "    <item value=\"06\" label=\"июнь\"/>\n"
    "    <item value=\"07\" label=\"июль\"/>\n"
    "    <item value=\"08\" label=\"август\"/>\n"
    "    <item value=\"09\" label=\"сентябрь\"/>\n"
    "    <item value=\"10\" label=\"октябрь\"/>\n"
    "    <item value=\"11\" label=\"ноябрь\"/>\n"
    "    <item value=\"12\" label=\"декабрь\"/>\n"
    "</parameter>\n"
    "<parameter type=\"list\" name=\"raion\" value=\"0\" label=\"Район\">\n"
    "    <item src=\"select id value, fname label from raion\" dbindex=\"0\"/>\n"
    "</parameter>\n"
    "<parameter type=\"date\" name=\"date\" value=\"--firstdaycurmonth\" label=\"Дата\" format=\"dd.mm.yyyy\">\n"
    "</parameter>\n"
    "<parameter type=\"string\" name=\"test\" value=\"строка\" label=\"Наименование\">\n"
    "</parameter>\n"
    "</parameters>\n";

    MessageBoxInf(str);

}
//---------------------------------------------------------------------------
// Ctrl+C
// Скопировать текст запроса
void __fastcall TForm1::ActionCopyQueryExecute(TObject *Sender)
{
    AnsiString str = GetSQL(CurrentQueryItem->querytext);
    Clipboard()->AsText = str;
    //FormShowQuery->ShowQuery(str, CurrentQueryItem->queryname);
}


//---------------------------------------------------------------------------
// Ctrl+S
// Отобразить текст основного запроса
void __fastcall TForm1::ActionShowMainQueryExecute(TObject *Sender)
{
    if (CurrentQueryItem->querytext != "") {
        AnsiString str = GetSQL(CurrentQueryItem->querytext);
        FormShowQuery->ShowQuery(str, "SQL-текст основного запроса \"" + CurrentQueryItem->queryname + "\"");
    } else {
        MessageBoxInf("Текст основного запроса отсутствует.\n");
    }

}

//---------------------------------------------------------------------------
// Ctrl+Alt+S
// Отобразить текст вспомогательного запроса
void __fastcall TForm1::ActionShowSecondaryQueryExecute(TObject *Sender)
{
    if (CurrentQueryItem->querytext2 != "") {
        AnsiString str = GetSQL(CurrentQueryItem->querytext2);
        FormShowQuery->ShowQuery(str, "SQL-текст вспомогательного запроса \"" + CurrentQueryItem->queryname + "\"");
    } else {
        MessageBoxInf("Текст вспомогательного запроса отсутствует.\n");
    }
}


//---------------------------------------------------------------------------
//
void __fastcall TForm1::FormShow(TObject *Sender)
{
    ParamsLV->Height = Panel2->Height - BitBtn1->Height;
    BitBtn1->Top = ParamsLV->Top + ParamsLV->Height;
    BitBtn2->Top = BitBtn1->Top;
    FormResize(NULL);

    ShortDateFormat = "dd.MM.yyyy";
    DateSeparator = '.';


     // Don't show the seconds, Sekunden nicht anzeigen
    //SendMessage(DateTimePicker1->Handle, DTM_SETFORMAT, 0, long(PChar("dd.MM.yyyy")));

}

//---------------------------------------------------------------------------
//
void __fastcall TForm1::EsaleSessionAfterDisconnect(TObject *Sender)
{
    //ShowMessage("Ваше соединение разорвано. Пожалуйста выйдите из программы.");
}

//---------------------------------------------------------------------------
// Проверка блокировки БД
// -1 - заблокированы запросы к любой БД, >= 0 - заблокированы запросы по индексу БД
bool __fastcall TForm1::CheckLock(int dbindex)
{
    CheckLockQuery->ParamByName("dbindex")->AsInteger = dbindex;
    if (CheckLockQuery->Active)
        CheckLockQuery->Refresh();
    else
        CheckLockQuery->Execute();

    if (CheckLockQuery->RecordCount > 0) {
        if (bAdmin) {
            String msg = "Запросы к базе данных " + m_sessions[dbindex]->Name + " заблокированы.\nПродолжить в любом случае?";
            //if (MessageBoxQuestion(msg) != IDNO) {
            return MessageBoxQuestion(msg) == IDNO;
        } else {
            String msg = "Запросы к базе данных " + m_sessions[dbindex]->Name + " заблокированы.\nПопробуйте выполнить данный запрос позднее.";
            MessageBoxInf(msg);

            return true;
        }
    }

    return false;
}

//---------------------------------------------------------------------------
// Нажатие кнопки - PopUp-Меню Экспорт в DBASE4, Excel
void __fastcall TForm1::BitBtn2Click(TObject *Sender)
{

    GetParentForm(BitBtn2)->ActiveControl = NULL;
    TPoint point = BitBtn2->ClientToScreen(TPoint(0,0));
    PopupMenu1->Popup(point.x, point.y);
}

//---------------------------------------------------------------------------

// Задает параметры потока
// и запускает его
void __fastcall TForm1::DoExport(THREADOPTIONS* threadopt)
{
    // Выполнение выбранного запроса
    if (ListBox1->ItemIndex < 0) {          // Если не выбран запрос в списке
        MessageBoxStop("Выберите запрос!");
        return;
    }

    // Выбор необходимого OraSession
    TOraSession *orasession = NULL;
    TOraSession *orasession2 = NULL;
    try {                                       // Если в запросе в БД не указан индекс в поле DBNAME
        int dbname = 0;
        int dbname2 = 0;
        dbname = StrToInt(CurrentQueryItem->dbname);
        orasession = m_sessions[dbname];       // Основная сессия

        if (CurrentQueryItem->dbname2 != "") {
            dbname2 = StrToInt(CurrentQueryItem->dbname2);
            orasession2  = m_sessions[dbname2];    // Дополнительная сессия
        }
    } catch(...) {
        MessageBoxStop("Исходная база данных указана не верно!\nОбратитесь к администратору.");
        return;
    }

    OdacLog->WriteLog("Execute query", CurrentQueryItem->queryname);    // Запись в Лог-таблицу

    // ФОРМИРОВАНИЕ СТРОКИ ЗАПРОСА
    AnsiString querytext;
    AnsiString querytext2;
    querytext = GetSQL(CurrentQueryItem->querytext);    // Основной запрос
    querytext2 = GetSQL(CurrentQueryItem->querytext2);  // Дополнительный запрос, может не использоваться (используется в отчетах MS Word)

    threadopt->querytext = querytext;
    threadopt->querytext2 = querytext2;
    threadopt->queryitem = CurrentQueryItem;
    threadopt->ParentFormHandle = this->Handle;

    threadopt->TemplateOraSession = orasession;
    threadopt->TemplateOraSession2 = orasession2;

    // СОЗДАНИЕ И ЗАПУСК ПОТОКА
    ts = new ThreadSelect(true);    // Создаем приостановленный поток
    ts->SetThreadOpt(threadopt);    // Передаем параметры
    ts->Resume();                   // Запускаем
}

//---------------------------------------------------------------------------
// О программе...
void __fastcall TForm1::ActionAboutAppExecute(TObject *Sender)
{
    //
    String MsgStr = "Программа для подготовки отчетов\nSQL2Excel v." + AppVersion + " (" + AppBuild + ")"
        "\n"
        "\nCopyright © 2014-2016"
        "\nЦентральный филиал ОАО \"Челябэнергосбыт\""
        "\n"
        "\nАвтор:"
        "\nИнженер-программист Отдела ИТ"
        "\nЦентрального филиала ОАО \"Челябэнергосбыт\""
        "\nВ.С. Овчинников"
        "\n"
        "\ne-mail: V.Ovchinnikov@cf.esbt.ru";
    MessageBoxInf(MsgStr, "О программе SQL2Excel");
}

//---------------------------------------------------------------------------
//
void __fastcall TForm1::PopupMenu1Popup(TObject *Sender)
{
    // Скрытие/отображение пунктов меню
    ActionExportDbfFile->Enabled = CurrentQueryItem->fDbfFile;
    ActionExportWordFile->Enabled = CurrentQueryItem->fWordFile;
    //ActionExportExcelFile->Enabled = CurrentQueryItem->fExcelFile;
    //ActionExportExcelMemory->Enabled = CurrentQueryItem->fExcelFile;
}

//---------------------------------------------------------------------------
// Реакция на щелчек мыши
void __fastcall TForm1::ListBox1MouseDown(TObject *Sender,
      TMouseButton Button, TShiftState Shift, int X, int Y)
{
    ListBox1->ItemIndex = ListBox1->ItemAtPos(TPoint(X,Y),false);
    FillParametersLV();
}

//---------------------------------------------------------------------------
// Реакция на щелчек мыши, а также на клавиши вверх, вниз
void __fastcall TForm1::ListBox1Click(TObject *Sender)
{
    FillParametersLV();
}

//---------------------------------------------------------------------------
// Процедура для NumEdit1 - TEdit вместо компонента (TNumEdit)
void __fastcall TForm1::NumEdit1Change(TObject *Sender)
{
   if (IsNumber(NumEdit1->Text, NumEdit_bUseDot, NumEdit_bUseSign)) {
        NumEdit_TextOld = NumEdit1->Text;
        NumEdit_SelStartOld = NumEdit1->SelStart;
    } else {
        TNotifyEvent event = NumEdit1->OnChange;
        NumEdit1->OnChange = NULL;
        try {
            NumEdit1->Text = NumEdit_TextOld;
            NumEdit1->SelStart = NumEdit_SelStartOld;
        }
        __finally {
            NumEdit1->OnChange = event;
        }
    }
}

//---------------------------------------------------------------------------
// Процедура для NumEdit1 - TEdit вместо компонента (TNumEdit)
void __fastcall TForm1::NumEdit1KeyPress(TObject *Sender, char &Key)
{
    NumEdit_SelStartOld = NumEdit1->SelStart;
}

//---------------------------------------------------------------------------
// Выход из программы
void __fastcall TForm1::ActionApplictionExitExecute(TObject *Sender)
{
    //if (MessageBoxQuestion("Вы уверены что хотите выйти из программы?") != IDNO) {
        Close();
    //}
}

//---------------------------------------------------------------------------

