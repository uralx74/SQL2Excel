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
#pragma link "MemDS"
#pragma link "EditAlt"
#pragma resource "*.dfm"
TForm1 *Form1;
const String TASKNAME = "SQL2EXCEL";

const String mainSpr = "spr_task_sql2excel";
const String envSpr = "spr_task_sql2excel_env";

const String SYSTEM_VARIABLES_PREFIX = "$";
const String CUSTOM_VARIABLES_PREFIX = "_";



#include <assert.h>

/*        #ifndef NDEBUG
        ndebag_counter++;
        //OutputDebugString("Now starting dangerous function");
        //assert(1==1);
        #endif
  /*
/* ���������� ��� ��������
   2016-09-20 IS DEPRECATED!
*/
String function_compare(const std::vector<String>& parameters)
{
    if (parameters.size() != 2)
    {
        return "error";
    }
    return parameters[0] == parameters[1]? OleXml::TRUE_STR_VALUE : OleXml::FALSE_STR_VALUE;

}

/* ��������� ������ �� �������� � ������
     parameters[0] - ����������� ��������
     parameters[1 + N] - �������� ������ (N �� 0 �� �������������)
 */
String function_in(const std::vector<String>& parameters)
{
    if (parameters.size() < 2) {
        return "error";
    }

    for (std::vector<String>::const_iterator it = parameters.begin()+1; it != parameters.end(); ++it)
    {
        if (parameters[0] == *it ) {
            return OleXml::TRUE_STR_VALUE;
        }
    }

    return OleXml::FALSE_STR_VALUE;
}

/* ��������� ����
     parameters[0] - ��������� ����
     parameters[1] - ��������� �������
     parameters[2] - ����� ������� ���� (0 - �������, 1 - ������, 2 - ���������)
     parameters[3] - ����� ������� �������
     parameters[4] - ������ ������
 */
String function_date(const std::vector<String>& parameters)
{
    if (parameters.size() != 5)
    {
        return "error";
    }

    TDateTime ResultDate = Date();

    String param_day = parameters[0];   // ���-�� ����
    String param_month = parameters[1]; // ���-�� �������
    String param_option_day = parameters[2];    // ����� ������� ����
    String param_option_month = parameters[3];  // ����� ������� �������
    String param_format = parameters[4];

    // ��������� ����
    // ������� ��������� ����� ������� (���� � �����), ���� ������ ����������� �����
    // ������� ����� (0), ������ ����� (1), ��������� ����� (2)
    if (param_option_month == "1" || param_option_month == "first")
    {
        ResultDate = EncodeDate(YearOf(ResultDate), 1, DayOf(ResultDate));
    }
    else if (param_option_month == "2" || param_option_month == "last")
    {
        ResultDate = EncodeDate(YearOf(ResultDate), 12, DayOf(ResultDate));
    }

    // ������� ����� (0), ������ ���� ������ (1), ��������� ���� ������ (2)
    if (param_option_day == "1" || param_option_day == "first")
    {
        ResultDate = EncodeDate(YearOf(ResultDate), MonthOf(ResultDate), 1);
    }
    else if (param_option_day == "2" || param_option_day == "last")
    {
        ResultDate = EncodeDate(YearOf(ResultDate), MonthOf(ResultDate), DaysInAMonth(ResultDate));
    }

    // ���������� ��� � ������
    ResultDate = IncMonth(ResultDate, StrToInt(param_month));
    ResultDate = ResultDate + StrToInt(param_day);

    String format = ExplodeByBackslash2(param_format, "'", "'", false)[0].text;  // ��������� ������ �� �������
    String result;

    DateTimeToString(result, format, ResultDate);

    return result;
}


/* ���������� ��������� ������ ����� ����������� ���������� � �������
*/
String calculateValue(const String& value)
{
    ParameterizedText paramText(value);
    paramText.replaceVariables(systemVariables);
    paramText.replaceVariables(systemFunctions);
    String result = paramText.getText();

    return result;
}

//---------------------------------------------------------------------------
//
__fastcall TForm1::TForm1(TComponent* Owner)
	: TForm(Owner)
{
    // ������ ������ �������
    m_vTabColor.push_back(static_cast<TColor>(RGB(180,255,20)));     // green
    m_vTabColor.push_back(static_cast<TColor>(RGB(120,230,90)));     // green
    m_vTabColor.push_back(static_cast<TColor>(RGB(0,190,90)));       // green
    m_vTabColor.push_back(static_cast<TColor>(RGB(0,190,210)));      // blue
    m_vTabColor.push_back(static_cast<TColor>(RGB(90,225,255)));     // blue
    m_vTabColor.push_back(static_cast<TColor>(RGB(100,176,255)));    // blue
    m_vTabColor.push_back(static_cast<TColor>(RGB(200,145,255)));    // violet
    m_vTabColor.push_back(static_cast<TColor>(RGB(255,100,220)));    // violet
    m_vTabColor.push_back(static_cast<TColor>(RGB(255,130,170)));    // red light
    m_vTabColor.push_back(static_cast<TColor>(RGB(255,100,0)));      // red
    m_vTabColor.push_back(static_cast<TColor>(RGB(255,180,50)));     // orange
    m_vTabColor.push_back(static_cast<TColor>(RGB(255,255,0)));      // yellow


    // ������ �������
    systemFunctions.addFunction("$compare", function_compare);
    systemFunctions.addFunction("$in", function_in);
    systemFunctions.addFunction("$date", function_date);
    systemFunctions.addFunction("$sql", function_date);

    //TParamRecord::addEditor("date", NULL);
    //TParamRecord::addEditor("string", NULL);
    //TParamRecord::addEditor(showListEditor);

    TParamRecord::setValueCalculator(calculateValue);

    // ������ "�������" ����
    // ��� ������ � ����������, �������� �������������
    dangerWords.reserve(4);
    dangerWords.push_back("execute");
    dangerWords.push_back("truncate");
    dangerWords.push_back("commit");
    dangerWords.push_back("drop");
    dangerWords.push_back("insert");
    dangerWords.push_back("update");
    dangerWords.push_back("delete");

    //TOraLogger* Logger = &TOraLogger::getInstance();
    OdacLog = new TOraLogger(EsaleSession, "", _username, TASKNAME, AppFullVersion);

    AppPath = ExtractFilePath(Application->ExeName);
}


//---------------------------------------------------------------------------
// ���������� ���� �� �������
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

    if (Auth() && PrepareForm())
    {
        OdacLog->WriteLog("START_APP");    // ������ � ���-�������
        //FormResize(NULL);     // ���� this->WindowState = wsMaximized
    }
    else
    {
        Application->ShowMainForm = false;
        Application->Terminate();
    }
}

//---------------------------------------------------------------------------
// ���������� ��������� ����������
bool __fastcall TForm1::PrepareForm()
{
    int result = LoadQueryList();

    bAdmin = _username.UpperCase() == "ADMIN";
    //miExecute->Visible = bAdmin;

    switch (result) {
    case -2:
		MessageBoxStop("����������� ��������� ��� �������� ������������ �������. ��������� ����� �������!");
        return false;
    case -1:
	    MessageBoxStop("�� ������� ������� ������� ����������� ��������. ��������� ����� �������!");
        return false;
    default:
        PrepareTabs();
        FillFieldsLB();
        FillParametersLV();

        ListBox1->ItemIndex = 0;    // �������� ����� ������ ������ � ������ ��������
        StatusBar1->Panels->Add();
        StatusBar1->Panels->Items[0]->Text = "�����";

        TabControl1->DoubleBuffered = true;
        Form1->Caption = "��������� ��� ������� " + AppFullVersion + " - " + _username;

        #ifndef NDEBUG
        Form1->Caption = Form1->Caption + " (Debuging...)";
        //OutputDebugString("Now starting dangerous function");
        assert(1==1);
        #endif

        return true;
    }
}

//---------------------------------------------------------------------------
// ����������� ������������ � ���������
bool __fastcall TForm1::Auth()
{
    LoginForm = new TLoginForm(Application);
    bool loggedon = LoginForm->Execute(EsaleSession);

    _username = UpperCase(LoginForm->getUsername());
    AddSystemVariable("username", _username);
    LoginForm->Free();
    return loggedon;
}

//---------------------------------------------------------------------------
// ��������� ������
int __fastcall TForm1::LoadQueryList()
{
    // ����� �������� � ������������ � ����� ������������ � ��� ������
    AnsiString Str = "select * from ("
        " select * from ("
        " SELECT qt.*, nvl(SYS.DBA_ROLE_PRIVS.GRANTED_ROLE, null) GRANTED_ROLE, row_number() over (partition by SPR_TASK_SQL2EXCEL_ID order by queryname) N FROM " + mainSpr + " qt "
        " LEFT join SYS.DBA_ROLE_PRIVS on GRANTEE = '" + _username + "'"
        " and upper(userlist) like '%ROLE=\"' || SYS.DBA_ROLE_PRIVS.GRANTED_ROLE || '\"%'"
        " ) where N=1"
        " )"
        " where fvisible=1 and (upper(userlist) like '%USER=\"" + _username + "\"%' or GRANTED_ROLE is not null)"
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

    // ������� ������
    OraQuery_SprTask->Close();
    delete OraQuery_SprTask;
	OraQuery_SprTask = NULL;

    return RecCount;
}

//---------------------------------------------------------------------------
// ��������� ������ �������� � ������
int __fastcall TForm1::DataSetToQueryList(TOraQuery* oraquery, std::vector<TQueryItem>& query_list, std::vector<TTabItem>& tab_list)
{
    InitEnvVariables();  // ������������� ���������� �����

    int RecCount = oraquery->RecordCount;
    if (RecCount <= 0) {
        return NULL;
    }

    query_list.reserve(RecCount);


	// ��������� ������ ����� ���������� � ��������
   	// ��������� ������ ��������
    oraquery->First();		// ��������� � ������ ������ (�� ������ ������)
    int i = 0;
    int tabindex = 0;
    AnsiString PrevTabName = "";

   	for( ; !oraquery->Eof; oraquery->Next()) {
        // � ���������� ���������� �� ������ ����������, ����� �� ����������� ������ ��� ��������� � ������
        TQueryItem query;
        query.fExcelFile = false;  // ���� Excel � ����
        query.fWordFile = false;   // ���� Word � ����
        query.fDbfFile = false;    // ���� Dbf � ����

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

        if (oraquery->FieldByName("fieldslist")->IsNull) {      // ������ - ����������� � ������� (�������� ��������� �����)
            if (query.param_excel.Fields.size() > 0) {                 // ���� �� ��������� ������, �� ����� �� ��������� �������� � Excel
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
// ��������� �������� �� SQL-�������
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

        OraQuery->Close();
        delete OraQuery;
    } catch(...) {
        if (OraQuery != NULL)
            delete OraQuery;
    }

    return result;
}

/* ��������� xml-����� �� ���� userparams
   ��������� ������ ���� TParamRecord
   ��������� ������ TParamRecord � ListParams ����������� ���������� TQueryItem* queryitem
 */
void TForm1::ParseUserParamsStr(AnsiString ParamStr, TQueryItem* queryitem)
{
    if (ParamStr == "")
    {
        return;
    }

    // ������������ ������ ����������
    OleXml msxml;

   	// ��������� ������ ����������
	AnsiString xmlParams;
    QueryVariables* params = &queryitem->UserParams;

    msxml.LoadXMLText(ParamStr);

    if (msxml.GetParseError() != "")
    {
        return;
    }

    Variant RootNode = msxml.GetRootNode();
    Variant node = msxml.GetFirstNode(RootNode);

    // ���������������� 2016-11-17
    //while (!node.IsEmpty())
    while ( !VarIsClear(node) )
    {
        TParamRecord* param = TParamRecord::createParameter(msxml, node);
        params->push_back(param);

        //param->visibleflg = true;
        node = msxml.GetNextNode(node);
    }
}

/* ������ xml-������ ���������� ��������
 */
void TForm1::ParseExportParamsStr(AnsiString ParseStr, TQueryItem* queryitem)
{
    if (ParseStr == "") {
        queryitem->DefaultExportType = EM_EXCEL_BLANK; // ���� ��������� �����������, �� �� ��������� ��������� ������ ��� SELECT
        return;
    }


    //EXPORTMODE FirstExportMode = EM_UNDEFINITE;
    AnsiString FirstId = "";
    queryitem->DefaultExportType = EM_UNDEFINITE;

    try {
        String attribute;
        OleXml msxml;
        msxml.LoadXMLText(ParseStr);

        if (msxml.GetParseError() != "") {
            return;
        }

        Variant RootNode = msxml.GetRootNode();
        Variant node = msxml.GetFirstNode(RootNode);


        // �� ���� id ����������, �������� ������������� ������������� �������
        queryitem->exportparam_id = "m_" + msxml.GetAttributeValue(RootNode, "default");

        int unassigned_id = 0;
        //while (!node.IsEmpty())
        while ( !VarIsClear(node) )
        {
            if (LowerCase(msxml.GetAttributeValue(node, "enable")) == "false") {
                node = msxml.GetNextNode(node);
                continue;
            }


            // ���� �� ����� �������� �������� �� ���������,
            // �� ������������ ������ ��������, � ������� �������� (���������� ��� � id = "0")
            if (queryitem->exportparam_id == "m_") {
                queryitem->exportparam_id = msxml.GetAttributeValue(node, "id", AnsiString("0"));
            }


            String sNodeName = msxml.GetNodeName(node);


            if (sNodeName == "excel")   // exportparams - excel
            {
                if (queryitem->fExcelFile)    // ��������� ������ ������ �������� ����� ����
                {
                    break;
                }

                queryitem->param_excel.id = "m_" + msxml.GetAttributeValue(node, "id");
                if (queryitem->param_excel.id == "m_")
                {
                    queryitem->param_excel.id = IntToStr(unassigned_id++);
                }

                queryitem->param_excel.title_label = msxml.GetAttributeValue(node, "title", queryitem->queryname);
                queryitem->param_excel.title_height = msxml.GetAttributeValue(node, "title-height", -1); // ������ ��������� � �������
                queryitem->param_excel.template_name = msxml.GetAttributeValue(node, "template", AnsiString(""));
                queryitem->param_excel.fUnbounded = msxml.GetAttributeValue(node, "unbounded", false);
                queryitem->param_excel.table_range_name = msxml.GetAttributeValue(node, "table_range", AnsiString(""));

                std::vector<EXCELFIELD>* ListFields = &queryitem->param_excel.Fields;

                Variant subnode = msxml.GetFirstNode(node);
                //while (!subnode.IsEmpty())
                while ( !VarIsClear(subnode) )
                {
                    if (msxml.GetNodeName(subnode) == "field")
                    {
                        EXCELFIELD field;
                        field.format = LowerCase(msxml.GetAttributeValue(subnode, "format"));
                        field.name = msxml.GetAttributeValue(subnode, "name");
                        field.width = msxml.GetAttributeValue(subnode, "width", -1);    // ������ �������
                        attribute = LowerCase(Trim(msxml.GetAttributeValue(subnode, "wraptext")));  // ������� �� ������
                        if (attribute == "false")
                        {
                            field.bwraptext = 0;
                        }
                        else if (attribute == "true")
                        {
                            field.bwraptext = 1;
                        }
                        else
                        {
                            field.bwraptext = -1;
                        }

                        ListFields->push_back(field);
                    }

                    subnode = msxml.GetNextNode(subnode);
                }

                if (queryitem->param_excel.id == queryitem->exportparam_id)
                {
                    if (queryitem->param_excel.template_name == "")    // � ������ ���� ������ ������
                    {
                        queryitem->DefaultExportType = EM_EXCEL_BLANK;
                    } else {
                        queryitem->DefaultExportType = EM_EXCEL_TEMPLATE;
                    }
                }

            }
            else if (sNodeName == "dbase4")
            {
                if (queryitem->fDbfFile)    // ��������� ������ ������ �������� ����� ����
                {
                    break;
                }
                queryitem->fDbfFile = true;
                queryitem->param_dbase.fAllowUnassignedFields = msxml.GetAttributeValue(node, "allowunassigned", false);

                queryitem->param_dbase.id = "m_" + msxml.GetAttributeValue(node, "id");
                if (queryitem->param_dbase.id == "m_")
                {
                    queryitem->param_dbase.id = IntToStr(unassigned_id++);
                }

                // ������ ������ ����� dbase4
                std::vector<DBASEFIELD>* ListFields = &queryitem->param_dbase.Fields;
                Variant subnode = msxml.GetFirstNode(node);
                // ���������������� 2016-11-17
                //while (!subnode.IsEmpty())
                while ( !VarIsClear(subnode) )
                {
                    if (msxml.GetNodeName(subnode) == "field")
                    {
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

            } else if (sNodeName == "word")
            {
                if (queryitem->fWordFile)   // ��������� ������ ������ �������� ����� ����
                {
                    break;
                }
                queryitem->fWordFile = true;

                queryitem->param_word.id = "m_" + msxml.GetAttributeValue(node, "id");
                if (queryitem->param_word.id == "m_")
                {
                    queryitem->param_word.id = IntToStr(unassigned_id++);
                }

                queryitem->param_word.template_name = msxml.GetAttributeValue(node, "template");
                queryitem->param_word.filter_main_field = msxml.GetAttributeValue(node, "filter_main_field", AnsiString(""));
                queryitem->param_word.filter_sec_field = msxml.GetAttributeValue(node, "filter_sec_field", AnsiString(""));
                queryitem->param_word.filter_infix_sec_field = msxml.GetAttributeValue(node, "filter_infix_sec_field", AnsiString(""));
                queryitem->param_word.page_per_doc = msxml.GetAttributeValue(node, "page_per_doc", 0);

                if (queryitem->param_word.id == queryitem->exportparam_id)
                {
                    queryitem->DefaultExportType = EM_WORD_TEMPLATE;
                }
            }
            else if (sNodeName == "execute")
            {
                queryitem->param_execute.id = "m_" + msxml.GetAttributeValue(node, "id");
                if (queryitem->param_execute.id == "m_") {
                    queryitem->param_execute.id = IntToStr(unassigned_id++);
                }

                if (queryitem->param_execute.id == queryitem->exportparam_id) {
                    queryitem->DefaultExportType = EM_PROCEDURE;
                }
            }

            // ������ ����� ���� ����������� �������� ���������� � �������
            /*if (FirstId == "") {
                AnsiString sId = "m_" + msxml.GetAttributeValue(node, "id")
                if (sId == "m_")
                    sId = "0";
                FirstId = sId;
            }*/

            node = msxml.GetNextNode(node);
        }

        if (queryitem->DefaultExportType == EM_UNDEFINITE) {
            if (queryitem->param_excel.template_name == "")    // � ������ ���� ������ ������
            {
                queryitem->DefaultExportType = EM_EXCEL_BLANK;
            }
            else
            {
                queryitem->DefaultExportType = EM_EXCEL_TEMPLATE;
            }
        }
    }
    catch (...)
    {
    }
}

/*
 */
void __fastcall TForm1::Run(EXPORTMODE ExportMode, int Tag)
{
    // ��������� �� ������������� �� ������� � ��
    if (CheckLock(StrToInt(CurrentQueryItem->dbname)))
    {
        return;
    }

    // ��������� �� ������������� �� ������� � ��
    if (CurrentQueryItem->dbname2 !="" && CheckLock(StrToInt(CurrentQueryItem->dbname2)))
    {
        return;
    }

    THREADOPTIONS* threadopt = new THREADOPTIONS;
    threadopt->queryName = CurrentQueryItem->queryname;
    switch (ExportMode)
    {
        case EM_PROCEDURE: {
        // �������������, ��� ����� ��������� ����������� ��������� ������
            String msg = "��������! ���������� ������� ������� ����� �������� � ������������ ��������� ������. ����������?";
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
                // ����� ���� ���������� ���������
                SaveDialog1->Options.Clear();
                SaveDialog1->Options << ofFileMustExist;
                SaveDialog1->Filter = "MS Excel ����� (*.xlsx)|*.xlsx|��� ����� (*.*)|*.*";
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
                // ����� ���� ���������� ���������
                SaveDialog1->Options.Clear();
                SaveDialog1->Options << ofFileMustExist;
                SaveDialog1->Filter = "MS Excel ����� (*.xlsx)|*.xlsx|��� ����� (*.*)|*.*";
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
            // ��������� �� �������������� �����-�������
            String TemplateFullName = AppPath + CurrentQueryItem->param_word.template_name;
            if(!FileExists(TemplateFullName)) {
                MessageBoxStop("���� ������� " + TemplateFullName + " �� ������.");
                return;
            }

            // ����� ���� ���������� ���������
            SaveDialog1->Options.Clear();
            SaveDialog1->Options << ofFileMustExist;
            SaveDialog1->Filter = "MS Word ����� (*.doc)|*.doc|��� ����� (*.*)|*.*";
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
            SaveDialog1->Filter = "DBase4 ����� (*.dbf)|*.dbf|��� ����� (*.*)|*.*";
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

/* ���������� �������, �������������� "�� ���������"
 */
void __fastcall TForm1::ActionDefaultRunExecute(TObject *Sender)
{
    Run(CurrentQueryItem->DefaultExportType);
}

/**/
void __fastcall TForm1::ActionAsProcedureExecute(TObject *Sender)
{
    Run(CurrentQueryItem->DefaultExportType);
}


/* ������� � ���� Excel
 */
void __fastcall TForm1::ActionExportExcelFileExecute(TObject *Sender)
{
    if (CurrentQueryItem->param_excel.template_name == "")
    {
        Run(EM_EXCEL_BLANK, 1);
    }
    else
    {
        Run(EM_EXCEL_TEMPLATE, 1);
    }

}

/* ������� � Excel (� ������)
 */
void __fastcall TForm1::ActionExportExcelBlankExecute(TObject *Sender)
{
    if (CurrentQueryItem->param_excel.template_name == "")
    {
        Run(EM_EXCEL_BLANK, 0);
    }
    else
    {
        Run(EM_EXCEL_TEMPLATE, 0);
    }
}

/* ������� � MS Word
 */
void __fastcall TForm1::ActionExportWordFileExecute(TObject *Sender)
{
    Run(EM_WORD_TEMPLATE);
}

/* ������� � ���� DBASE4
 */
void __fastcall TForm1::ActionExportDbfFileExecute(TObject *Sender)
{
    Run(EM_DBASE4_FILE);
}

/* ������� ��� �������� �����
 */
void __fastcall TForm1::FormClose(TObject *Sender, TCloseAction &Action)
{
    QueryList.clear();
    TabList.clear();
    dangerWords.clear();

    OdacLog->WriteLog("CLOSE_APP");    // ������ � ���-�������

}

/* ������������ ���������� �� ������� ������� (�����������) ��������
 */
bool TForm1::TestParameters(const QueryVariables* queryParams)
{
   for(QueryVariables::const_iterator variable = queryParams->begin(); variable != queryParams->end(); variable++)
   {
        for (std::vector<String>::iterator injection = dangerWords.begin(); injection != dangerWords.end(); injection++ )
        {
            if ( (*variable)->getType() == "string" && (*variable)->getValue().Pos( (*injection) ) )
            {
                return false;   // test failed
            }
        }
   }
   return true; // test passed
}

/* ������������ � ������ sql-������� ����� ����������� � / * * ... * * /
  �������� ��������� � ������ �� ��������
  ������� / * * � * * /
  �������� ������ �������
 */
String TForm1::GetSQL(const String& SQLText, QueryVariables* queryParams) const
{
    DinamicControlExit(NULL);

    int nDangerWords = dangerWords.size();

    // ������ �� ��������  (��������� ����������)
    // ��������� � ��������� �������!!!

    // ��� �������� ��������� ������ �� ����������.!!!!!!!!!!!!!!!!!!!!! 2016-19-09
    if ( queryParams!= NULL && !TestParameters(queryParams) )
    {
        return "SQL text contents injections.";
    }

    // ������� ������ (�������� ��������� �� ��������)
	std::vector<EXPLODESTRING> sqlstring;
    sqlstring = ExplodeByBackslash(SQLText, "/**", "**/", true);

	for (unsigned int i = 0; i < sqlstring.size(); i++)   // ���� �� ���������� � �������� ������
    {
        EXPLODESTRING *item;
        item = &sqlstring[i];
    	if (item->fBacksleshed)  			// ������ (�����������) ���������� � ������ ������� ���� ������ ��������  /** �������� **/
        {
            TReplaceFlags replaceflags = TReplaceFlags() << rfReplaceAll << rfIgnoreCase;
            item->text = StringReplace(item->text, "/**", "", replaceflags);
 			item->text = StringReplace(item->text, "**/", "", replaceflags);     // ������� **_/

            // ����������� ����������
            item->text = calculateValue(item->text);


            // ���� ������ ��������� ������������
            if ( queryParams != NULL )
            {
                for (QueryVariables::iterator it = (*queryParams).begin(); it != (*queryParams).end(); it++)
                {   //�������� --�������� �� ��������
                    TParamRecord *param = (*it);
                    String paramName = param->getName();
                    String paramValue = param->getValue();

                    if ( paramName != "" && item->text.Pos(":" + paramName) > 0 )
                    {
                        if ( param->isDeleted() )
                        {
                            item->text = "";    // ������ ���� /_** **_/  �������
                        }
                        else
                        {
 			                item->text = StringReplace(item->text, ":" + paramName, paramValue, replaceflags);
                        }
                    }
                }
            }
        }
    }

    // �������� ������ � ������
    AnsiString result = Implode(sqlstring, "");

    return result;
}


//---------------------------------------------------------------------------
// ������������� ���������� �����
// ��������! ��� ������� ������� �� ������� ��������� ������ � ���� ������
// �������� ��������� EsaleSession.
void TForm1::InitEnvVariables()
{
    // ����������� ���� �������
    // � ���� �������
    // ����� ���������� ���������� ���������
    // �������� ������� ������������ ���� � �� Oracle
    TOraQuery *OraQuery = new TOraQuery(NULL);
    OraQuery->Session = EsaleSession;

    //OraQuery->SQL->Add("select * from raion where substr(:username,3,2) = substr(uuser,3,2)");

    TOraQuery *EnvVarQueries = new TOraQuery(NULL);



    OraQuery->SQL->Add("select * from " + envSpr);
    OraQuery->Open();
    while(!OraQuery->Eof)
    {
        EnvVarQueries->Session = m_sessions[OraQuery->FieldByName("dbname")->AsInteger];
        EnvVarQueries->SQL->Clear();

        String sqlText = GetSQL(OraQuery->FieldByName("sqltext1")->AsString);
        EnvVarQueries->SQL->Add(sqlText);
        EnvVarQueries->Open();

        // ��������� ������ ��������� ����� - �� ��������
        for (int i = 1; i <= EnvVarQueries->FieldCount; i++)
        {
            AddCustomVariable(EnvVarQueries->Fields->FieldByNumber(i)->DisplayName,
                EnvVarQueries->Fields->FieldByNumber(i)->AsString);
        }
        // ��������� ������ ��������� ����� - �� �������
        //while (!EnvVarQueries->Eof) {
        //    AddEnvVariable(EnvVarQueries->FieldByName("name")->AsString,
        //        EnvVarQueries->FieldByName("value")->AsString);
        //    EnvVarQueries->Next();
        //}

        OraQuery->Next();
    }
    delete EnvVarQueries;
    EnvVarQueries = NULL;

    OraQuery->SQL->Clear();
    //delete OraQuery;
    //OraQuery = NULL;



    // ��������� ������ ����� � ���� ������ _roles
    String roles = "{";
    //TOraQuery *OraQuery = new TOraQuery(NULL);
    //OraQuery->Session = EsaleSession;
    OraQuery->SQL->Add("select * from session_roles");
    OraQuery->Open();
    while (!OraQuery->Eof)
    {
        roles += "'"+OraQuery->FieldByName("role")->AsString+"'";
        OraQuery->Next();
        if (!OraQuery->Eof)
        {
            roles += ",";
        }
    }
    roles +="}";
    OraQuery->Close();
    delete OraQuery;

    AddSystemVariable("roles", roles.LowerCase());
}

/* ��������� ���������������� ���������� */
void __fastcall TForm1::AddCustomVariable(const String& name, const String& value)
{
    systemVariables.addVariable(CUSTOM_VARIABLES_PREFIX + name, value);
}

/* ��������� ��������� ���������� */
void __fastcall TForm1::AddSystemVariable(const String& name, const String& value)
{
    systemVariables.addVariable(SYSTEM_VARIABLES_PREFIX + name, value);
}

/* ���������� ������� � ������ */
void __fastcall TForm1::threadListener(unsigned int threadId, TThreadSelectMessage message)
{

    //if ( message.getStatus() == WM_THREAD_PROCEED_BEGIN_SQL )
    //{
    //}

    switch ( message.getStatus() ) {
        case WM_THREAD_PROCEED_BEGIN_SQL:
        {
            OdacLog->WriteLog("START_THREAD", threadId, message.getMessage());    // ������ � ���-�������
            this->Enabled = false;
            Application->CreateForm(__classid(TForm_Wait), &Form_Wait);
            Form_Wait->Label3->Caption = "���������� �������...";
            TotalTime = 0;
            Timer1->Enabled = true;
            Form_Wait->Show();
            break;
        }
        case WM_THREAD_PROCEED_BEGIN_FETCH:
        {
            Form_Wait->Label3->Caption = "���������� ������...";
            break;
        }
        case WM_THREAD_PROCEED_BEGIN_DOCUMENT:
        {
            Form_Wait->Label3->Caption = "�������� ���������...";
            break;
        }
        case WM_THREAD_PROCEED_EXCEL:
            break;

        case WM_THREAD_COMPLETED_SUCCESSFULLY:
        {
            this->Enabled = true;
            Timer1->Enabled = false;

            OdacLog->WriteLog("STOP_THREAD", threadId, message.getMessage());    // ������ � ���-�������
            if (message.getFiles().size() > 0)
            {
                showResults(message.getFiles());
            }

            Form_Wait->Release();
            ts = NULL;
            break;
            // ����� ����������� OnThreadSuccess
        }


        case WM_THREAD_USER_CANCEL:      // ������ �������������
        {
            this->Enabled = true;
            Timer1->Enabled = false;
            Form_Wait->Release();
            OdacLog->WriteLog("STOP_THREAD", threadId, message.getMessage());    // ������ � ���-�������
            break;
        }
        case WM_THREAD_ERROR_BD_CANT_CONNECT:
        {
            this->Enabled = true;
            Timer1->Enabled = false;
            Form_Wait->Release();
            OdacLog->WriteLog("STOP_THREAD_ERR", threadId, message.getMessage());    // ������ � ���-�������
            AnsiString msg = "�� ������� ������������ � ���� ������. \n���������� ��������� ������ �������.";
            MessageBoxInf(msg);
            break;
        }
        case WM_THREAD_ERROR_NULL_RESULTS:
        {
            this->Enabled = true;
            Timer1->Enabled = false;
            Form_Wait->Release();
            OdacLog->WriteLog("STOP_THREAD_ERR", threadId, message.getMessage());    // ������ � ���-�������
            AnsiString msg = "� ������ �������� ���������� �������� 0 �����.\n���������� �������� ��������� ������.";
            MessageBoxInf(msg);
            break;
        }
        case WM_THREAD_ERROR_TOO_MORE_RESULTS:
        {
            this->Enabled = true;
            Timer1->Enabled = false;
            Form_Wait->Release();
            OdacLog->WriteLog("STOP_THREAD_ERR", threadId, message.getMessage());    // ������ � ���-�������
            AnsiString msg = "� ������ �������� ���������� �������� ����� 1 ���. �����.\n���������� �������� ��������� ������.";
            MessageBoxInf(msg);
            break;
        }
        case WM_THREAD_ERROR_PARAMS_INCORRECT:
        {
            this->Enabled = true;
            Timer1->Enabled = false;
            Form_Wait->Release();
            OdacLog->WriteLog("STOP_THREAD_ERR", threadId, message.getMessage());    // ������ � ���-�������
            AnsiString msg = "�� ������� ��������� ������.\n��������� ������������ ���������� ������.";
            MessageBoxStop(msg);
            break;
        }
        case WM_THREAD_ERROR_IN_PROCESS:
        {
            this->Enabled = true;
            Timer1->Enabled = false;
            Form_Wait->Release();
            OdacLog->WriteLog("STOP_THREAD_ERR", threadId, message.getMessage());    // ������ � ���-�������
            AnsiString msg = "�������� ������ � �������� ��������� ������.\n" + message.getMessage();
            MessageBoxStop(msg);
            break;
        }
        case WM_THREAD_ERROR_IN_PROCESS_ALT:
        {
            this->Enabled = true;
            Timer1->Enabled = false;
            Form_Wait->Release();
            OdacLog->WriteLog("STOP_THREAD_ERR", threadId, message.getMessage());    // ������ � ���-�������
            AnsiString msg = message.getMessage();
            MessageBoxStop(msg);
            break;
        }
        case WM_THREAD_ERROR_OPEN_QUERY:
        {
            this->Enabled = true;
            Timer1->Enabled = false;
            Form_Wait->Release();
            OdacLog->WriteLog("STOP_THREAD_ERR", threadId, message.getMessage());    // ������ � ���-�������
            AnsiString msg = "�������� ������ ��� �������� ��������� �������.\n" + message.getMessage();
            MessageBoxStop(msg);
            break;
        }
        case WM_THREAD_ERROR_OPEN_QUERY2:
        {
            this->Enabled = true;
            Timer1->Enabled = false;
            Form_Wait->Release();
            OdacLog->WriteLog("STOP_THREAD_ERR", threadId, message.getMessage());    // ������ � ���-�������
            AnsiString msg = "�������� ������ ��� �������� ���������������� �������.\n" + message.getMessage();
            MessageBoxStop(msg);
            break;
        }
        case WM_THREAD_EXECUTE_DONE:
        {
            this->Enabled = true;
            Timer1->Enabled = false;
            Form_Wait->Release();
            OdacLog->WriteLog("STOP_THREAD", threadId, message.getMessage());    // ������ � ���-�������
            AnsiString msg = "��������� �������.\n";
            MessageBoxInf(msg);
            break;

        }
        case WM_THREAD_EXECUTE_ERROR:
        {
            this->Enabled = true;
            Timer1->Enabled = false;
            Form_Wait->Release();
            OdacLog->WriteLog("STOP_THREAD_ERR", threadId, message.getMessage());    // ������ � ���-�������
            AnsiString msg = "�������� �������������� ������ � �������� ���������� �������.\n";
            MessageBoxStop(msg);
            break;
        }
    }

}

//---------------------------------------------------------------------------
// ���������� ���������� ������
void __fastcall TForm1::showResults(std::vector<String> fileList)
{
    std::vector<String>::size_type n = fileList.size();
    if (n == 0)
    {
        return;
    }
    else
    {
        String s = "";
        int MaxOut = 5;         // ���������!!!!!!!!!!!!!!!!!!!!
        int nOut = n > MaxOut ? MaxOut : n;
        for (int i = 0; i < nOut; i++)     // ������� ����� ������ ������ MaxOut ������
        {
            s += "\n" + fileList[i] ;
        }
        if (n > MaxOut)
        {
            s += "\n...";       // ���� ������ > MaxOut
        }

        AnsiString filepath = ExtractFilePath(fileList[0]);
        MessageBoxInf("��������� �������� � �������� " + filepath +
            "\n����� (" + IntToStr(n) + " ��.):" + s);
        try
        {
            ExploreFile(this->Handle, fileList[0]);
        }
        catch (...)
        {
        }
    }
}

//---------------------------------------------------------------------------
// ���������� ���������� ������
/*void __fastcall TForm1::OnThreadSuccess(EXPORTMODE ExportMode, std::vector<String> vResultFiles)
{
    switch (ExportMode) {
    case EM_EXCEL_TEMPLATE:
    case EM_EXCEL_BLANK:
        if (vResultFiles.size() > 0) {
            MessageBoxInf("��������� �������� � ����� " + vResultFiles[0]);
            try {
                ExploreFile(this->Handle, vResultFiles[0]);
            } catch (...) {
            }
        }
        break;
    case EM_DBASE4_FILE:
        //AnsiString filepath = ExtractFilePath(threadopt->filename);
        MessageBoxInf("��������� �������� � ����� " + vResultFiles[0]);
        try {
            ExploreFile(this->Handle, vResultFiles[0]);
        } catch (...) {
        }
        break;
    case EM_WORD_TEMPLATE:
        String s = "";
        int MaxOut = 5;         // ���������!!!!!!!!!!!!!!!!!!!!
        int n = vResultFiles.size();
        int nOut = n > MaxOut ? MaxOut : n;
        for (int i = 0; i < nOut; i++) {    // ������� ����� ������ ������ MaxOut ������
            s += "\n" + vResultFiles[i] ;
        }
        if (n > MaxOut)
            s += "\n...";       // ���� ������ > MaxOut

        AnsiString filepath = ExtractFilePath(vResultFiles[0]);
        MessageBoxInf("��������� �������� � �������� " + filepath +
            "\n����� (" + IntToStr(n) + " ��.):" + s);
        try {
            ExploreFile(this->Handle, vResultFiles[0]);
        } catch (...) {
        }
        break;
     }
}*/

//---------------------------------------------------------------------------
//
void __fastcall TForm1::ListBox1DrawItem(TWinControl *Control, int Index,
      TRect &Rect, TOwnerDrawState State)
{
    // ����������� �����
    static const TColor colorText1 = static_cast<TColor>(RGB(0,0,0));   // ���� ������ ������ ������
    static const TColor colorText1Sel = static_cast<TColor>(RGB(255,255,255));  // ���� ������ ������ ������
    static const TColor colorText2 = static_cast<TColor>(RGB(80,80,80));        // ���� ������ ������ ������
    static const TColor colorText2Sel = static_cast<TColor>(RGB(255,255,255));  // ���� ������ ������ ������
    static const TColor colorBkOdd = static_cast<TColor>(RGB(240,240,240));     // ���� ���� ��������� ��������
    static const TColor colorBkEven = static_cast<TColor>(RGB(255,255,255));    // ���� ���� ������� ��������

    TListBox *pListBox = static_cast <TListBox *> (Control);
    TCanvas *pCanvas = pListBox->Canvas;


    std::string text1;      // ������� ������
    std::string text2;      // ������ ������

    std::string str = pListBox->Items->Strings[Index].c_str();


    
    //���������� ������ �� ������� �������� ������ \n
    //AnsiString str = pListBox->Items->Strings[Index];
    // ��������� ������
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

    // ���������
    pCanvas->Lock();

    if (Index % 2 == 1) {       //������ �� ������ ������
        pCanvas->Brush->Color = colorBkOdd;
        pCanvas->FillRect(Rect);
    } else {
        pCanvas->Brush->Color = colorBkEven;
        pCanvas->FillRect(Rect);    // ������� ������� (������ ���)
    }

    // if the item is selected...
    if (State.Contains(odSelected)) {
        pCanvas->Font->Color = clHighlightText;
        pCanvas->Brush->Color = clHighlight;
        pCanvas->FillRect(Rect);
    }

    // ����� ������
    const int offset = 3;       // set this to offset the text

    if (State.Contains(odSelected))     // ���� ��� ������ ������
        pCanvas->Font->Color = colorText1Sel;    // ���� ������
    else
        pCanvas->Font->Color = colorText1;    // ���� ������


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
        if (State.Contains(odSelected))     // ���� ��� ������ ������
            pCanvas->Font->Color = colorText2Sel;    // ���� ������
        else
            pCanvas->Font->Color = colorText2;    // ���� ������

        //pCanvas->Font->Color = RGB(80,80,80);    // ���� ������
        pCanvas->Font->Style = pCanvas->Font->Style >> fsBold;// >> fsUnderline;
        //pCanvas->Font->Size=pCanvas->Font->Size + 12;
        pCanvas->Font->Height=pCanvas->Font->Height - 6;
        //pCanvas->Font->Style = fsNormal;
        pCanvas->TextOut(
            Rect.Left + offset, Rect.Top + th + 3,
            text2.c_str() );
    }

    if (State.Contains(odFocused)) {    // ������� ����� ������
        // remove the focus rect (i.e., XOR it away)
        DrawFocusRect(pCanvas->Handle, &Rect);
    }

    pCanvas->Unlock();     
}


//---------------------------------------------------------------------------
// ������� ������ ���������� � ListView
void TForm1::FillParametersLV()
{
    //MessageBoxStop("��������� �������� ������������ ��������!\n�������� �������� ����������!");

    if (ListBox1->ItemIndex < 0)
    {
        ListBox1->ItemIndex = 0;
    }

    TQueryItem* qi = TabList[TabControl1->TabIndex].queryitem[ListBox1->ItemIndex];
    if (CurrentQueryItem == qi)
    {
        return;
    }

    // ������ �����!!! ��������� ���������� �������
    CurrentQueryItem = qi;

    ParamsLV->Items->BeginUpdate();
    ParamsLV->Items->Clear();


	for (unsigned int i = 0; i < CurrentQueryItem->UserParams.size(); i++)
    {
        TParamRecord *record = CurrentQueryItem->UserParams[i];

        if ( !record->isVisible() ) {
            continue;
        }

        /*if (record->visibleif != "" && CheckCondition(record->visibleif) != true) {
            continue;
        }       */

        TListItem *Item = ParamsLV->Items->Add();
        Item->Caption = record->getCaption();
        Item->SubItems->Add(record->getDisplay());
    }
    ParamsLV->Items->EndUpdate();

    // ������ ��� ������ - ����������� ������ "�� ���������"
    // �����, ��������, ����� ����������, ��� ��� �� ���������
    // ���� ����������� � Excel, �� ������ ����������� � ������
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
    /*case EM_WORD_MEMORY:    // ���� �� �����������
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
// ���������� ����������� ���������  - �� ��������� ���������� false
bool TForm1::CheckCondition(AnsiString condition)
{
    if (condition.Trim() == "")
        return false;

    vector<AnsiString> t;
    t = Explode(condition, "=", false);
    if (t.size() == 1)


    if (t.size() == 1) {
        t[0] = t[0].LowerCase();
        if ( t[0]== "true")
            return true;
        else
            return false;
    }
    else if (t.size() != 2) {
        return false;
    } else
        return t[0] == t[1];


/*    String lparam = ReplaceVariables(&m_env_var, t[0]);  // ����������� ���������������� �������� � �����
    lparam = ReplaceVariables(&queryitem->Variables, t[0]);  // ����������� ��������, ������������ � QUERYITEM

    String rparam = ReplaceEnvVariables(&m_env_var, t[1]);
    rparam = ReplaceEnvVariables(&queryitem->Variables, t[1]); */

//    return lparam == rparam;
}

/*
//---------------------------------------------------------------------------
// ����������� ����������, ������������ � m_env_var (����������)
String TForm1::ReplaceEnvVariables(AnsiString condition)
{
    TReplaceFlags replaceflags = TReplaceFlags() << rfReplaceAll << rfIgnoreCase;

    for (std::vector<ENVITEM>::iterator it = m_env_var.begin() ; it != m_env_var.end(); ++it) {
        condition = StringReplace(condition, it->name, it->value, replaceflags);     // ������� * * /
    }

    return condition;
} */


//---------------------------------------------------------------------------
// ��������� �������
void __fastcall TForm1::PageControl1DrawTab(TCustomTabControl *Control,
      int TabIndex, const TRect &Rect, bool Active)
{
    TTabControl *pTabControl = static_cast <TTabControl *> (Control);
    TCanvas *pCanvas = Control->Canvas;

    // �����, ������� ����� �������� �� �������
    AnsiString TabCaption = TabControl1->Tabs->Strings[TabIndex];


    // ���������� �����
    TColor colorBk;     // ���� ����
    TColor colorText = static_cast<TColor>(RGB(0,0,0));     // ���� ������

    // ���� �� 8 ������ (� �.�. ������ �����)
    //int m_ColorIndex = TabIndex % m_vTabColor.size();  // ��������� ����� �� �����

    //colorBk = m_vTabColor[m_ColorIndex];
    colorBk = ColorByIndex(TabIndex);

    pCanvas->Brush->Color = colorBk;
    pCanvas->Font->Color = colorText;    // ���� ������


    // ������� ���������� �� 90 �������� �����
    HFONT hfontTimes;         // Font handle
    LOGFONT logfont;          // Logical font structure

    // First, clear all fields.
    memset (&logfont, 0, sizeof (logfont));

    // ������� ������������ �����
    logfont.lfHeight = pTabControl->Font->Height;   //-13;
    logfont.lfWidth = 0;
    logfont.lfEscapement = 900;         // ��������
    logfont.lfOrientation = 900;        // ��������  900
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

    // ���������
    pCanvas->Lock();    // �������� ������ ����� ����������
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
// ������� ������ �������� � ListBox 
void TForm1::FillFieldsLB()
{
    Panel3->Color = m_vTabColor[TabControl1->TabIndex % m_vTabColor.size()];

    TTabItem* TabItem = &TabList[TabControl1->TabIndex];
    ListBox1->Items->BeginUpdate();
    ListBox1->Clear();
    for (QueryItemList::size_type i = 0; i < TabItem->queryitem.size(); i++) {
        AnsiString sName = TabItem->queryitem[i]->queryname;   // QueryName
        AnsiString sFields = TabItem->queryitem[i]->fieldslist; // Fields
        ListBox1->Items->Add(sName + "\\n" + sFields);
     }
    ListBox1->Items->EndUpdate();
    if (ListBox1->Items->Count > 0)
        ListBox1->ItemIndex = 0;
}

//---------------------------------------------------------------------------
// �������� �������
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
// �������� ������ �������� TabControl ��� ������� ������� �����
void __fastcall TForm1::FormResize(TObject *Sender)
{
    TabControl1->Width = TabControl1->RowCount() * TabControl1->TabHeight;  // �������� ������ TabControl1 � ����������� �� ���������� ����� � ���
}

//---------------------------------------------------------------------------
// ������ ���������� �������
void __fastcall TForm1::Timer1Timer(TObject *Sender)
{
    TotalTime += 0.001 * Timer1 -> Interval;
    AnsiString sec = IntToStr((int) TotalTime % 60);
    AnsiString min = IntToStr((int) TotalTime / 60);
    sec = str_pad(sec.c_str(), 2, "0", STR_PAD_LEFT).c_str();
    StatusBar1->Panels->Items[0]->Text =  min + ":" + sec;

    Application->ProcessMessages();
}

/* ������ ������� ���������� ��� �������������� ���������
   ������������� �������� ������������� ��������
*/
void __fastcall TForm1::DinamicControlExit(TObject *Sender)
{
    if (Sender != NULL)
    {
        TControl *Control = (TControl*)Sender;
        Control->Visible = false;
        TParamRecord *param = CurrentQueryItem->UserParams[Control->Tag];
        if (param->type == "list")
        {
            TComboBox *ComboBox = (TComboBox*)Sender;
            if (ComboBox->ItemIndex >=0) {
                param->setValue(ComboBox->ItemIndex);
            }
        }
        else if (param->type == "date")
        {
            TDateTimePicker* DateTimePicker = (TDateTimePicker*)Sender;
            param->setValue(DateTimePicker->DateTime);
        } else if (param->type == "string") {
            String text = "";
             try {
                text = dynamic_cast<TEdit*>(Sender)->Text;
             } catch (...) {
                text = dynamic_cast<TMaskEdit*>(Sender)->Text;
             }
             param->setValue(text);
        } else if (param->type == "integer" || param->type == "float" ) {
            param->setValue(NumEdit1->Text);
        }

        ParamsLV->Items->Item[ParamsLV->Tag]->SubItems->Strings[0] = param->getDisplay();
    } else {
        DateTimePicker1->Visible = false;
        Edit1->Visible = false;
        MaskEdit1->Visible = false;
        ComboBox1->Visible = false;
        NumEdit1->Visible = false;
    }
}

/* ������������ ��������� KeyDown �� ������������ ����� �������������� ����������
 */
void __fastcall TForm1::DinamicControlOnKeyDown(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
    if (Key == VK_ESCAPE) { // �������� (����������� �������������� ������� Esc)
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

/*void TForm1::ShowDateTimePicker(TParamRecord* param)
{
} */

/* ������ �������������� �������� ���������
 */
void __fastcall TForm1::OnEditParam()
{
    if (ParamsLV->Selected == NULL)
    {
        return;
    }

    TRect rect = ParamsLV->Items->Item[0]->DisplayRect(drLabel);

    int top = ParamsLV->Selected->Top;
    int left = ParamsLV->Columns->Items[0]->Width+1;
    int width =  ParamsLV->Columns->Items[1]->Width;
    int height = rect.Height();

    // ���������� ������ ��������� � ������ ����� visible
    int LV_itemindex = ParamsLV->Selected->Index;  // ������ �������� � ParamsLV
    ParamsLV->Tag = LV_itemindex;      // ������� ���������� ������� � ParamsLV

    int n = 0;
    int paramitem_index = 0;  // ������ �������� � ������� ���������� CurrentQueryItem->Parameters
    for (paramitem_index = 0; paramitem_index < CurrentQueryItem->UserParams.size(); paramitem_index++) {
        TParamRecord *param = CurrentQueryItem->UserParams[paramitem_index];

        if ( !param->isVisible() ) { // ������� ������� ��������� (visibleflg = false)
            continue;
        }
        if (n == LV_itemindex) {
            break;
        }
        n++;
    }

    TParamRecord *param;
    param = CurrentQueryItem->UserParams[paramitem_index];
    //param->show();




    //TWinControl *Control;

    Edit1->Visible = false;
    ComboBox1->Visible = false;
    DateTimePicker1->Visible = false;


    if (param->type == "date")
    {
        DateTimePicker1->Parent = ParamsLV;       // ������������� ������� ����������� ��������
        DateTime_SetFormat(DateTimePicker1->Handle, "dd.MM.yyyy");
        DateTimePicker1->Width = width;
        DateTimePicker1->Top = top;
        DateTimePicker1->Left = left;
        DateTimePicker1->Height = height-2;
        DateTimePicker1->Font = ParamsLV->Font;
        DateTimePicker1->Font->Size = 10;
        DateTimePicker1->Tag = paramitem_index;  // ������� ���������� ������� � �������

        try {
            DateTimePicker1->Date = StrToDate(param->getDisplay());
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
        ComboBox1->Tag = paramitem_index;  // ������� ���������� ������� � �������

        //ComboBox1->Clear();

        TStringList* list = dynamic_cast<TListParameter*>(param)->getItems();
        ComboBox1->Items->Assign(list);
        list->Free();
        ComboBox1->ItemIndex = dynamic_cast<TListParameter*>(param)->getItemIndex();


        ComboBox1->Visible = true;
        ComboBox1->SetFocus();

    } else if (param->type == "string") {

        if (  static_cast<TStringParameter*>(param)->mask == "" )
        {
            //TEdit* EditBox = new TEdit(this);
            Edit1->Parent = ParamsLV;
            Edit1->Width = width;
            Edit1->Top=top;
            Edit1->Left=left;
            Edit1->Height=height-2;
            Edit1->Font = ParamsLV->Font;
            Edit1->Font->Size = 10;
            Edit1->Tag = paramitem_index;
            Edit1->Text = param->getDisplay();
            Edit1->Visible = true;
            Edit1->SetFocus();
        }
        else
        {
            MaskEdit1->EditMask = ((TStringParameter*)param)->mask;
            MaskEdit1->Parent = ParamsLV;
            MaskEdit1->Width = width;
            MaskEdit1->Top=top;
            MaskEdit1->Left=left;
            MaskEdit1->Height=height-2;
            MaskEdit1->Font = ParamsLV->Font;
            MaskEdit1->Font->Size = 10;
            MaskEdit1->Tag = paramitem_index;
            MaskEdit1->Text = param->getDisplay();
            MaskEdit1->Visible = true;
            MaskEdit1->SetFocus();
        }
    }
    else if (param->type == "integer" || param->type == "float")
    {
        NumEdit_bUseSign = true;
        NumEdit_bUseDot = param->type == "float";
        NumEdit1->Parent = ParamsLV;
        NumEdit1->Width = width;
        NumEdit1->Top = top;
        NumEdit1->Left = left;
        NumEdit1->Height = height-2;
        NumEdit1->Font = ParamsLV->Font;
        NumEdit1->Font->Size = 10;
        NumEdit1->Tag = paramitem_index;
        NumEdit1->Text = param->getDisplay();
        NumEdit1->Visible = true;
        NumEdit1->SetFocus();
    };
}

//---------------------------------------------------------------------------
// ������������ Click � ������ ����������
void __fastcall TForm1::ParamsLVClick(TObject *Sender)
{
    OnEditParam();
}

//---------------------------------------------------------------------------
// ���������� ������ ����������
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
        TListItem *Item = ParamsLV->Items->Item[ParamsLV->Tag];  // ������� ���������� ������� � ParamsLV
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

        bool bSeparator = CurrentQueryItem->UserParams[Item->Index]->type == "separator";

        if (!bSeparator) {
            // ���� �� �����������
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
            // ���� �����������
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
    AnsiString str = "������� �������:\n"
    "F1 \t- ��������� �������\n"
    "Ctrl+F1 \t- ���������� ������ ���������� �����\n"
    "Ctrl+S \t- ���������� ����� ��������� �������\n"
    "Ctrl+Alt+S \t- ���������� ����� ���������������� �������\n"
    "F8 \t- ��������� ������\n"
    "Esc \t- ����� �� ���������\n"
    "\n"
    "�������:\n"
    "���������� �������� ��������: " + mainSpr + "\n"
    "���������� �������� ��� ���������� �����: " + envSpr + "\n";

    MessageBoxInf(str, "������� SQL2Excel");
}
//---------------------------------------------------------------------------
// Ctrl+C
// ����������� ����� �������
void __fastcall TForm1::ActionCopyQueryExecute(TObject *Sender)
{
    AnsiString str = GetSQL(CurrentQueryItem->querytext, &CurrentQueryItem->UserParams);
    Clipboard()->AsText = str;
    //FormShowQuery->ShowQuery(str, CurrentQueryItem->queryname);
}


//---------------------------------------------------------------------------
// Ctrl+S
// ���������� ����� ��������� �������
void __fastcall TForm1::ActionShowMainQueryExecute(TObject *Sender)
{
    if (CurrentQueryItem->querytext != "") {
        AnsiString str = GetSQL(CurrentQueryItem->querytext, &CurrentQueryItem->UserParams);
        FormShowQuery->ShowQuery(str, "SQL-����� ��������� ������� \"" + CurrentQueryItem->queryname + "\"");
    } else {
        MessageBoxInf("����� ��������� ������� �����������.\n");
    }

}

//---------------------------------------------------------------------------
// Ctrl+Alt+S
// ���������� ����� ���������������� �������
void __fastcall TForm1::ActionShowSecondaryQueryExecute(TObject *Sender)
{
    if (CurrentQueryItem->querytext2 != "") {
        AnsiString str = GetSQL(CurrentQueryItem->querytext2, &CurrentQueryItem->UserParams);
        FormShowQuery->ShowQuery(str, "SQL-����� ���������������� ������� \"" + CurrentQueryItem->queryname + "\"");
    } else {
        MessageBoxInf("����� ���������������� ������� �����������.\n");
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
    //ShowMessage("���� ���������� ���������. ���������� ������� �� ���������.");
}

//---------------------------------------------------------------------------
// �������� ���������� ��
// -1 - ������������� ������� � ����� ��, >= 0 - ������������� ������� �� ������� ��
bool __fastcall TForm1::CheckLock(int dbindex)
{
    CheckLockQuery->ParamByName("dbindex")->AsInteger = dbindex;
    if (CheckLockQuery->Active)
        CheckLockQuery->Refresh();
    else
        CheckLockQuery->Execute();

    if (CheckLockQuery->RecordCount > 0) {
        if (bAdmin) {
            String msg = "������� � ���� ������ " + m_sessions[dbindex]->Name + " �������������.\n���������� � ����� ������?";
            //if (MessageBoxQuestion(msg) != IDNO) {
            return MessageBoxQuestion(msg) == IDNO;
        } else {
            String msg = "������� � ���� ������ " + m_sessions[dbindex]->Name + " �������������.\n���������� ��������� ������ ������ �������.";
            MessageBoxInf(msg);

            return true;
        }
    }

    return false;
}

//---------------------------------------------------------------------------
// ������� ������ - PopUp-���� ������� � DBASE4, Excel
void __fastcall TForm1::BitBtn2Click(TObject *Sender)
{

    GetParentForm(BitBtn2)->ActiveControl = NULL;
    TPoint point = BitBtn2->ClientToScreen(TPoint(0,0));
    PopupMenu1->Popup(point.x, point.y);
}

//---------------------------------------------------------------------------


/* ������ ��������� ������
   � ��������� ���
*/
void __fastcall TForm1::DoExport(THREADOPTIONS* threadopt)
{
    // ���������� ���������� �������
    if (ListBox1->ItemIndex < 0) {          // ���� �� ������ ������ � ������
        MessageBoxStop("�������� ������!");
        return;
    }

    // ����� ������������ OraSession
    TOraSession *orasession = NULL;
    TOraSession *orasession2 = NULL;
    try {                                       // ���� � ������� � �� �� ������ ������ � ���� DBNAME
        int dbname = 0;
        int dbname2 = 0;
        dbname = StrToInt(CurrentQueryItem->dbname);
        orasession = m_sessions[dbname];       // �������� ������

        if (CurrentQueryItem->dbname2 != "") {
            dbname2 = StrToInt(CurrentQueryItem->dbname2);
            orasession2  = m_sessions[dbname2];    // �������������� ������
        }
    } catch(...) {
        MessageBoxStop("�������� ���� ������ ������� �� �����!\n���������� � ��������������.");
        return;
    }


    // ���������������� 2016-11-18
    //OdacLog->WriteLog("Execute query", CurrentQueryItem->queryname);    // ������ � ���-�������

    // ������������ ������ �������
    AnsiString querytext;
    AnsiString querytext2;
    querytext = GetSQL(CurrentQueryItem->querytext, &CurrentQueryItem->UserParams);    // �������� ������
    querytext2 = GetSQL(CurrentQueryItem->querytext2, &CurrentQueryItem->UserParams);  // �������������� ������, ����� �� �������������� (������������ � ������� MS Word)

    threadopt->querytext = querytext;
    threadopt->querytext2 = querytext2;
    threadopt->queryitem = CurrentQueryItem;
    threadopt->ParentFormHandle = this->Handle;

    threadopt->TemplateOraSession = orasession;
    threadopt->TemplateOraSession2 = orasession2;

    // �������� � ������ ������
    ts = new ThreadSelect(true, threadopt);    // ������� ���������������� �����
    //ts = new ThreadSelect(true, threadopt, threadListener);    // ������� ���������������� �����

    //void (*f)(const String&, int);
    //f = threadListener;



    //ts->SetThreadOpt(threadopt);    // �������� ���������
    ts->Resume();                   // ���������
}

//---------------------------------------------------------------------------
// � ���������...
void __fastcall TForm1::ActionAboutAppExecute(TObject *Sender)
{
    //
    String MsgStr = "��������� ��� ���������� �������\nSQL2Excel v." + AppVersion + " (" + AppBuild + ")"
        "\n"
        "\nCopyright � 2014-2016"
        "\n����������� ������ ��� \"���������������\""
        "\n"
        "\n�����:"
        "\n�������-����������� ������ ��"
        "\n������������ ������� ��� \"���������������\""
        "\n�.�. ����������"
        "\n"
        "\ne-mail: V.Ovchinnikov@cf.esbt.ru";
    MessageBoxInf(MsgStr, "� ��������� SQL2Excel");
}

//---------------------------------------------------------------------------
//
void __fastcall TForm1::PopupMenu1Popup(TObject *Sender)
{
    // �������/����������� ������� ����
    ActionExportDbfFile->Enabled = CurrentQueryItem->fDbfFile;
    ActionExportWordFile->Enabled = CurrentQueryItem->fWordFile;
    //ActionExportExcelFile->Enabled = CurrentQueryItem->fExcelFile;
    //ActionExportExcelMemory->Enabled = CurrentQueryItem->fExcelFile;
}

//---------------------------------------------------------------------------
// ������� �� ������ ����
void __fastcall TForm1::ListBox1MouseDown(TObject *Sender,
      TMouseButton Button, TShiftState Shift, int X, int Y)
{
    ListBox1->ItemIndex = ListBox1->ItemAtPos(TPoint(X,Y),false);
    FillParametersLV();
}

//---------------------------------------------------------------------------
// ������� �� ������ ����, � ����� �� ������� �����, ����
void __fastcall TForm1::ListBox1Click(TObject *Sender)
{
    FillParametersLV();
}

//---------------------------------------------------------------------------
// ��������� ��� NumEdit1 - TEdit ������ ���������� (TNumEdit)
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
// ��������� ��� NumEdit1 - TEdit ������ ���������� (TNumEdit)
void __fastcall TForm1::NumEdit1KeyPress(TObject *Sender, char &Key)
{
    NumEdit_SelStartOld = NumEdit1->SelStart;
}

//---------------------------------------------------------------------------
// ����� �� ���������
void __fastcall TForm1::ActionApplictionExitExecute(TObject *Sender)
{
    //if (MessageBoxQuestion("�� ������� ��� ������ ����� �� ���������?") != IDNO) {
        Close();
    //}
}

/* ������� ������ ���������� �����
*/
void __fastcall TForm1::ActionShowEnvironmentExecute(TObject *Sender)
{
    AnsiString str;
    str += "���������� �����:\n";
    str += systemVariables.getVariables();

    MessageBoxInf(str, "������ ���������� ����� SQL2Excel");
}
//---------------------------------------------------------------------------

