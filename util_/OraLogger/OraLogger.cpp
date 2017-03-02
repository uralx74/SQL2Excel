//---------------------------------------------------------------------------


#pragma hdrstop

#include "OraLogger.h"

using namespace std;

/*
 class TOraLogger
*/


TOraLogger::TOraLogger(TOraSession* OraSession, AnsiString s_os_mac_address, AnsiString s_task_user_name, AnsiString s_task_name, AnsiString s_app_ver) :
    _oraSession(OraSession),
    _osMacAddress(s_os_mac_address),
    _taskUsername(s_task_user_name),
    _taskName(s_task_name),
    _appVer(s_app_ver)

{
	_oraQueryLog = new TOraQuery(NULL);
    _oraQueryLog->Session = OraSession;
    _oraQueryLog->SQL->Clear();
 	_oraQueryLog->CreateProcCall("pk_nasel_otdel.p_log_task_write", 0);
    _oraQueryLog->Prepare();

    randomize();
    _appId = random(9999999999);
}

TOraLogger::~TOraLogger()
{
    delete _oraQueryLog;
    _oraQueryLog = NULL;
}

//------------------------------------------------------------------------------
// Записывает в таблицу БД лог-строку
bool TOraLogger::WriteLog(const AnsiString& funcName, const AnsiString& threadId, AnsiString descr, const AnsiString& prntThreadId)
{

    if (threadId != "")
    {
        _oraQueryLog->ParamByName("p_thread_id")->Value = threadId ;

        if (prntThreadId != "")
        {
            _oraQueryLog->ParamByName("p_prnt_thread_id")->Value = prntThreadId;
        }
        else
        {
            _oraQueryLog->ParamByName("p_prnt_thread_id")->Value = _appId;
        }
    }
    else
    {
        _oraQueryLog->ParamByName("p_thread_id")->Value = _appId;
    }


    //_oraQueryLog->ParamByName("p_prnt_thread_id")->Value = (threadId == "" ? "" : _appId);
    //_oraQueryLog->ParamByName("p_thread_id")->Value = ( threadId == "" ? _appId : threadId;
    _oraQueryLog->ParamByName("p_descr")->Value = descr;
    _oraQueryLog->ParamByName("p_pc_mac")->Value = _osMacAddress;
    _oraQueryLog->ParamByName("p_task_name")->Value = _taskName;
    _oraQueryLog->ParamByName("p_func_name")->Value = funcName;
    _oraQueryLog->ParamByName("p_task_user_name")->Value = _taskUsername;
    _oraQueryLog->ParamByName("p_app_ver")->Value = _appVer;

    try {
        _oraQueryLog->ExecSQL();
        _oraQueryLog->Close();
        //_oraQueryLog->ClearFields();
    } catch (...) {
        return false;
    }

    return true;
}

//---------------------------------------------------------------------------

#pragma package(smart_init)
 