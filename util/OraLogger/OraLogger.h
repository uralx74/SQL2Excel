//---------------------------------------------------------------------------
#ifndef OraLoggerH
#define OraLoggerH

#include "Ora.hpp"
//------------------------------------------------------------------------------
//
class TOraLogger
{
private:
    AnsiString _osMacAddress;
    AnsiString _osUsername;
    AnsiString _taskUsername;
    AnsiString _taskName;
    AnsiString _appVer;
    TOraSession* _oraSession;
    TOraQuery* _oraQueryLog;
    AnsiString _appId;

public:
    TOraLogger(TOraSession* OraSession, AnsiString s_os_mac_address, AnsiString s_task_user_name, AnsiString s_task_name, AnsiString s_app_ver);
    ~TOraLogger();
    void setParameters(const AnsiString& funcName, const AnsiString& threadId = "", AnsiString descr = "");
    bool WriteLog(const AnsiString& funcName, const AnsiString& threadId = "", AnsiString descr = "", const AnsiString& prntThreadId = "");
};
//---------------------------------------------------------------------------
#endif
