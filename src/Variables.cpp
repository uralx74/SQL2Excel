//---------------------------------------------------------------------------
/* ����� ��� ������ � �����������
 */

#include <vcl.h>
#pragma hdrstop

#include "Variables.h"

/*
 */
void split(const String &s, char delim, vector<String> &elems)
{
    stringstream ss;
    ss.str(s.c_str());
    string item;
    while ( getline(ss, item, delim) )
    {
        elems.push_back(item.c_str());
    }
}

/*
 */
EnvFunction::EnvFunction():
    _value(),
    _func(NULL)
{
}

/*
 */
EnvFunction::EnvFunction(const String& value):
    _value(value),
    _func(NULL)
{

}

/*
 */
EnvFunction::EnvFunction( String (*func)(const std::vector<String>&) ):
    _func(func)
{
}

/*EnvFunction::EnvFunction( String (*func)(const String&) ):
    _func(func)
{

} */

/*
 */
String EnvFunction::getValue()
{
    if ( _func == NULL )
    {
        return _value;
    }
    else
    {
        return _func(_parameters);
    }
}

/*
 */
bool EnvFunction::isVariable()
{
    return _func == NULL;
}

/*
 */
void EnvFunction::setParameters(const String& parameters)
{
    _parameters.clear();
    split(parameters, ',', _parameters);

    //_parameters = parameters;
}


/*
 */
Variables::Variables():
    _prefix("")
{
    //_variables.reserve(10);
    replaceflags = TReplaceFlags() << rfReplaceAll << rfIgnoreCase;
}


/*
 */
void Variables::setPrefix(const String& prefix)
{
    _prefix = prefix;
}

/* ��������� ���������� �����
 */
void __fastcall Variables::addVariable(const String& name, const String& value)
{
    //envVariables.find(name)

    //if (!envVariables.constains(name))
    {
        _variables[_prefix + name] = value;
    }
}

/*
 */
void __fastcall Variables::addFunction(const String& name, String (*func)(const std::vector<String>&) )
{
    //varList[name] = EnvFunction(func);
    _variables.insert( std::pair<String, EnvFunction >( name, EnvFunction(func) ) );
}
/*void __fastcall Variables::addFunction(const String& name, String (*func)(const String&) )
{
    //varList[name] = EnvFunction(func);
    _variables.insert( std::pair<String, EnvFunction >( name, EnvFunction(func) ) );
}*/

/* ����������� ����������, ������������ ����������������� �����
*/
String Variables::replaceInText(const String& text)
{
    if (_variables.size() < 1 || text.Length() < 1)
    {
        return text;
    }

    String result = text;


    /*for (EnvVariables::const_iterator it = _variables.begin(); it != _variables.end(); it ++)
    {
        Result = StringReplace(Result, it->first, it->second, replaceflags);
    }*/

    for (VariableList::const_iterator it = _variables.begin(); it != _variables.end(); it ++)
    {
        int length = result.Length();
        String name = it->first;

        if (it->second.isVariable()) {
            // ���� ����������, �� ������ �������� � ������ �� ��������
            result = StringReplace(result, name, it->second.getValue(), replaceflags);
        }
        else
        {
            // ���� �������, �� ������� �� ���������, ��������� ��������
            // � �������� � ������ ������ ��� ������� � ����������� � �������
            // �� ����������� ��������
            int pos = result.Pos(name + "(");
            int nameLength = name.Length();
            if (pos == 0)
            {
                continue;
            }
            else
            {

                for(int i = pos; i <= length; i++)
                {
                    if (result[i] == ')')
                    {
                        int startPos = pos + nameLength + 1;
                        String parameters = result.SubString(startPos, i - startPos );
                        it->second.setParameters(parameters);
                        String fullName = name + "(" + parameters + ")";    // ��� ������� � �����������
                        result = StringReplace(result, fullName, it->second.getValue(), replaceflags);
                    }
                }

            }
        }
    }
    return result;
}


/*
//---------------------------------------------------------------------------
// � ����������� �������� ��� ������� � taskutil.h
String __fastcall Variables::GetValue(String value)
{
    if (value.Length() < 2 || value[1] != '_' )
    {
        return value;
    }

    //String f_date = '_date('
    //vector<String>::iterator cur;
    //for (cur = m_env_func.begin(); cur <m_env_func.end() - 1; cur++) {


    String Result;
    int n = m_env_func.size();
    for (int i = 0; i < n; i++)
    {
        if (value.Pos(m_env_func[i]) != 1)
        {
            continue;
        }

        // �������� ������ � �����������
        std::vector<EXPLODESTRING2> sqlstring;
        sqlstring = ExplodeByBackslash2(value, m_env_func[i], ")");
        std::vector<AnsiString> params;

        // ��������� ������ � ����������� � ������������ - (,)
        if (sqlstring[1].fBacksleshed)
        {
            params = Explode(sqlstring[1].text, ",", false);
        }

        int n_params = params.size();
        switch (i) {
            // ������� _date(v1, v2, p1, p2, format)
            // ���������� ����, �������������� � ������ ���������� ������� ������ �������
        case 0:
        {
            TDateTime ResultDate = Date();

             // ��������� ����������
            if ( n_params == 5)
            {
                String param_day = params[0];   // ���-�� ����
                String param_month = params[1]; // ���-�� �������
                String param_option_day = params[2];    // ����� ������� ����
                String param_option_month = params[3];  // ����� ������� �������
                String param_format = params[4];
                //break;

                // ��������� ����
                // ������� ��������� ����� ������� (���� � �����), ���� ������ ����������� �����
                // ������� ����� (0), ������ ����� (1), ��������� ����� (2)
                if (param_option_month == "1" || param_option_month == "first")
                {
                    ResultDate = EncodeDate(YearOf(ResultDate), 1, DayOf(ResultDate));
                } else if (param_option_month == "2" || param_option_month == "last")
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
                DateTimeToString(Result, format, ResultDate);
            }

            break;
        }
        case 1: // m_env_func[i] = "_sql("
        {
            Result = GetValueFromSQL(params[0], params[1]);
            break;
        }

        // ������� _compare(val1, val2)
        // ������������ ��� �������� �� ���������
        case 2:
        {
            if (n_params != 2)
            {
                Result = "error";//value;
            }
            else
            {
                Result = params[0] == params[1]? "true" : "false";
            }
            break;


        }
        // ������� _in(val1, {v1,v2,v3,...})
        // ��������� ��������� �������� �� ���������
        case 3:
        {
            if (n_params != 2)
            {
                Result = "error";
            }
            else
            {
                Result = "false";
                String value = params[0];

                String tmp = ExplodeByBackslash2(params[1], "{", "}", false)[0].text;
                std::vector<AnsiString> vset;
                vset = Explode(tmp, ",", false);

                int n_vsetsize = vset.size();
                if (n_vsetsize > 0)
                {
                    for (int j = 0; j < n_vsetsize; j++)
                    {
                        if (value == vset[j])
                        {
                            Result = "true";
                        }
                    }
                }
                else
                {        // ������
                    Result = "error";
                }
            }
            break;
        }
        }  // end of switch

    }
    return Result;
}         */











//---------------------------------------------------------------------------

#pragma package(smart_init)
