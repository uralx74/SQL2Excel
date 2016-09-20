//---------------------------------------------------------------------------
/* Класс для работы с переменными
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

/* Добавляет переменную среды
 */
void __fastcall Variables::addVariable(const String& name, const String& value)
{
    //envVariables.find(name)
    //if (!envVariables.constains(name))
    _variables[_prefix + name] = value;
}

/*
 */
String Variables::getVariables()
{
    String result = "";
    for (VariableList::const_iterator it = _variables.begin(); it != _variables.end(); it ++)
    {
        result += it->first + " = \"" + it->second.getValue() + "\"\n";
    }
    return result;
}

/*
 */
void __fastcall Variables::addFunction(const String& name, String (*func)(const std::vector<String>&) )
{
    //varList[name] = EnvFunction(func);
    _variables.insert( std::pair<String, EnvFunction >( name, EnvFunction(func) ) );
}

/* Подстановка переменных, определенных параметризованный текст
*/
String Variables::replaceInText(const String& text)
{
    if (_variables.size() < 1 || text.Length() < 1)
    {
        return text;
    }

    String result = text;


    for (VariableList::const_iterator it = _variables.begin(); it != _variables.end(); it ++)
    {
        int length = result.Length();
        String name = it->first;

        if (it->second.isVariable())
        {
            // Если переменная, то просто заменяем в тексте на значение
            result = StringReplace(result, name, it->second.getValue(), replaceflags);
        }
        else
        {
            // Если функция, то находим ее параметры, вычисляем значение
            // и заменяем в тексте полное имя функции с параметрами в скобках
            // на вычисленное значение
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
                        String fullName = name + "(" + parameters + ")";    // Имя функции с параметрами
                        result = StringReplace(result, fullName, it->second.getValue(), replaceflags);
                        length = result.Length();
                    }
                }

            }
        }
    }
    return result;
}


/*
//---------------------------------------------------------------------------
// В последующем вставить эту функцию в taskutil.h
String __fastcall Variables::GetValue(String value)
{
         case 1: // m_env_func[i] = "_sql("
        {
            Result = GetValueFromSQL(params[0], params[1]);
            break;
        }

    }  // end of switch

    }
    return Result;
}         */











//---------------------------------------------------------------------------

#pragma package(smart_init)
