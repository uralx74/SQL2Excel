//---------------------------------------------------------------------------

#ifndef VariablesH
#define VariablesH

#include <map.h>
#include <vector.h>
#include <sstream>
/* Класс для храннения значения или ссылки на функцию,
   использующихся в качестве параметров запросов
*/
typedef String (*EnvFunctionType)(const std::vector<String>&);

class EnvFunction
{
public:
    EnvFunction();
    EnvFunction(const String& value);
    EnvFunction( String (*func)(const std::vector<String>&) );

    String getValue();
    void setParameters(const String& parameters);
    bool isVariable();

public:
    String _value;
    String (*_func)(const std::vector<String>&);
    std::vector<String> _parameters;
};



/* Класс для работы с EnvFunction
*/
typedef std::map<String, EnvFunction> VariableList;
// contains      ^ Name  ^ Value


class Variables
{
public:
    Variables();
    String replaceInText(const String& text);
    void __fastcall addVariable(const String& name, const String& value);
    void __fastcall addFunction(const String& name, String (*func)(const std::vector<String>&) );

    void setPrefix(const String& prefix);
    String getVariables();

private:
    String _prefix;
    VariableList _variables;
    TReplaceFlags replaceflags;
};


//---------------------------------------------------------------------------
#endif // VariablesH
