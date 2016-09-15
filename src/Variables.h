//---------------------------------------------------------------------------

#ifndef VariablesH
#define VariablesH

#include <map.h>
#include <vector.h>
#include <sstream>
/*


*/


class EnvFunction
{
public:
    EnvFunction();
    EnvFunction(const String& value);
    //EnvFunction( String (*func)(const String&) );
    EnvFunction( String (*func)(const std::vector<String>&) );

    String getValue();
    void setParameters(const String& parameters);
    bool isVariable();

public:
    String _value;
    String (*_func)(const std::vector<String>&);
    //String (*_func)(const String&);
    std::vector<String> _parameters;
    //String _parameters;
};



/*


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
    //void __fastcall addFunction(const String& name, String (*func)(const String&) );

    void setPrefix(const String& prefix);

private:
    String _prefix;
    VariableList _variables;
    TReplaceFlags replaceflags;
};





//---------------------------------------------------------------------------
#endif // VariablesH
