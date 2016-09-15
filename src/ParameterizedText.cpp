//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "ParameterizedText.h"


ParameterizedText::ParameterizedText(const String& text) :
    _sourceText(text),
    _text("")

{
    replaceflags = TReplaceFlags() << rfReplaceAll << rfIgnoreCase;
    //replaceflags = replaceflags << rfReplaceAll << rfIgnoreCase;
}

bool ParameterizedText::replaceVariables(const Variables& variables)
{
    // Подстановка переменных среды
    /*for (EnvVariables::const_iterator it = variables.begin(); it != variables.end(); it ++)
    {
        _text = StringReplace(_sourceText, it->first, it->second, replaceflags);
    }*/

    _text = variables.replaceInText(_sourceText);

}

void ParameterizedText::reset()
{
    _text = "";
}

String ParameterizedText::getText()
{
    return _text;
}


void ParameterizedText::setSouceText(const String& text)
{
}


//---------------------------------------------------------------------------

#pragma package(smart_init)

