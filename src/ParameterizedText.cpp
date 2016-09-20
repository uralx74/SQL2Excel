#include <vcl.h>
#pragma hdrstop

#include "ParameterizedText.h"

/*
 */
ParameterizedText::ParameterizedText(const String& text) :
    _sourceText(text),
    _text(text)

{
    replaceflags = TReplaceFlags() << rfReplaceAll << rfIgnoreCase;
    //replaceflags = replaceflags << rfReplaceAll << rfIgnoreCase;
}

/* Производит замену именованных фрагментов в тексте
 */
bool ParameterizedText::replaceVariables(const Variables& variables)
{
    _text = variables.replaceInText(_text);
}

/* Производит сброс буффера в исходное состояние
 */
void ParameterizedText::reset()
{
    _text = _sourceText;
}

/* Возвращает текущее состояние
 */
String ParameterizedText::getText()
{
    return _text;
}

void ParameterizedText::setSouceText(const String& text)
{
}


//---------------------------------------------------------------------------

#pragma package(smart_init)

