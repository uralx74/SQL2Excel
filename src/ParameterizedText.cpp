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

/* ���������� ������ ����������� ���������� � ������
 */
bool ParameterizedText::replaceVariables(const Variables& variables)
{
    _text = variables.replaceInText(_text);
}

/* ���������� ����� ������� � �������� ���������
 */
void ParameterizedText::reset()
{
    _text = _sourceText;
}

/* ���������� ������� ���������
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

