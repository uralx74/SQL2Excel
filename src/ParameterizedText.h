/* ParameterizedText.h
   Класс для работы с параметризованным текстом.

   Автор: vsovchinnikov
   Дата создания: 2016-09-12
 */

#ifndef ParameterizedTextH
#define ParameterizedTextH

#include "variables.h"
#include "vector.h"
#include "map.h"

class ParameterizedText
{
public:
    ParameterizedText(const String& text);
    void setSouceText(const String& text);
    void reset();
    bool replaceVariables(const Variables& variables);
    String getText();

private:
    String _sourceText;
    String _text;
    TReplaceFlags replaceflags;
};



//---------------------------------------------------------------------------
#endif // ParameterizedTextH
