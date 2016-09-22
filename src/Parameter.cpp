//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Parameter.h"


/*TStrings TParamRecord::getSubItems()
{
    throw Exception("This type of parameter doesn't have subitems.");
}*/

bool TParamRecord::isVisible()
{
    return visibleflg;
}

String TParamRecord::getCaption()
{
    return label;
}

bool TParamRecord::isDeleted()
{
    return deleteifflg == true && value.UpperCase() == deleteifvalue.UpperCase();
}

String TParamRecord::getDisplay()
{
    return display;
}

String TParamRecord::getType()
{
    return type;
}

String TParamRecord::getName()
{
    return name;
}

String TParamRecord::getValue()
{
    return value;
}

void TParamRecord::setValue(const String& value)
{
    this->value = value;
    this->display = value;
}

void TParamRecord::setValue(int index)
{
    throw Exception("This type of parameter doesn't take integer value.");
}

void TParamRecord::setValue(const TDateTime& dt)
{
    throw Exception("This type of parameter doesn't take DateTime value.");
}

void TDateTimeParameter::setValue(const TDateTime& dt)
{
    display = DateToStr(dt);
    if (format == "")
    {
        value = display;
    }
    else
    {
        value = FormatDateTime(format, dt);
    }
}


void TStringParameter::setValue(const String& value)
{
    display = value;
    this->value = display;
}

/* Устанавливает значение парамтера в соответствии с выбранным элементом из списка
   Производит поиск элемента по индексу
 */
void TListParameter::setValue(int index)
{

    int n = 0;  // Индекс элемента в векторе не считая скрытые
    int i = 0;  // индекс итерации (или элемента считая скрытые), а также
    //TParamlistItem* item = NULL;
    for(TListParameter::ListItemIterator it = listitem.begin(); it != listitem.end(); it++, i++)
    {
        if (!it->visibleflg) {
            continue;
        }
        if (n == index) {
            _currentItem = it;
            break;
        }
        if (n > index) {
            break;
        }
        n++;
    }

    if (_currentItem != NULL)
    {
        value = _currentItem->value;
        //result = _currentItem->result;
        display = _currentItem->label;
        _itemIndex = n;
    } else {
        if (n > 0) {
            setValue(0);
            _itemIndex = 0;
        } else {
            _currentItem = NULL;
            _itemIndex = -1;
            //value = "Ошибка";
            //result = "Ошибка";
            //display = "Ошибка";
        }
    }
}

/* Устанавливает значение парамтера в соответствии с выбранным элементом из списка
   Производит поиск элемента по значению
 */
void TListParameter::setValue(const String& value)
{
    int n = 0;  // Индекс элемента в векторе не считая скрытые
    int i = 0;  // индекс итерации (или элемента считая скрытые), а также
   //TParamlistItem* item = NULL;
    for(TListParameter::ListItemIterator it = listitem.begin(); it != listitem.end(); it++, i++)
    {
        if (!it->visibleflg) {
            continue;
        }
        if (value == it->value) {
            _currentItem = it;
            //item = it;
            break;
        }
        n++;
    }

    if (_currentItem != NULL)
    {
        this->value = _currentItem->value;
        //result = _currentItem->result;
        display = _currentItem->label;
        _itemIndex = n;
    } else {
        if (n > 0) {
            setValue(0);
        } else {
            _currentItem = NULL;
            _itemIndex = -1;
            //value = "Ошибка";
            //result = "Ошибка";
            //display = "Ошибка";
        }
    }
}

/* Возвращает все видимые элементы списка
 */
TStringList* TListParameter::getItems()
{
    TStringList* strings = new TStringList();
    for(TListParameter::ListItemIterator it = listitem.begin(); it != listitem.end(); it++)
    {
        if (it->visibleflg) {
            strings->Add(it->label);
        }
    }
    return strings;
}


String (*TParamRecord::_calculate)(const String&);

/* Задает указатель функции для обработки выражений
 */
void TParamRecord::setValueCalculator(const CalculateFunction &calculate)
{
    _calculate = calculate;
}

/* Производит вызов функции для обработки выражиния
   используя указатель на эту функцию
 */
String TParamRecord::calculate(const String& expression)
{
    if (_calculate != NULL) {
        return _calculate(expression);
    } else {
        return expression;
    }
}

/*void TParamRecord::addEditor(BeginEditFunction &beginEdit)
{
   //editor[name] = beginEdit;
   editor.insert( std::make_pair("name", beginEdit));
} */


/* Создает конкретный вид параметра, в зависимости от типа указанного в xml
 */
TParamRecord* TParamRecord::createParameter(const OleXml &xml, Variant node)
{
    TParamRecord* param;
    AnsiString type = LowerCase(xml.GetAttributeValue(node, "type"));

    if (type == "list")
    {
        param = new TListParameter(xml, node);
    }
    else if (type == "string")
    {
        param = new TStringParameter(xml, node);
    }
    else if (type == "date")
    {
        param = new TDateTimeParameter(xml, node);
    }
    else if (type == "integer")
    {
        param = new TIntegerParameter(xml, node);
    }
    else if (type == "variable")
    {
        param = new TVariableParameter(xml, node);
    }
    else
    {   // default
        param = new TSeparatorParameter(xml, node);
    }


    return param;
}

/*
 */
TParamRecord* TParamRecord::createDefault(const OleXml &xml, Variant node)
{
    type = LowerCase(xml.GetAttributeValue(node, "type"));
    name = xml.GetAttributeValue(node, "name");
    label = xml.GetAttributeValue(node, "label");

    src = xml.GetAttributeValue(node, "src");
    dbindex = xml.GetAttributeValue(node, "dbindex");
    visible = Trim(LowerCase(xml.GetAttributeValue(node, "visible")));
    visibleif = Trim(LowerCase(xml.GetAttributeValue(node, "visibleif")));

    value_src = xml.GetAttributeValue(node, "value");
    value = calculate(value_src);

    // Тест!!!!!!!
    parent = xml.GetAttributeValue(node, "parent");



    // visibleif
    if (visible == "" && visibleif != "") {  // visible имеет приоритет над visibleif
        String condition = calculate(condition);

        if (condition == "true")
        {
            visibleflg = true;
        }
        else
        {
            visibleflg = false;
        }
    } else
    {
        if (visible == "false")
        {   // visible имеет приоритет над visibleif
            visibleflg = false;
        }
        else
        {
           visibleflg = true;
        }
    }

    // deleteif  - удалять блок /**...**/ если флаг = true
    if (!xml.GetAttribute(node, "deleteif").IsEmpty())
    {// Если в xml отсутствует параметр value
        deleteifflg = true;
        deleteifvalue = xml.GetAttributeValue(node, "deleteif").UpperCase();
    }
    else
    {
        deleteifflg = false;
    }
}


/*
TSqlListParameter::TSqlListParameter()
{

     /*if (src != "") {    // Если задан sql-запрос

        int dbindex = 0;
        //AnsiString sdbindex = "";
        //Variant DbIndexAttribute = msxml.GetAttribute(subnode, "dbindex");
        //sdbindex = msxml.GetAttributeValue(subnode, "dbindex");
        if (dbindex != "") {
            try {
                dbindex = StrToInt(param.dbindex);
            } catch (...) {}
        } else {
            try {
                dbindex = StrToInt(queryitem->dbname);  //------------------------------------------------------------------------
            } catch (...) {}
            dbindex = IntToStr(dbindex);
        }

        try {
            //TOraSession *orasession = getSessionByIndex(dbindex);
            TOraSession *orasession = m_sessions[dbindex];
            orasession->Connected = true;

            TOraQuery *OraQuery = new TOraQuery(NULL);
            OraQuery->Session = orasession;
            OraQuery->SQL->Add(param.src);
            OraQuery->Open();

            TParamlistItem item;

            while (!OraQuery->Eof) {
                item.value = OraQuery->FieldByName("value")->AsString;
                item.label = OraQuery->FieldByName("label")->AsString;
                param.listitem.push_back(item);
                OraQuery->Next();
            }

            OraQuery->Close();
            delete OraQuery;
        } catch (...) {}
    } else {            // Если задан список значений
*/

/* Создает параметр типа List
 */
TListParameter::TListParameter(const OleXml &xml, Variant node) :
    _currentItem(NULL)
{
    createDefault(xml, node);
    listitem.reserve(10);

    Variant subnode = xml.GetFirstNode(node);

    // Если в опциях list-a отсутствует параметр value
    // то value будет равно первому значению из списка
    bool bParamValueExist = !xml.GetAttribute(node, "value").IsEmpty();

    // Если в списке значений компонента list отсутствует параметр value
    // то проставляем значение value для каждого item-a
    bool bValueAutoInc = xml.GetAttribute(subnode, "value").IsEmpty();
    int i = 0;

    // Заполняем список элементов
    // в список попадают и невидимые (скрытые) элементы
    while (!subnode.IsEmpty()) {
        TParamlistItem item;
        item.value = bValueAutoInc? IntToStr(i++) : xml.GetAttributeValue(subnode, "value");
        item.result = xml.GetAttributeValue(subnode, "result", item.value);
        item.label = xml.GetAttributeValue(subnode, "label", item.value);

        if (xml.GetAttribute(subnode, "visible").IsEmpty() && !xml.GetAttribute(subnode, "visibleif").IsEmpty())
        {
            item.visibleif = Trim(LowerCase(xml.GetAttributeValue(subnode, "visibleif")));
            item.visibleflg = calculate(item.visibleif) ==  OleXml::TRUE_STR_VALUE;
        } else {
            item.visible = Trim(LowerCase(xml.GetAttributeValue(subnode, "visible", OleXml::TRUE_STR_VALUE)));
            item.visibleflg = item.visible == OleXml::TRUE_STR_VALUE;
        }

        listitem.push_back(item);
        subnode = xml.GetNextNode(subnode);
    }

    // Задаем значение по умолчанию равным первому из списка
    // но только если параметр param.value не задан явно
    if ( bParamValueExist )
    {
        setValue(value);
    }
    else
    {
        setValue(0);
    }
}

/* Возвращает индекс элемента (скрытые пропускаются)
 */
int TListParameter::getItemIndex()
{
    return _itemIndex;
}


String TListParameter::getValue()
{
    if (_currentItem != NULL)
    {
        return _currentItem->result;
    }
    else
    {
        return "";
    }
}

/* Создает параметр типа String
 */
TStringParameter::TStringParameter(const OleXml &xml, Variant node)
{
    createDefault(xml, node);
    display = value;
    mask = xml.GetAttributeValue(node, "mask");
}

/* Создает параметр типа Date
 */
TDateTimeParameter::TDateTimeParameter(const OleXml &xml, Variant node)
{
    createDefault(xml, node);
    format = xml.GetAttributeValue(node, "format");

    if (value == "") {
        value = DateToStr(Now());
    }
    display = value;

    try {
        // Преобразовываем дату в строку в нужном формате
        if (format != "") {
            AnsiString oldShortDateFormat = ShortDateFormat;
            AnsiString oldDateSeparator = DateSeparator;
            ShortDateFormat = "dd.MM.yyyy";
            DateSeparator = '.';
            value = FormatDateTime(format, StrToDate(value));
            AnsiString ShortDateFormat = oldShortDateFormat;
            AnsiString DateSeparator = oldDateSeparator;
        }
    } catch (...){
    }
}

/* Создает параметр типа Integer
 */
TIntegerParameter::TIntegerParameter(const OleXml &xml, Variant node)
{
    createDefault(xml, node);
    display = value;
}

/* Создает параметр типа Float
 */
TFloatParameter::TFloatParameter(const OleXml &xml, Variant node)
{
    display = value;
}

/* Создает параметр типа Separator
 */
TSeparatorParameter::TSeparatorParameter(const OleXml &xml, Variant node)
{
    createDefault(xml, node);
    //display = value;
}

/* Создает параметр типа Variable
 */
TVariableParameter::TVariableParameter(const OleXml &xml, Variant node)
{
    visibleflg = false;
}


//---------------------------------------------------------------------------

#pragma package(smart_init)
