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

/* ������������� �������� ��������� � ������������ � ��������� ��������� �� ������
 */
void TListParameter::setValue(int index)
{

    int n = 0;  // ������ �������� � ������� �� ������ �������
    int i = 0;  // ������ �������� (��� �������� ������ �������), � �����
    TParamlistItem* item = NULL;
    for(TListParameter::ListItemIterator it = listitem.begin(); it != listitem.end(); it++, i++)
    {
        if (!it->visibleflg) {
            continue;
        }
        if (n == index) {
            item = it;
            break;
        }
        if (n > index) {
            break;
        }
        n++;
    }

    if (item != NULL)
    {
        value = item->value;
        result = item->result;
        display = item->label;

    } else {
        if (n > 0) {
            setValue(0);
        } else {
            value = "";
            result = "";
            display = "";
        }
    }
}


String (*TParamRecord::_calculate)(const String&);

/* ������ ��������� ������� ��� ��������� ���������
 */
void TParamRecord::setValueCalculator(const CalculateFunction &calculate)
{
    _calculate = calculate;
}

/* ���������� ����� ������� ��� ��������� ���������
   ��������� ��������� �� ��� �������
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


/* ������� ���������� ��� ���������, � ����������� �� ���� ���������� � xml
 */
TParamRecord* TParamRecord::createParameter(const MsxmlWorks &xml, Variant node)
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
TParamRecord* TParamRecord::createDefault(const MsxmlWorks &xml, Variant node)
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

    // ����!!!!!!!
    parent = xml.GetAttributeValue(node, "parent");



    // visibleif
    if (visible == "" && visibleif != "") {  // visible ����� ��������� ��� visibleif
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
        {   // visible ����� ��������� ��� visibleif
            visibleflg = false;
        }
        else
        {
           visibleflg = true;
        }
    }

    // deleteif  - ������� ���� /**...**/ ���� ���� = true
    if (!xml.GetAttribute(node, "deleteif").IsEmpty())
    {// ���� � xml ����������� �������� value
        deleteifflg = true;
        deleteifvalue = xml.GetAttributeValue(node, "deleteif").UpperCase();
    }
    else
    {
        deleteifflg = false;
    }

}

/* ������� �������� ���� List
 */
TListParameter::TListParameter(const MsxmlWorks &xml, Variant node)
{
    createDefault(xml, node);
    Variant subnode = xml.GetFirstNode(node);

    // ���� � ������ list-a ����������� �������� value
    // �� value ����� ����� ������� �������� �� ������
    bool bParamValueExist = !xml.GetAttribute(node, "value").IsEmpty();

     /*if (src != "") {    // ���� ����� sql-������

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
    } else {            // ���� ����� ������ ��������
    */
        // ���� � ������ �������� ���������� list ����������� �������� value
        // �� ����������� �������� value ��� ������� item-a
        bool bValueAutoInc = xml.GetAttribute(subnode, "value").IsEmpty();
        int i = 0;
        while (!subnode.IsEmpty()) {
            TParamlistItem item;
            item.value = bValueAutoInc? IntToStr(i++) : xml.GetAttributeValue(subnode, "value");
            item.result = xml.GetAttributeValue(subnode, "result", item.value);
            item.label = xml.GetAttributeValue(subnode, "label");
            item.visible = Trim(LowerCase(xml.GetAttributeValue(subnode, "visible")));
            item.visibleif = Trim(LowerCase(xml.GetAttributeValue(subnode, "visibleif")));


            //visible = true;
            // ��������� ������� ���������
            // � ������ ����, ��� �������� visible ������������ ������� visibleif
            if (item.visible == "" && item.visibleif != "")
            {
                String condition = calculate(item.visibleif);

                if (condition == "true")
                {
                    item.visibleflg = true;
                }
                else
                {
                    item.visibleflg = false;
                }

                //item.visibleflg = CheckCondition(condition);
                //item.visibleflg = CheckCondition(item.visibleif);
                //if (record->visibleif != "" && CheckCondition(record->visibleif) != true) {
                //AnsiString s = "s";
            }
            else
            {
                if (item.visible == "false")
                {  // visible ����� ��������� ��� visibleif
                    item.visibleflg = false;
                }
                else
                {
                    item.visibleflg = true;
                }
            }

            listitem.push_back(item);
            subnode = xml.GetNextNode(subnode);
        }

        // ������ �������� �� ��������� ������ ������� �� ������
        // �� ������ ���� �������� param.value �� ����� ����
        if (!bParamValueExist && value == "" && listitem.size() > 0)
        {
            value = listitem[0].value;
        }


    // ������ ����������� param.display � ������������ � param.value (�������� �� ������)
    for (int j = 0; j < listitem.size(); j++)
    {
        if (value == listitem[j].value)
        {
            display = listitem[j].label;
            break;
        }
    }
}

String TListParameter::getValue()
{
    //return value;
    return result;
}

/* ������� �������� ���� String
 */
TStringParameter::TStringParameter(const MsxmlWorks &xml, Variant node)
{
    createDefault(xml, node);
    display = value;
    mask = xml.GetAttributeValue(node, "mask");
}

/* ������� �������� ���� Date
 */
TDateTimeParameter::TDateTimeParameter(const MsxmlWorks &xml, Variant node)
{
    createDefault(xml, node);
    format = xml.GetAttributeValue(node, "format");

    if (value == "") {
        value = DateToStr(Now());
    }
    display = value;

    try {
        // ��������������� ���� � ������ � ������ �������
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

/* ������� �������� ���� Integer
 */
TIntegerParameter::TIntegerParameter(const MsxmlWorks &xml, Variant node)
{
    createDefault(xml, node);
    display = value;
}

/* ������� �������� ���� Float
 */
TFloatParameter::TFloatParameter(const MsxmlWorks &xml, Variant node)
{
    display = value;
}

/* ������� �������� ���� Separator
 */
TSeparatorParameter::TSeparatorParameter(const MsxmlWorks &xml, Variant node)
{
    createDefault(xml, node);
    //display = value;
}

/* ������� �������� ���� Variable
 */
TVariableParameter::TVariableParameter(const MsxmlWorks &xml, Variant node)
{
    visibleflg = false;
}


//---------------------------------------------------------------------------

#pragma package(smart_init)
