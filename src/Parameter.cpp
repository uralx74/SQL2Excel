//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Parameter.h"


TParamRecord* TParamRecord::createParameter(const MsxmlWorks &xml, Variant node)
{
    TParamRecord* param;
    AnsiString type = xml.GetAttributeValue(node, "type");

    if (type == "list")
    {
        param = new TListParameter(xml, node);
    }
    else if (type == "edit")
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
    } else {   // default
        param = new TSeparatorParameter(xml, node);
    }

    return param;
}

TParamRecord* TParamRecord::createDefault(const MsxmlWorks &xml, Variant node)
{
    type = LowerCase(xml.GetAttributeValue(node, "type"));
    name = xml.GetAttributeValue(node, "name");
    label = xml.GetAttributeValue(node, "label");
    format = xml.GetAttributeValue(node, "format");

    src = xml.GetAttributeValue(node, "src");
    dbindex = xml.GetAttributeValue(node, "dbindex");
    visible = Trim(LowerCase(xml.GetAttributeValue(node, "visible")));
    visibleif = Trim(LowerCase(xml.GetAttributeValue(node, "visibleif")));

        // ����!!!!!!!
    parent = xml.GetAttributeValue(node, "parent");

    value_src = xml.GetAttributeValue(node, "value");

    //value = xml.GetAttributeValue(node, "value");

    //value_src = ReplaceVariables(envVariables, value_src);



    //value = ReplaceVariables(queryitem->Variables, param.value_src);
    //value = GetDefinedValue(value);     // ���������� �����!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!



/*    // visibleif
    if (visible == "" && visibleif != "") {  // visible ����� ��������� ��� visibleif

        //String condition = ReplaceVariables(envVariables, visibleif);  // ����������� ���������������� �������� � �����

        //condition = ReplaceVariables(queryitem->Variables, condition);  // ����������� ��������, ������������ � QUERYITEM

        ParameterizedText paramText(condition);
        paramText.replaceVariables(systemVariables);
        String condition = paramText.getText();


        if (GetDefinedValue(condition) == "true")
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





    // deleteif
    if (!xml.GetAttribute(node, "deleteif").IsEmpty())
    {// ���� � xml ����������� �������� value
        deleteifflg = true;
        deleteifvalue = xml.GetAttributeValue(node, "deleteif").UpperCase();
    }
    else
    {
        deleteifflg = false;
    }




*/
}


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


            if (item.visible == "" && item.visibleif != "")
            {  // visible ����� ��������� ��� visibleif



                // ! ��� ���������� �����������������!
                //String condition = ReplaceVariables(envVariables, item.visibleif);  // ����������� ���������������� �������� � �����
                //condition = ReplaceVariables(queryitem->Variables, condition);  // ����������� ��������, ������������ � QUERYITEM






/*
                if (GetDefinedValue(condition) == "true")
                {
                    item.visibleflg = true;
                }
                else
                {
                    item.visibleflg = false;
                }

*/

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
    //}

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

TStringParameter::TStringParameter(const MsxmlWorks &xml, Variant node)
{
    createDefault(xml, node);
    display = value;
    mask = xml.GetAttributeValue(node, "mask");
}

TDateTimeParameter::TDateTimeParameter(const MsxmlWorks &xml, Variant node)
{
    createDefault(xml, node);
    if (value == "")
        value = DateToStr(Now());

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

TIntegerParameter::TIntegerParameter(const MsxmlWorks &xml, Variant node)
{
    createDefault(xml, node);
    display = value;
}

TSeparatorParameter::TSeparatorParameter(const MsxmlWorks &xml, Variant node)
{
    createDefault(xml, node);
    //display = value;
}


TFloatParameter::TFloatParameter(const MsxmlWorks &xml, Variant node)
{
    display = value;
}


//---------------------------------------------------------------------------

#pragma package(smart_init)
