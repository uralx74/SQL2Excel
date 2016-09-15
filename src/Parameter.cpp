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

        // “ест!!!!!!!
    parent = xml.GetAttributeValue(node, "parent");

    value_src = xml.GetAttributeValue(node, "value");

    //value = xml.GetAttributeValue(node, "value");

    //value_src = ReplaceVariables(envVariables, value_src);



    //value = ReplaceVariables(queryitem->Variables, param.value_src);
    //value = GetDefinedValue(value);     // ƒоработать здесь!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!



/*    // visibleif
    if (visible == "" && visibleif != "") {  // visible имеет приоритет над visibleif

        //String condition = ReplaceVariables(envVariables, visibleif);  // ѕодстановка предопределенных значений в среде

        //condition = ReplaceVariables(queryitem->Variables, condition);  // ѕодстановка значений, определенных в QUERYITEM

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
        {   // visible имеет приоритет над visibleif
            visibleflg = false;
        }
        else
        {
           visibleflg = true;
        }
    }





    // deleteif
    if (!xml.GetAttribute(node, "deleteif").IsEmpty())
    {// ≈сли в xml отсутствует параметр value
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

    // ≈сли в опци€х list-a отсутствует параметр value
    // то value будет равно первому значению из списка
    bool bParamValueExist = !xml.GetAttribute(node, "value").IsEmpty();

     /*if (src != "") {    // ≈сли задан sql-запрос

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
    } else {            // ≈сли задан список значений
    */
        // ≈сли в списке значений компонента list отсутствует параметр value
        // то проставл€ем значение value дл€ каждого item-a
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
            {  // visible имеет приоритет над visibleif



                // ! это необходимо расскоментировать!
                //String condition = ReplaceVariables(envVariables, item.visibleif);  // ѕодстановка предопределенных значений в среде
                //condition = ReplaceVariables(queryitem->Variables, condition);  // ѕодстановка значений, определенных в QUERYITEM






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
                {  // visible имеет приоритет над visibleif
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

        // «адаем значение по умолчанию равным первому из списка
        // но только если параметр param.value не задан €вно
        if (!bParamValueExist && value == "" && listitem.size() > 0)
        {
            value = listitem[0].value;
        }
    //}

    // «адаем отображение param.display в соответствии с param.value (выбираем из списка)
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
        // ѕреобразовываем дату в строку в нужном формате
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
