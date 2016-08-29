//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "FShowQuery.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TFormShowQuery *FormShowQuery;
//---------------------------------------------------------------------------
__fastcall TFormShowQuery::TFormShowQuery(TComponent* Owner)
    : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void TFormShowQuery::ShowQuery(const AnsiString& Text, AnsiString Caption)
{

    this->SQLText = Text;
    //this->Caption = "SQL-текст запроса \"" + Caption + "\"";
    this->Caption = Caption;
    this->ShowModal();
}

//---------------------------------------------------------------------------
// Отображает текст запроса, раскрашивает его
void __fastcall TFormShowQuery::FormShow(TObject *Sender)
{
    SQLTextRichEdit->Text = SQLText;
    SQLTextRichEdit->Update();


    SQLTextRichEdit->SelStart = 0;
    SQLTextRichEdit->SelLength = SQLText.Length();
    SQLTextRichEdit->SelAttributes->Color = clNavy;
    SQLTextRichEdit->SelAttributes->Style = SQLTextRichEdit->SelAttributes->Style >> fsBold;

    SQLText = SQLText.UpperCase();

    std::vector <AnsiString> KeyWords;
    KeyWords.reserve(120);

    KeyWords.push_back("ABS");
    KeyWords.push_back("AND");
    KeyWords.push_back("AS");
    KeyWords.push_back("ALL");
    KeyWords.push_back("ASC");
    KeyWords.push_back("AVG");
    KeyWords.push_back("ANY");
    KeyWords.push_back("AUTOMATIC");

    KeyWords.push_back("BY");
    KeyWords.push_back("BETWEEN");
    KeyWords.push_back("BEGIN");

    KeyWords.push_back("CASE");
    KeyWords.push_back("CEIL");
    KeyWords.push_back("COALESCE");
    KeyWords.push_back("COUNT");
    KeyWords.push_back("CONNECT");
    KeyWords.push_back("CUBE");
    KeyWords.push_back("COMMIT");

    KeyWords.push_back("DECLARE");
    KeyWords.push_back("DELETE");
    KeyWords.push_back("DECODE");
    KeyWords.push_back("DISTINCT");
    KeyWords.push_back("DESC");
    KeyWords.push_back("DIMENSION");

    KeyWords.push_back("ELSE");
    KeyWords.push_back("ELSIF");
    KeyWords.push_back("END");
    KeyWords.push_back("EXIT");
    KeyWords.push_back("EXISTS");

    KeyWords.push_back("FOR");

    KeyWords.push_back("GROUP");
    KeyWords.push_back("GROUPING");
    KeyWords.push_back("GREATEST");

    KeyWords.push_back("HAVING");

    KeyWords.push_back("JOIN");

    KeyWords.push_back("FLOOR");
    KeyWords.push_back("FROM");


    KeyWords.push_back("IF");
    KeyWords.push_back("INNER");
    KeyWords.push_back("IS");
    KeyWords.push_back("IN");
    KeyWords.push_back("INSERT");
    KeyWords.push_back("INTO");
    KeyWords.push_back("INTERSECT");
    KeyWords.push_back("IGNORE");
    KeyWords.push_back("ITERATE");


    KeyWords.push_back("ORDER");
    KeyWords.push_back("OVER");
    KeyWords.push_back("OR");
    KeyWords.push_back("ON");
    KeyWords.push_back("OUT");
    KeyWords.push_back("ONLY");
    KeyWords.push_back("OFFSET");

    KeyWords.push_back("PARTITION");
    KeyWords.push_back("PIVOT");

    KeyWords.push_back("RANK");
    KeyWords.push_back("REVERSE");
    KeyWords.push_back("REPLACE");
    KeyWords.push_back("ROLLUP");
    KeyWords.push_back("ROUND");
    KeyWords.push_back("RULES");
    KeyWords.push_back("RIGHT");
    KeyWords.push_back("ROW");
    KeyWords.push_back("ROWS");
    KeyWords.push_back("ROWID");
    KeyWords.push_back("ROWNUM");

    KeyWords.push_back("THEN");
    KeyWords.push_back("TRUNC");
    KeyWords.push_back("TRIM");
    KeyWords.push_back("TABLE");
    KeyWords.push_back("TO");

    KeyWords.push_back("SOME");
    KeyWords.push_back("SUBSTR");
    KeyWords.push_back("SELECT");
    KeyWords.push_back("SYSDATE");
    KeyWords.push_back("SUM");
    KeyWords.push_back("SETS");

    KeyWords.push_back("UPDATE");
    KeyWords.push_back("UPPER");
    KeyWords.push_back("UNION");

    KeyWords.push_back("LOOP");
    KeyWords.push_back("LEFT");
    KeyWords.push_back("LIKE");
    KeyWords.push_back("LAG");
    KeyWords.push_back("LEAD");
    KeyWords.push_back("LIST");
    KeyWords.push_back("LENGTH");
    KeyWords.push_back("LEAST");
    KeyWords.push_back("LEVEL");
    KeyWords.push_back("LOWER");


    KeyWords.push_back("MODEL");
    KeyWords.push_back("MOD");
    KeyWords.push_back("MINUS");
    KeyWords.push_back("MAX");
    KeyWords.push_back("MIN");
    KeyWords.push_back("MEASURES");

    KeyWords.push_back("NAV");
    KeyWords.push_back("NVL");
    KeyWords.push_back("NUMBER");
    KeyWords.push_back("NULL");
    KeyWords.push_back("NOT");
    KeyWords.push_back("NATURAL");
    KeyWords.push_back("NEXT");

    KeyWords.push_back("XOR");

    KeyWords.push_back("VALUES");

    KeyWords.push_back("WITH");
    KeyWords.push_back("WHERE");
    KeyWords.push_back("WHEN");
    KeyWords.push_back("WHILE");


    KeyWords.push_back("FETCH");
    KeyWords.push_back("FOR");



    TStringList* SepList = new TStringList;
    SepList->Capacity = 36;

    SepList->Add(" ");
    SepList->Add("(");
    SepList->Add(")");
    SepList->Add("[");
    SepList->Add("]");
    SepList->Add("{");
    SepList->Add("}");
    SepList->Add("<");
    SepList->Add(">");
    SepList->Add("-");
    SepList->Add("+");
    SepList->Add("*");
    SepList->Add("/");
    SepList->Add("\\");
    SepList->Add("&");
    SepList->Add("$");
    SepList->Add("!");
    SepList->Add("@");
    SepList->Add("#");
    SepList->Add("$");
    SepList->Add("%");
    SepList->Add("^");
    SepList->Add("?");
    SepList->Add("=");
    SepList->Add("|");
    SepList->Add("""");
    SepList->Add("~");
    SepList->Add("\n");
    SepList->Add("\r");
    SepList->Add("\t");
    SepList->Add("'");
    SepList->Add(".");
    SepList->Add(",");
    SepList->Add(":");
    SepList->Add("`");
    SepList->Add(";");


    int ntextlen = SQLText.Length();

    for (int i=0; i<KeyWords.size(); i++) {

        int npos = PosEx(KeyWords[i], SQLText, 1);
        int nlen = KeyWords[i].Length();
        while (npos > 0) {
            if (npos+nlen == ntextlen) {
                npos = 0;
                continue;
            }

            AnsiString prevchar;  // Символ перед
            AnsiString nextchar;  // Символ после
            npos == 1? prevchar = " " : prevchar = SQLText[npos-1];
            npos+nlen > ntextlen? nextchar = " " : nextchar = SQLText[npos+nlen];


            bool bFirstDelim = false;
            bool bLastDelim = false;
            for (int j = 0; j < SepList->Count; j++) {
                if (!bFirstDelim && SepList->Strings[j] == prevchar) {
                    bFirstDelim = true;
                    //break;
                }
                if (!bLastDelim && SepList->Strings[j] == nextchar) {
                    bLastDelim = true;
                    //break;
                }
            }
            if (!bFirstDelim || !bLastDelim)  // Если слева или справа не делиметры тогда пропускаем
            {
                npos =  PosEx(KeyWords[i], SQLText, npos+1);
                continue;
            }

            SQLTextRichEdit->SelStart = npos-1;
            SQLTextRichEdit->SelLength  = nlen;
            SQLTextRichEdit->SelAttributes->Color = clGreen;
            SQLTextRichEdit->SelAttributes->Style = SQLTextRichEdit->SelAttributes->Style << fsBold;
            npos =  PosEx(KeyWords[i], SQLText, npos+1);
        }
    }



    // Выделение блоков
    std::vector <TTextBlock> BlockList;
    BlockList.reserve(3);
    BlockList.push_back(TTextBlock("'","'", clBlue));
    BlockList.push_back(TTextBlock("--","\n", clRed));
    BlockList.push_back(TTextBlock("/*","*/", clRed));



    int startpos = 1;


    while (startpos > 0) {
        int minpos = ntextlen;
        int minelem = -1;
        int npos = 0;
        int nblocksCount = BlockList.size();

        for (int i=0; i < nblocksCount; i++) {
            npos =  PosEx(BlockList[i].sOpen, SQLText, startpos);
            if (npos > 0 && minpos > npos) {
                minpos = npos;
                minelem = i;
            }
        }


        if (minelem < 0)
            break;

        TTextBlock* tb = &BlockList[minelem];

        int npos2 = PosEx(tb->sClose, SQLText, minpos + 1);
        if (npos2 == 0)
            npos2 = SQLText.Length();
        /*if (npos2 == 0)
            break;   */


        //int openlen = tb->sOpen.Length();
        int closelen = tb->sClose.Length();

        SQLTextRichEdit->SelStart = minpos - 1;
        SQLTextRichEdit->SelLength = npos2 - minpos + closelen;

        SQLTextRichEdit->SelAttributes->Color = tb->FontColor;
        SQLTextRichEdit->SelAttributes->Style = SQLTextRichEdit->SelAttributes->Style >> fsBold;

        startpos = npos2 + closelen;
    }

    SQLTextRichEdit->SelStart = 0;
    SQLTextRichEdit->SelLength = 0;

    BlockList.clear();
    //SepList.
    KeyWords.clear();


}

//---------------------------------------------------------------------------
//
void __fastcall TFormShowQuery::CancelExecute(TObject *Sender)
{
    Close();        
}
//---------------------------------------------------------------------------
//
void __fastcall TFormShowQuery::FormResize(TObject *Sender)
{
    SQLTextRichEdit->Refresh();
}
//---------------------------------------------------------------------------
//
void __fastcall TFormShowQuery::FormClose(TObject *Sender,
      TCloseAction &Action)
{
    SQLTextRichEdit->Text = "";

}
//---------------------------------------------------------------------------
// Копировать текст запроса в буфер обмена
void __fastcall TFormShowQuery::ItemCopyClick(TObject *Sender)
{
    SQLTextRichEdit->CopyToClipboard();
}
//---------------------------------------------------------------------------
// Скрывает/отображает пункты контекстного меню
void __fastcall TFormShowQuery::RichEditPopupMenuPopup(TObject *Sender)
{
    ItemCopy->Enabled = SQLTextRichEdit->SelText != "";
    ItemSelectAll->Enabled = SQLTextRichEdit->Text != "";
}
//---------------------------------------------------------------------------
// Выделить все
void __fastcall TFormShowQuery::ItemSelectAllClick(TObject *Sender)
{
    SQLTextRichEdit->SelectAll();
}

//---------------------------------------------------------------------------
// Поиск Вниз
void __fastcall TFormShowQuery::SearchDownBtnClick(TObject *Sender)
{
    AnsiString FindText = Edit1->Text;
    int FoundAt, StartPos, ToEnd;
    TSearchTypes mySearchTypes = TSearchTypes();
  /*if (FindDialog1->Options.Contains(frMatchCase))
	mySearchTypes << stMatchCase;
  if (FindDialog1->Options.Contains(frWholeWord))
	mySearchTypes << stWholeWord; */

    if (SQLTextRichEdit->SelLength)
	    StartPos = SQLTextRichEdit->SelStart + SQLTextRichEdit->SelLength;
    else
	    StartPos = 0;

    ToEnd = SQLTextRichEdit->Text.Length() - StartPos;
    FoundAt = SQLTextRichEdit->FindText(FindText, StartPos, ToEnd, mySearchTypes);
    if (FoundAt != -1)
    {
	    SQLTextRichEdit->SelStart = FoundAt;
	    SQLTextRichEdit->SelLength = FindText.Length();
	    SQLTextRichEdit->SetFocus();
        //SQLTextRichEdit->
    } else
        Beep();    
}

//---------------------------------------------------------------------------
// Поиск Вверх
void __fastcall TFormShowQuery::SearchUpBtnClick(TObject *Sender)
{
    int FoundAt = -1;
    int p;

    do
    {
        p = SQLTextRichEdit->FindText(Edit1->Text, FoundAt + 1, SQLTextRichEdit->SelStart - FoundAt, TSearchTypes() << stMatchCase);
        if(p >= 0)
            FoundAt = p;
    }
    while(p >= 0);
 
    if (FoundAt != -1)
    {
        SQLTextRichEdit->SelStart = FoundAt;
        SQLTextRichEdit->SelLength = Edit1->Text.Length();
        SQLTextRichEdit->SetFocus();
    } else
        Beep();    
}
//---------------------------------------------------------------------------
// Copy
void __fastcall TFormShowQuery::CopyBtnClick(TObject *Sender)
{
    //Clipboard()->AsText = SQLText;
    Clipboard()->AsText = SQLTextRichEdit->Text;

}
//---------------------------------------------------------------------------



