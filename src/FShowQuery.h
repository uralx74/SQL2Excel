//---------------------------------------------------------------------------

#ifndef FShowQueryH
#define FShowQueryH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <Buttons.hpp>
#include <ComCtrls.hpp>
#include <Mask.hpp>
#include <ActnList.hpp>
#include <Clipbrd.hpp>
#include <Menus.hpp>
#include <ImgList.hpp>
#include <ToolWin.hpp>
#include <ExtCtrls.hpp>
#include <Graphics.hpp>
#include <vector>
#include <list>
#include "fmain.h"

class TTextBlock {
public:
    __fastcall TTextBlock(AnsiString sOpen, AnsiString sClose, TColor FontColor);
    AnsiString sOpen;
    AnsiString sClose;
    TColor FontColor;
};

__fastcall TTextBlock::TTextBlock(AnsiString sOpen, AnsiString sClose, TColor FontColor)
{
    this->sOpen = sOpen;
    this->sClose = sClose;
    this->FontColor = FontColor;
}



//---------------------------------------------------------------------------
class TFormShowQuery : public TForm
{
__published:	// IDE-managed Components
    TRichEdit *SQLTextRichEdit;
    TActionList *ActionList1;
    TAction *Cancel;
    TLabel *Label1;
    TPopupMenu *RichEditPopupMenu;
    TMenuItem *ItemCopy;
    TMenuItem *N2;
    TMenuItem *ItemSelectAll;
    TToolBar *ToolBar1;
    TToolButton *SearchDownBtn;
    TEdit *Edit1;
    TToolButton *SearchUpBtn;
    TToolButton *ToolButton3;
    TToolButton *CopyBtn;
    TImageList *ImageList1;
    TToolButton *ToolButton1;
    TImage *Image1;
    void __fastcall FormShow(TObject *Sender);
    void __fastcall CancelExecute(TObject *Sender);
    void __fastcall FormResize(TObject *Sender);
    void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
    void __fastcall ItemCopyClick(TObject *Sender);
    void __fastcall RichEditPopupMenuPopup(TObject *Sender);
    void __fastcall ItemSelectAllClick(TObject *Sender);
    void __fastcall SearchDownBtnClick(TObject *Sender);
    void __fastcall SearchUpBtnClick(TObject *Sender);
    void __fastcall CopyBtnClick(TObject *Sender);
private:	// User declarations
public:		// User declarations
    __fastcall TFormShowQuery(TComponent* Owner);
    void ShowQuery(const AnsiString& Text, AnsiString Caption);
    AnsiString SQLText;
};
//---------------------------------------------------------------------------
extern PACKAGE TFormShowQuery *FormShowQuery;
//---------------------------------------------------------------------------
#endif
