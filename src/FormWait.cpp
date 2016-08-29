//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "FormWait.h"
#include "FMain.h"

//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm_Wait *Form_Wait;

//---------------------------------------------------------------------------
__fastcall TForm_Wait::TForm_Wait(TComponent* Owner)
    : TForm(Owner)
{
}
//---------------------------------------------------------------------------
// Обновление изображения
void __fastcall TForm_Wait::Timer1Timer(TObject *Sender)
{
    if (++FrameIndex >= ImageList1->Count)
        FrameIndex=0;

    Image1->Canvas->Lock();

    Image1->Canvas->FillRect(Image1->ClientRect); // Очищаем от старой картинки
    ImageList1->GetBitmap(FrameIndex, Image1->Picture->Bitmap);
    this->Refresh();
}
//---------------------------------------------------------------------------
void __fastcall TForm_Wait::FormCreate(TObject *Sender)
{
    this->DoubleBuffered = true;
}
//---------------------------------------------------------------------------
void __fastcall TForm_Wait::FormShow(TObject *Sender)
{
    Timer1->Enabled = true;
}
//---------------------------------------------------------------------------
void __fastcall TForm_Wait::FormHide(TObject *Sender)
{
    Timer1->Enabled = false;
}
//---------------------------------------------------------------------------
// Предотвращаем закрытие формы пользователем
void __fastcall TForm_Wait::FormCloseQuery(TObject *Sender, bool &CanClose)
{
    CanClose = false;
}

//---------------------------------------------------------------------------
void __fastcall TForm_Wait::Activate(TObject *Sender)
{
    ((TForm*)Sender)->Enabled = false;
    Show();

}

//---------------------------------------------------------------------------
void __fastcall TForm_Wait::Deactivate(TObject *Sender)
{
    ((TForm*)Sender)->Enabled = true;
    Hide();
}

//---------------------------------------------------------------------------
void __fastcall TForm_Wait::SpeedButton1Click(TObject *Sender)
{
    SendMessage(Application->MainForm->Handle, WM_SYSCOMMAND, SC_MINIMIZE, 0);
}
//---------------------------------------------------------------------------

void __fastcall TForm_Wait::CancelBtnClick(TObject *Sender)
{
//    Form1->CancelThread();
}
//---------------------------------------------------------------------------

