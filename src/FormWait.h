//---------------------------------------------------------------------------

#ifndef FormWaitH
#define FormWaitH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ExtCtrls.hpp>
#include <ImgList.hpp>
#include <Buttons.hpp>
//---------------------------------------------------------------------------
class TForm_Wait : public TForm
{
__published:	// IDE-managed Components
    TImageList *ImageList1;
    TImage *Image1;
    TTimer *Timer1;
    TLabel *Label1;
    TLabel *Label2;
    TBevel *Bevel1;
    TLabel *Label3;
    TSpeedButton *SpeedButton1;
    TSpeedButton *CancelBtn;
    void __fastcall Timer1Timer(TObject *Sender);
    void __fastcall FormCreate(TObject *Sender);
    void __fastcall FormShow(TObject *Sender);
    void __fastcall FormHide(TObject *Sender);
    void __fastcall FormCloseQuery(TObject *Sender, bool &CanClose);
    void __fastcall SpeedButton1Click(TObject *Sender);
    void __fastcall CancelBtnClick(TObject *Sender);
private:	// User declarations
    int FrameIndex;
public:		// User declarations
    __fastcall TForm_Wait(TComponent* Owner);
    void __fastcall Activate(TObject *Sender);
    void __fastcall Deactivate(TObject *Sender);
};



//---------------------------------------------------------------------------
extern PACKAGE TForm_Wait *Form_Wait;
//---------------------------------------------------------------------------
#endif
