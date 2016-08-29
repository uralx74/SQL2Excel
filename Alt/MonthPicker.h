//---------------------------------------------------------------------------
// ��������� TMonthPicker
// �����: �.�.����������
// ���������� ������ ��� "���������������"
// 2016 �.
// e-mail.: x74@list.ru
//---------------------------------------------------------------------------
// ��������:
// ������ ��������� ��������� ������������ ������� ��� � �����, ��� �����������
// ������ ����� ������.
// ��������:
// TDateTime Date - ���./���. ��������� ���� (��� ����� ����� ������)
// TDateTime MinDate - ���������� ����� ����������� �����
// TDateTime MaxDate - ���������� ����� ������������ �����
// unsigned short Month - ���./���. ��������� �����
// unsigned short Year - ���./���. ��������� ���
// unsigned short LastDay - ���������� ��������� ���� � ��������� ������
//---------------------------------------------------------------------------

#ifndef MonthPickerH
#define MonthPickerH
//---------------------------------------------------------------------------
#include <SysUtils.hpp>
#include <Controls.hpp>
#include <Classes.hpp>
#include <Forms.hpp>
#include <Buttons.hpp>

class PACKAGE TMonthPicker : public TPanel
{
protected:
    void __fastcall SetFEnabled(bool Value);
    bool FEnabled;
    unsigned short FMonth;
    unsigned short FYear;
    TDateTime FMinDate;
    TDateTime FMaxDate;

private:
    bool __fastcall IsNumber(String str);
    void __fastcall MonthButtonClick(TObject *Sender);
    void __fastcall SpeedButtonClick(TObject *Sender);
    void __fastcall EditKeyPress(TObject *Sender, char &Key);
    void __fastcall EditChange(TObject *Sender);
    bool __fastcall CheckBounds(TDateTime dt);
    inline bool __fastcall CheckBoundMin(TDateTime dt);
    inline bool __fastcall CheckBoundMax(TDateTime dt);

    void __fastcall UpdateControl();
    void __fastcall FixDate();
    void __fastcall FixMonth();

    //__fastcall unsigned short GetMonthByDateTime(TDateTime dt);
    TPanel* panel;      // ������, �� ������� ����������� ��������� ����������
    TEdit* edit;        // ���� ���
    TSpeedButton* button1;  // ������ ���������� ����
    TSpeedButton* button2;  // ������ ���������� ����
    TShape* shape;          // ������� ����� ��������� ������ ���� � ��������
    TSpeedButton* btnMonthsList[12];    // ������ �������
    TLabel* labelColorLines[12];        // ������� ������ �������
    TNotifyEvent FOnChange;

__published:
    __property TNotifyEvent OnChange = {read=FOnChange, write=FOnChange};
    __property bool Enabled = {read=FEnabled, write=SetFEnabled, default=true};
    __property TDateTime Date = {read=GetDate, write=SetDate};
    __property unsigned short Month = {read=GetMonth, write=SetMonth};
    __property unsigned short Year = {read=GetYear, write=SetYear};
    __property unsigned short LastDay = {read=GetLastDay};
    __property TDateTime MinDate = {read=FMinDate, write=SetMinDate};
    __property TDateTime MaxDate = {read=FMaxDate, write=SetMaxDate};


public:
    __fastcall TMonthPicker(TComponent* Owner);
    __fastcall virtual ~TMonthPicker(void);
    __fastcall Create(TWinControl *Owner);
    void __fastcall SetYear(unsigned short year);
    void __fastcall SetYear(TDateTime dt);
    void __fastcall SetMonth(unsigned short month);
    void __fastcall SetMonth(TDateTime dt);
    void __fastcall SetDate(TDateTime dt);
    TDateTime __fastcall GetDate();
    String __fastcall GetDate(String format);
    unsigned short __fastcall GetYear();
    unsigned short __fastcall GetMonth();
    unsigned short __fastcall GetLastDay();
    void __fastcall SetMinDate(TDateTime dt);
    void __fastcall SetMaxDate(TDateTime dt);

};

//---------------------------------------------------------------------------
#endif
 