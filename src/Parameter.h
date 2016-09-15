/* ����� ������������ ������ xml ��� ���������� ������ ����������
   ���������� � ������ ������ �� ������������.
 */

#ifndef ParameterH
#define ParameterH


#include <vector.h>
#include "..\util\MSXMLWorks.h"


// ��������� �������� List � ���������� ������������
class TParamlistItem {
public:
    AnsiString value;       // ����������� ��������
    AnsiString label;       // ������������ ��������
    AnsiString result;      // ������������ ��������� (�� ��������� ����� value)
    AnsiString visible;     // ����������� ���� ���������
    AnsiString visibleif;   // �������, ��� ������� ������� ������������
    bool visibleflg;        // ������� ��������� ��������� � ������ visible � visibleif
};


// ��������� ��� �������� ���������� �������
class TParamRecord
{
public:
    static TParamRecord* createParameter(const MsxmlWorks &xml, Variant node);


    virtual void setVisible(bool visible = true) {
        //control->Visible = visible;
    };

    virtual ~TParamRecord() {};

    AnsiString type;    // ���
    AnsiString name;    // ���������� ��� ��������
    AnsiString value;   // ����������? �������� ���������
    AnsiString value_src;   // ���������� (��������) �������� ���������
    AnsiString label;   // ������������ ��� ��������
    AnsiString display; // ������������ �������� ���������
    AnsiString format;  // ������ ������ ������
    AnsiString dbindex; // ������ ���� ������ ��� �������� ������ �������� (���� � xml src )
    AnsiString src;     // ������ ���� ������ ��� �������� ������ �������� (���� � xml src )
    AnsiString visible;         // ����
    bool deleteifflg;   // ���� ������� ���� ���� value ��������� ����� ������� deleteifval
    AnsiString deleteifvalue;  // ���� ������� ���� ���� value ��������� ����� ������� deleteifval
    //std::vector <TParamlistItem> variables;   // ������ ��������� ��������
    //std::vector <TParamlistItem> listitem;   // ������ �������� (��� list � variables)
    AnsiString visibleif;   // �����������
    AnsiString disableif;   // �����������
    AnsiString parent;      // ��� ������������� ��������� (���� �� ����������)

    bool visibleflg;    // ����������� ��������
    TObject *control;

protected:
    TParamRecord* createDefault(const MsxmlWorks &xml, Variant node);
};


class TListParameter: public TParamRecord
{
public:
    TListParameter(const MsxmlWorks &xml, Variant node);
    std::vector <TParamlistItem> listitem;   // ������ �������� (��� list � variables)
    //TComboBox1* control;
};

class TStringParameter: public TParamRecord
{
public:
    TStringParameter(const MsxmlWorks &xml, Variant node);
    AnsiString mask;    // ����� �����
    //TEdit* control;
};

class TDateTimeParameter: public TParamRecord
{
public:
    TDateTimeParameter(const MsxmlWorks &xml, Variant node);
    //TDateTimePicker* control;
};

class TIntegerParameter: public TParamRecord
{
public:
    TIntegerParameter(const MsxmlWorks &xml, Variant node);
    //TDateTimePicker* control;
};

class TSeparatorParameter: public TParamRecord
{
public:
    TSeparatorParameter(const MsxmlWorks &xml, Variant node);
    //TDateTimePicker* control;
};

class TFloatParameter: public TParamRecord
{
public:
    TFloatParameter(const MsxmlWorks &xml, Variant node);
};













//---------------------------------------------------------------------------
#endif // ParameterH
