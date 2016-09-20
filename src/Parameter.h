/* ����� ������������ ������ xml ��� ���������� ������ ����������
   ���������� � ������ ������ �� ������������.
 */

#ifndef ParameterH
#define ParameterH


#include <vector.h>
#include "ParameterizedText.h"
#include "..\util\MSXMLWorks.h"


class ParamterEditor
{
public:
    setBeginEdit();
};


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


/* ������� ����� ����������������� ���������
 */
class TParamRecord
{
public:
    typedef String (*CalculateFunction)(const String&);
    typedef void (*BeginEditFunction)(const TParamRecord&);

    static TParamRecord* createParameter(const MsxmlWorks &xml, Variant node);
    static void setValueCalculator(const CalculateFunction &calculate);

    virtual ~TParamRecord() {};
    String getType();
    virtual String getName();
    virtual String getValue();
    virtual String getDisplay();
    virtual String getCaption();

    //virtual TStrings getSubItems();


    virtual void setValue(const String& value);
    virtual void setValue(int index);
    virtual void setValue(const TDateTime& dt);
    virtual bool isVisible();
    virtual bool isDeleted();

    AnsiString type;    // ���
    AnsiString name;    // ���������� ��� ��������

protected:
    AnsiString value;   // ����������? �������� ���������
    AnsiString value_src;   // ���������� (��������) �������� ���������
    AnsiString label;   // ������������ ��� ��������
    AnsiString display; // ������������ �������� ���������

    AnsiString dbindex; // ������ ���� ������ ��� �������� ������ �������� (���� � xml src )
    AnsiString src;     // ������ ���� ������ ��� �������� ������ �������� (���� � xml src )

    bool deleteifflg;   // ���� ������� ���� ���� value ��������� ����� ������� deleteifval
    AnsiString deleteifvalue;  // ���� ������� ���� ���� value ��������� ����� ������� deleteifval

    AnsiString visible;         // ����
    AnsiString visibleif;   // �����������
    bool visibleflg;    // ����������� ��������
    AnsiString parent;      // ��� ������������� ��������� (���� �� ����������)


protected:
    TParamRecord* createDefault(const MsxmlWorks &xml, Variant node);
    String calculate(const String& expression);

private:
    static CalculateFunction _calculate;

};


class TListParameter: public TParamRecord
{
typedef std::vector<TParamlistItem>::iterator ListItemIterator;
public:
    virtual String getValue();
    virtual void setValue(int index);
    TListParameter(const MsxmlWorks &xml, Variant node);
    std::vector <TParamlistItem> listitem;   // ������ �������� (��� list � variables)
private:
    String result;  
};

class TStringParameter: public TParamRecord
{
public:
    TStringParameter(const MsxmlWorks &xml, Variant node);
    virtual void setValue(const String& value);

    AnsiString mask;    // ����� �����
};

class TDateTimeParameter: public TParamRecord
{
public:
    TDateTimeParameter(const MsxmlWorks &xml, Variant node);
    virtual void setValue(const TDateTime& dt);

private:
    AnsiString format;  // ������ ������ ������

};

class TIntegerParameter: public TParamRecord
{
public:
    TIntegerParameter(const MsxmlWorks &xml, Variant node);
};

class TSeparatorParameter: public TParamRecord
{
public:
    TSeparatorParameter(const MsxmlWorks &xml, Variant node);
};

class TFloatParameter: public TParamRecord
{
public:
    TFloatParameter(const MsxmlWorks &xml, Variant node);
};

class TVariableParameter: public TParamRecord
{
public:
    TVariableParameter(const MsxmlWorks &xml, Variant node);
};








//---------------------------------------------------------------------------
#endif // ParameterH
