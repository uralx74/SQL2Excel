/* Класс обеспечивает разбор xml для заполнения списка параметров
   Вычисление в данном классе не производится.
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


// Структура элемента List в параметрах пользователя
class TParamlistItem {
public:
    AnsiString value;       // Фактическое значение
    AnsiString label;       // Отображаемое значение
    AnsiString result;      // Возвращаемый результат (по умолчанию равен value)
    AnsiString visible;     // Безусловный флаг видимости
    AnsiString visibleif;   // Условие, при котором элемент отображается
    bool visibleflg;        // Текущее состояние видимости с учетом visible и visibleif
};


/* Базовый класс пользовательского параметра
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

    AnsiString type;    // Тип
    AnsiString name;    // Внутреннее имя парамера

protected:
    AnsiString value;   // Внутреннее? значение параметра
    AnsiString value_src;   // Внутреннее (исходное) значение параметра
    AnsiString label;   // Отображаемое имя парамера
    AnsiString display; // Отображаемое значение параметра

    AnsiString dbindex; // Индекс базы данных для загрузки списка значений (если в xml src )
    AnsiString src;     // Индекс базы данных для загрузки списка значений (если в xml src )

    bool deleteifflg;   // Флаг удалять блок если value параметра равен значени deleteifval
    AnsiString deleteifvalue;  // Флаг удалять блок если value параметра равен значени deleteifval

    AnsiString visible;         // Флаг
    AnsiString visibleif;   // Зависимость
    bool visibleflg;    // вычисляемый параметр
    AnsiString parent;      // Имя родительского параметра (пока не доработано)


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
    std::vector <TParamlistItem> listitem;   // Список значений (для list и variables)
private:
    String result;  
};

class TStringParameter: public TParamRecord
{
public:
    TStringParameter(const MsxmlWorks &xml, Variant node);
    virtual void setValue(const String& value);

    AnsiString mask;    // Маска ввода
};

class TDateTimeParameter: public TParamRecord
{
public:
    TDateTimeParameter(const MsxmlWorks &xml, Variant node);
    virtual void setValue(const TDateTime& dt);

private:
    AnsiString format;  // Формат вывода данных

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
