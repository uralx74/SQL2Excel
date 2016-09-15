/* Класс обеспечивает разбор xml для заполнения списка параметров
   Вычисление в данном классе не производится.
 */

#ifndef ParameterH
#define ParameterH


#include <vector.h>
#include "..\util\MSXMLWorks.h"


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


// Структура для хранения параметров запроса
class TParamRecord
{
public:
    static TParamRecord* createParameter(const MsxmlWorks &xml, Variant node);


    virtual void setVisible(bool visible = true) {
        //control->Visible = visible;
    };

    virtual ~TParamRecord() {};

    AnsiString type;    // Тип
    AnsiString name;    // Внутреннее имя парамера
    AnsiString value;   // Внутреннее? значение параметра
    AnsiString value_src;   // Внутреннее (исходное) значение параметра
    AnsiString label;   // Отображаемое имя парамера
    AnsiString display; // Отображаемое значение параметра
    AnsiString format;  // Формат вывода данных
    AnsiString dbindex; // Индекс базы данных для загрузки списка значений (если в xml src )
    AnsiString src;     // Индекс базы данных для загрузки списка значений (если в xml src )
    AnsiString visible;         // Флаг
    bool deleteifflg;   // Флаг удалять блок если value параметра равен значени deleteifval
    AnsiString deleteifvalue;  // Флаг удалять блок если value параметра равен значени deleteifval
    //std::vector <TParamlistItem> variables;   // Список возможных значений
    //std::vector <TParamlistItem> listitem;   // Список значений (для list и variables)
    AnsiString visibleif;   // Зависимость
    AnsiString disableif;   // Зависимость
    AnsiString parent;      // Имя родительского параметра (пока не доработано)

    bool visibleflg;    // вычисляемый параметр
    TObject *control;

protected:
    TParamRecord* createDefault(const MsxmlWorks &xml, Variant node);
};


class TListParameter: public TParamRecord
{
public:
    TListParameter(const MsxmlWorks &xml, Variant node);
    std::vector <TParamlistItem> listitem;   // Список значений (для list и variables)
    //TComboBox1* control;
};

class TStringParameter: public TParamRecord
{
public:
    TStringParameter(const MsxmlWorks &xml, Variant node);
    AnsiString mask;    // Маска ввода
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
