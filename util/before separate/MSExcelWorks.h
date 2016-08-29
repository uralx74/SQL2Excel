//---------------------------------------------------------------------------
#ifndef MSEXCELWORKS
#define MSEXCELWORKS

/*******************************************************************************
	����� ��� ����� � OLE-������� Excel.Application
    ������ �� 10.11.2014


    �������� ���������� ������ � �����������:
    1. ������� ������ ������ MSExcelWorks � ������� ����� � ���� ��� ������ ������
        MSExcelWorks msexcel;
        workbook = msexcel.OpenExcel();
        Variant worksheet1 = msexcel.GetSheet(workbook, 1);
    2. ���� ���������� �������� ������ ������ � �������, �� ������� ������ AnsiString:
        std::vector<AnsiString> format_body;
    3. ���� ���������� �������� ���� ������, ������, ������������ � �������, �� ������� ���������� Variant
        � ������ CELLFORMAT:
        Variant region_body;
        CELLFORMAT cf_body;
        cf_body.BorderStyle = CELLFORMAT::xlContinuous;
        cf_body.FontStyle = cfHead.FontStyle << CELLFORMAT::fsBold;
        cf_body.bSetFontColor = true;
        cf_body.FontColor = clRed;
        cf_body.bWrapText = false;
    4. ������� ����������-������ Variant ��� ������:
        Variant data_body;
        data_body = CreateVariantArray(1, FieldCount);
    5. ��������� ����������-������ � �������
        data_body.PutElement("Value", i, j);
    6. ��������� ������
        format_body.push_back("@");
        format_body.push_back("0");
        format_body.push_back("��.��.����");
        format_body.push_back("��:��:��");
    7. ������� ������ � ���� Excel
        region_body = msexcel.WriteTable(worksheet1, ArrayDataBody, 4 <����� ������>, 1 <����� �������>, format_body);
    8. ������ ������ �����:
        msexcel.SetRangeFormat(region_body, cf_body);
    9. ���������� ��������:
        msexcel.SetVisibleExcel(true, true);
    10. ��������� ������� ������:
        cf_body.clear();
        VarClear(ArrayDataHead);
        ArrayDataHead = NULL;

    ---
    CopyRange(worksheet, range_body, int Row, int Column, bool flag);
    ��� Row - ������ �� ������ ������ range_body
    Column - ������ �� ������� ������� range_body
    flag - ������� ���������� �� ������ � ��������������� ��������� ������

*******************************************************************************/

#include "system.hpp"
#include <utilcls.h>
#include "Comobj.hpp"
#include "sysutils.hpp"
#include "taskutils.h"


class CellFormat {
public:
    __fastcall CellFormat();
    enum TDataAlignment {daDefault = 0, daTop = 2, daBottom = 4, daLeft = 2, daCenter = 3, daRight=4};
    enum TFontStyle {fsDefault = 0, fsNormal, fsBold, fsItalic, fsUnderline, fsStrikeOut};
    enum TBorderStyle {bsDefault = -1, bsNone = 0, xlContinuous = 1, bsBold = 7, bsDash1 = 2, bsDash2 = 3, bsDash3 = 4};
    enum TBorderLine {xlEdgeLeft = 7, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideHorizontal, xlInsideVertical};
    AnsiString DataFormat;
    TDataAlignment HorizontalAlignment;
    TDataAlignment VerticalAlignment;
    TBorderStyle BorderStyle;
    int FontSize;
    bool bSetFontColor;
    unsigned long FontColor;
    bool bSetBorderColor;
    unsigned long BorderColor;
    bool bSetFillColor;
    unsigned long FillColor;
    Set <char, 0, 9> FontStyle;
    Set <char, 0, 12> BordeLine;
    char ShrinkToFit;
    char bWrapText;

protected:

private:

};

__fastcall CellFormat::CellFormat()
{
    BorderColor = RGB(0,0,0);
    FontColor = BorderColor;
    FontSize = 0;
    DataFormat = "";
    HorizontalAlignment = daDefault;
    VerticalAlignment = daDefault;
    FontStyle.Clear();
    //FontStyle = FontStyle << fsNormal;
    BordeLine = BordeLine << xlEdgeTop << xlEdgeLeft<< xlEdgeBottom << xlEdgeRight << xlInsideHorizontal << xlInsideVertical;
    BorderStyle = bsDefault;
    bSetFontColor = false;
    bSetBorderColor = false;
    bSetFillColor = false;
    ShrinkToFit = -1;
    bWrapText = -1;
}

typedef CellFormat CELLFORMAT;
typedef std::vector<AnsiString> DATAFORMAT;

class MSExcelWorks
{
private:

public:
    enum TExportStatus {ES_ERROR_NOT_ENOUGTH_FIELD = 0, ES_ERROR_RANGE_IS_NOT_SOLID, ES_ERROR_TOO_MUCH_RECORDS};
    //enum TDirection {Down = 1, Up, Left, Right};
    //enum xlBooleane {xlDefault = -1, xlFalse, xlTrue};
    Variant __fastcall OpenExcel(AnsiString TemplateName = "");     //������ Excel
    Variant __fastcall OpenWorksheetFromFile(AnsiString& FileName);
	void __fastcall CloseExcel(HWND hExW);                  // ������� �������� Excel
	void __fastcall CloseExcel();                           // ������� �������� Excel
    void __fastcall CloseWorkbook(Variant& Book);
    void __fastcall SaveAsDocument(Variant& workbook, AnsiString& FileName);
    Variant __fastcall AddSheet(Variant& Book, AnsiString& SheetName, int SheetIndex = -1);
    Variant __fastcall AddBook(AnsiString TemplateName = "");
    void __fastcall SetActiveWorkbook(Variant& Workbook);
    void __fastcall SetActiveWorksheet(Variant& Worksheet);
    void __fastcall SetActiveRange(Variant& Worksheet, int firstRow, int firstCol, int lastRow = 0, int lastCol = 0);
	Variant __fastcall GetActiveSheet();
	Variant __fastcall GetSheet(Variant& Workbook, int SheetIndex = 1);
    Variant __fastcall GetRange(Variant& Worksheet, int firstRow, int firstCol, int countRow = 1, int countCol = 1);
    Variant __fastcall GetRangeByName(Variant& Worksheet, AnsiString& RangeName);
    Variant __fastcall GetRangeFromRange(Variant& range, int firstRow, int firstCol, int countRow = 1, int countCol = 1);
    std::vector<AnsiString> __fastcall GetNamesFromWorkbook(Variant& WorksheetOrWorkbook);
    std::vector<AnsiString> __fastcall GetNamesFromWorksheet(Variant& Worksheet);
    int __fastcall GetRangeRowsCount(Variant& range);
    int __fastcall GetRangeColumnsCount(Variant& range);
    AnsiString __fastcall GetRangeFormat(Variant& range);

    void __fastcall InsertRows(Variant& worksheet, int RowIndex, int RowsCount);
    Variant __fastcall WriteTable(Variant& worksheet, const Variant &ArrayData,  int firstRow, int firstCol, std::vector<AnsiString> *DataFormat = NULL);

    Variant __fastcall WriteTable(Variant& worksheet, const Variant &ArrayData, AnsiString CellName, std::vector<AnsiString> *DataFormat = NULL);
    Variant __fastcall WriteTableToRange(Variant& range, const Variant &ArrayData,  int firstRow, int firstCol, std::vector<AnsiString> *DataFormat = NULL);


	Variant __fastcall WriteToRange(const AnsiString& txt, Variant range, AnsiString format = "");
	//Variant __fastcall WriteToRange(const AnsiString& txt, const AnsiString& sRangeName, AnsiString format = "");
	Variant __fastcall WriteToCell(Variant& worksheet, const AnsiString& txt, int Row, int Col, AnsiString format = "");
    Variant __fastcall WriteToCell(Variant& worksheet, const AnsiString& txt, AnsiString CellName, AnsiString format = "");
    Variant __fastcall WriteFormulaToCell(Variant& wst, const AnsiString& txt, int Row, int Col, bool fBold = false);
    Variant __fastcall WriteFormula(Variant& worksheet, const AnsiString& txt, int Row, int Col, int countRow = 1, int countCol = 1,  bool fBold = false);
    Variant __fastcall MergeCells(Variant& worksheet, int firstRow, int firstCol, int lastRow, int lastCol);

    void __fastcall SetRangeFormat(Variant& range, const CellFormat& cf, int firstRow, int firstCol, int countRow = 1, int countCol = 1);
    void __fastcall SetRangeFormat(Variant& range, const CellFormat& cf);
    void __fastcall SetRangeDataFormat(Variant& range, AnsiString& format);
	void __fastcall RangeFormat(Variant& wst, int firstRow, int CountRow, int firstCol, int lastCol, int Size, int Font_Color, int Inter_Color, bool Bold); // ������������� �����
    void __fastcall ClearFormats(Variant& range);
	void __fastcall DrawBorders(Variant& range, bool r7 = true, bool r8 = true, bool r9 = true, bool r10 = true, bool r11 = true, bool r12 = true); // ������������� ������� ������� ������� ������ ����� ��������� ���������
    void __fastcall SetRangeColumnsFormat(Variant& range, const std::vector<AnsiString> &cf);
	void __fastcall RangeShtrich(Variant& wst, int firstRow, int CountRow, int firstCol, int lastCol, int Shtrich);
	void __fastcall SetColumnsAutofit(Variant& range);
    void __fastcall SetAutoFilter(Variant& range);
	void __fastcall SetRowsAutofit(Variant& range);

    void __fastcall SetRowHeight(Variant& range, int Height);
	//void __fastcall SetRowHeight(Variant& worksheet, int ColumnIndex, int width);
    void __fastcall SetColumnWidth(Variant& worksheet, int ColumnIndex, int width);
    void __fastcall SetColumnWidth(Variant& range, int Width);

	void __fastcall SetVisibleExcel(bool fVisible = true, bool fForeground = true); // ���������� �������� Excel
    void __fastcall SetVisible(Variant Workbook, bool fVisible = true, bool fForeground = true);
	void __fastcall DateTimeCreateDoc(Variant& wst, int Row, int Col);

    Variant __fastcall ReadRange(Variant& worksheet, int firstRow, int firstCol, int lastRow, int lastCol);
    AnsiString __fastcall ReadCell(Variant& worksheet, int Row, int Col); // ������ ������ �� ������ ������� Excell
    Variant __fastcall ReadCellFormula(Variant& worksheet, int Row, int Col); // ������ ������ �� ������ ������� Excell

    std::vector<AnsiString> __fastcall GetDataFormat(const Variant& ArrayData, int RowIndex);//, std::vector<CELLFORMAT> *formats);
    Variant CreateVariantArray(int RowCount, int ColCount);
    void RedimVariantArray(Variant& ArrayData, int RowCount);
    void __fastcall CopyArray(const Variant& SrcArrayData, Variant* ArrayData,  int srcFirstRow, int srcFirstCol, int srcLastRow, int srcLastCol, int dstFirstRow, int dstFirstCol);

    inline int GetRangeFirstRow(Variant range);
    inline int GetRangeFirstColumn(Variant range);

    Variant __fastcall CopyRange(Variant& worksheet, const Variant& range, int Row = 0, int Column = 0, bool fCopyData = true);
    Variant __fastcall CopyRange(Variant& worksheet, AnsiString sRangeName, int Row = 0, int Column = 0, bool fCopyData = true);
    Variant __fastcall CopyRangeTo(Variant& worksheet, const Variant& range, int Row = 1, int Column = 1, bool fCopyData = true);

    void ExportToExcelFields(TOraQuery* QTable, Variant Worksheet);
    Variant ExportToExcelTable(TOraQuery* QTable, Variant Worksheet, String RangeName, bool fUnbounded = true);


    Variant ExcelApp;
    Variant WorkBooks;
};

//----------------------------------------------------------------------------
// ��������� ����� � ����
void __fastcall MSExcelWorks::SaveAsDocument(Variant& workbook, AnsiString& FileName)
{
    try {
        // ���������� ��������� � ����
		workbook.OleProcedure("SaveAs", FileName);
		//return true;
    } catch (...) {
        throw Exception("�� ������� ��������� �������� � ������ " + FileName);
		//return false;
    }
}

//----------------------------------------------------------------------------
// ��������/��������� ����������
void __fastcall MSExcelWorks::SetAutoFilter(Variant& range)
{
    // expression .AutoFilter(Field, Criteria1, Operator, Criteria2, VisibleDropDown)
    range.OleFunction("AutoFilter");
}

//----------------------------------------------------------------------------
// ��������� ������� ������� �� �����������
void __fastcall MSExcelWorks::SetColumnsAutofit(Variant& range)
{
    range.OlePropertyGet("Columns").OleProcedure("AutoFit");
}

//----------------------------------------------------------------------------
//
void __fastcall MSExcelWorks::SetColumnWidth(Variant& worksheet, int ColumnIndex, int width)
{
    worksheet.OlePropertyGet("Columns").OlePropertyGet("Item", ColumnIndex).OlePropertySet("ColumnWidth", width);
}

//----------------------------------------------------------------------------
// ��������� ������ ������� �� �����������
void __fastcall MSExcelWorks::SetRowsAutofit(Variant& range)
{
    range.OlePropertyGet("Rows").OleProcedure("AutoFit");
}

//----------------------------------------------------------------------------
//
void __fastcall MSExcelWorks::SetRowHeight(Variant& range, int Height)
{
    //range.OlePropertyGet("Rows").OlePropertySet("RowWidth", Width);
    range.OlePropertySet("RowHeight", Height);
}

//----------------------------------------------------------------------------
//
void __fastcall MSExcelWorks::SetColumnWidth(Variant& range, int Width)
{
    range.OlePropertySet("ColumnWidth", Width);
}


//----------------------------------------------------------------------------
// ��������� ���������� ����� � Range
int __fastcall MSExcelWorks::GetRangeRowsCount(Variant& range)
{
    return range.OlePropertyGet("Rows").OlePropertyGet("Count");
}

//----------------------------------------------------------------------------
// ��������� ���������� �������� � Range
int __fastcall MSExcelWorks::GetRangeColumnsCount(Variant& range)
{
    return range.OlePropertyGet("Columns").OlePropertyGet("Count");
}

//----------------------------------------------------------------------------
// ���������� ������ Range
AnsiString __fastcall MSExcelWorks::GetRangeFormat(Variant& range)
{
    AnsiString result;
    result = range.OlePropertyGet("NumberFormat");
    return result;
}

//----------------------------------------------------------------------------
// ������� ������ Range
void __fastcall MSExcelWorks::ClearFormats(Variant& range)
{
    range.OleProcedure("ClearFormats");
}


//----------------------------------------------------------------------------
// ���������� ������ � ������� ����� �� �����
std::vector<AnsiString> __fastcall MSExcelWorks::GetNamesFromWorksheet(Variant& Worksheet)
{
    Variant vNames = Worksheet.OlePropertyGet("Names");
    int nNamesCount = vNames.OlePropertyGet("Count");
    std::vector<AnsiString> vFields;
    vFields.reserve(nNamesCount);

    for(int i=1; i < nNamesCount + 1; i++) {
        AnsiString sName = vNames.OleFunction("Item", i).OlePropertyGet("Name");
        int n = sName.Pos("!");
        //AnsiString sRefers = vNames.OleFunction("Item", i).OlePropertyGet("RefersToR1C1");  // ����� ��������� ������� � ���������� ������ R1C1 Range
        sName = sName.SubString(n+1, sName.Length() - n);     // ����� ����� ������ ����� ! (������: ����1!���)
        vFields.push_back(sName);
    }
    return vFields;
}

//----------------------------------------------------------------------------
// ���������� ������ � ������� ����� � �����
std::vector<AnsiString> __fastcall MSExcelWorks::GetNamesFromWorkbook(Variant& Workbook)
{
    Variant vNames = Workbook.OlePropertyGet("Names");
    int nNamesCount = vNames.OlePropertyGet("Count");
    std::vector<AnsiString> vFields;
    vFields.reserve(nNamesCount);

    for(int i=1; i < nNamesCount + 1; i++) {
        AnsiString sName = vNames.OleFunction("Item", i).OlePropertyGet("Name");
        vFields.push_back(sName);
    }
    return vFields;
}





//----------------------------------------------------------------------------
// ��������� Range
Variant __fastcall MSExcelWorks::GetRange(Variant& Worksheet, int firstRow, int firstCol, int countRow, int countCol)
{
    Variant sell_left_top = Worksheet.OlePropertyGet("Cells", firstRow, firstCol);
	Variant sell_right_bottom = Worksheet.OlePropertyGet("Cells", firstRow+countRow-1, firstCol+countCol-1);
	Variant range = Worksheet.OlePropertyGet("Range", sell_left_top, sell_right_bottom);
    return range;
}

//----------------------------------------------------------------------------
// ��������� Range �� �����
Variant __fastcall MSExcelWorks::GetRangeByName(Variant& Worksheet, AnsiString& RangeName)
{
    try {
        //Variant range = Worksheet.OlePropertyGet("Cells", RangeName);       // ������� ����� ��������� �� �����
        Variant range = Worksheet.OlePropertyGet("Range", RangeName);       // ������� ����� ��������� �� �����
        return range;
    } catch (EOleSysError &e) {
        throw Exception("�� ������� ���������� Range �� ����� " + RangeName);
    }

/*    Variant Workbook = Worksheet.OlePropertyGet("Parent");

    Variant vNames = Workbook.OlePropertyGet("Names");
    AnsiString strName;
    //AnsiString strRefersTo, strSheetName;
    //String srCellName("");
    int nNameCount = vNames.OlePropertyGet("Count");

    for(int i=1; i < nNameCount + 1; i++) {
        strName = vNames.OleFunction("Item", i).OlePropertyGet("Name");
        int n = strName.Pos("!");
        strName = strName.SubString(n+1, strName.Length() - n);

         if(strName == RangeName) {
            return vNames.OleFunction("Item", i).OlePropertyGet("RefersToRange");

            //int nRow = vNames.OleFunction("Item", i).OlePropertyGet("RefersToRange").OlePropertyGet("Row");
            //int nCol = vNames.OleFunction("Item", i).OlePropertyGet("RefersToRange").OlePropertyGet("Column");
            //strRefersTo = vNames.OleFunction("Item", i).OlePropertyGet("RefersTo");
            //strSheetName = strRefersTo.SubString(2, strRefersTo.Pos("!") - 2);
            //return WriteTable(worksheet, ArrayData,  nRow, nCol, DataFormat);
        }
    }
    //VarClear(
    return Unassigned; */
}

//----------------------------------------------------------------------------
// ��������� Range ������ ��������� Range
Variant __fastcall MSExcelWorks::GetRangeFromRange(Variant& range, int firstRow, int firstCol, int countRow, int countCol)
{
    Variant Cells = range.OlePropertyGet("Cells");
    Variant sell_left_top = Cells.OlePropertyGet("Item", firstRow, firstCol);
	Variant sell_right_bottom = Cells.OlePropertyGet("Item", firstRow+countRow-1, firstCol+countCol-1);
    Variant Worksheet = range.OlePropertyGet("Worksheet");
    return Worksheet.OlePropertyGet("Range", sell_left_top, sell_right_bottom);
}

//------------------------------------------------------------------------------
// ��������� ������ ����� ���������� �� Range                   !!! �� ���������
Variant __fastcall MSExcelWorks::ReadRange(Variant& worksheet, int firstRow, int firstCol, int lastRow, int lastCol)
{

    Variant ArrayData;
    Variant sell_left_top = worksheet.OlePropertyGet("Cells", firstRow, firstCol);
	Variant sell_right_bottom = worksheet.OlePropertyGet("Cells", lastRow, lastCol);
	Variant range = worksheet.OlePropertyGet("Range", sell_left_top, sell_right_bottom);
    ArrayData = range.OlePropertyGet("Value");

    return ArrayData;
}

//------------------------------------------------------------------------------
// �������� ����� ������� � ������ ������                       !!! �� ��������
void __fastcall MSExcelWorks::CopyArray(const Variant &SrcArrayData, Variant* ArrayData,  int srcFirstRow, int srcFirstCol, int srcLastRow, int srcLastCol, int dstFirstRow, int dstFirstCol)
{
    for (int i = srcFirstRow; i <= srcLastRow; i ++)
        for (int j = srcFirstCol; j <= srcLastCol; j++) {
            ArrayData->PutElement(SrcArrayData.GetElement(i,j), i, j);   // ����������!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        }
}

//----------------------------------------------------------------------------
// ������ ������ ����� � Range
void __fastcall MSExcelWorks::SetRangeDataFormat(Variant& range, AnsiString& format)
{
    if (format != "") {   // "m/d/yyyy" "@" "0.00" "General"
        try {
            range.OlePropertySet("NumberFormat", format);
        }
        catch (...) {
        }
    }
}

//----------------------------------------------------------------------------
// ������ ������ ����� � Range �������� ������������
void __fastcall MSExcelWorks::SetRangeFormat(Variant& range,  const CellFormat& cf, int firstRow, int firstCol, int countRow, int countCol)
{
	// ������������ ��������� range
    //Variant cell_left_top = range.OlePropertyGet("Cells", firstRow, firstCol);
	//Variant cell_right_bottom = range.OlePropertyGet("Cells", firstRow + countRow - 1, firstCol + countCol - 1);
	//Variant range_tmp = range.OlePropertyGet("Range", cell_left_top, cell_right_bottom);

    Variant range_tmp = GetRangeFromRange(range, firstRow, firstCol, countRow, countCol);
 	SetRangeFormat(range_tmp, cf);
}

//----------------------------------------------------------------------------
// ������ ������ ����� � Range
void __fastcall MSExcelWorks::SetRangeFormat(Variant& range, const CellFormat& cf)
{
    if (cf.DataFormat != "") {   // "m/d/yyyy" "@" "0.00" "General"
        try {
            range.OlePropertySet("NumberFormat", cf.DataFormat);
        }
        catch (...) {
        }
    }

    // ��� 2 - �� ������ ����, 3 - �� ������, 4 - �� �������)
    if (cf.HorizontalAlignment)
        range.OlePropertySet ("HorizontalAlignment", cf.HorizontalAlignment);

    if (cf.VerticalAlignment)
        range.OlePropertySet ("VerticalAlignment", cf.VerticalAlignment - 1);

    if (cf.ShrinkToFit > -1) {  // ���������� ������� ������ �� ������ ������
        range.OlePropertySet("ShrinkToFit", cf.ShrinkToFit);
    }

    if (cf.bWrapText > -1) // ���������� �� ������
        range.OlePropertySet("WrapText", cf.bWrapText);


    Variant font = range.OlePropertyGet("Font");
    if (cf.FontStyle.Contains(CellFormat::fsNormal)) {
        font.OlePropertySet("Bold", false);
        font.OlePropertySet("Italic", false);
        font.OlePropertySet("Underline", false);
    }
    else {
        if (cf.FontStyle.Contains(CellFormat::fsBold))
            font.OlePropertySet("Bold", true);
        if (cf.FontStyle.Contains(CellFormat::fsItalic))
            font.OlePropertySet("Italic", true);
        if (cf.FontStyle.Contains(CellFormat::fsUnderline))
            font.OlePropertySet("Underline", true);
    }

    if (cf.FontSize > 0)
        font.OlePropertySet("Size", cf.FontSize);

    if (cf.bSetFontColor)     // ���� ������ � ������
        font.OlePropertySet("Color", cf.FontColor);

    if (cf.bSetFillColor)
        range.OlePropertyGet("Interior").OlePropertySet("Color", cf.FillColor);

    //if (cf.bWrapText)






    // ��������!!!!!
    if (cf.BorderStyle >= 0) {
        Variant borders = range.OlePropertyGet("Borders");
        //borders.OlePropertySet("LineStyle", cf.BorderStyle);
        for (int i = CellFormat::xlEdgeLeft; i <= CellFormat::xlInsideVertical; i++)
        {
            if (cf.BordeLine.Contains(i))
      	        //try {range.OlePropertyGet("Borders", i).OlePropertySet("LineStyle", cf.BorderStyle);} catch(...) {};
      	        range.OlePropertyGet("Borders", i).OlePropertySet("LineStyle", cf.BorderStyle);
        }
    }

}
//----------------------------------------------------------------------------
// ������ ������ �������� � Range
void __fastcall MSExcelWorks::SetRangeColumnsFormat(Variant& range, const std::vector<AnsiString> &df)
{
    int lastCol = GetRangeColumnsCount(range);
	Variant Columns = range.OlePropertyGet("Columns");
    int SizeDF = df.size();

    for (int i = 0; i < lastCol && i < SizeDF; i++) {         // ��� ������� ������� �� ������ ������
        SetRangeDataFormat(Columns.OlePropertyGet("Item", i+1), df[i]);
    }
}

//----------------------------------------------------------------------------
// ������� ������ �����
void __fastcall MSExcelWorks::InsertRows(Variant& worksheet, int RowIndex, int RowsCount)
{
    if (RowsCount < 1)
        return;
    Variant Rows = worksheet.OlePropertyGet("Rows");
    //Variant Row = Rows.OlePropertyGet("Item", RowIndex, 5);

    try {
        Variant Row = Rows.OlePropertyGet("Range", IntToStr(RowIndex) + ":" + IntToStr(RowIndex+RowsCount-1));
        Row.OleProcedure("Insert", 0xFFFFEFE7, 0);
    } catch (Exception &e) {
        throw Exception("�� ������� �������� ������ � ��������.\n���������� ����������� ����� " + IntToStr(RowsCount) + ".");
    }

    // OleProcedure("Insert", xlDown, xlFormatFromLeftOrAbove);
    // xlDown = 0xFFFFEFE7,
    // xlToLeft = 0xFFFFEFC1,
    // xlToRight = 0xFFFFEFBF,
    // xlUp = 0xFFFFEFBE
    // xlFormatFromLeftOrAbove = 0
    // xlFormatFromRightOrBelow = 1
}

//----------------------------------------------------------------------------
// ������� ���������� ������� � ��������� �� ����� �������
Variant __fastcall MSExcelWorks::WriteTableToRange(Variant& range, const Variant &ArrayData,  int firstRow, int firstCol, std::vector<AnsiString> *DataFormat)
{
    if (DataFormat != NULL) {
        SetRangeColumnsFormat(range, *DataFormat);
    }

    range.OlePropertySet("Value", ArrayData);		// ����� ������ � ��������. ��� ����� ���� ����� ��� ������� ���������� ������
    return range;
}

//----------------------------------------------------------------------------
// ������� ���������� ������� � ��������� �� ����� �������
Variant __fastcall MSExcelWorks::WriteTable(Variant& worksheet, const Variant &ArrayData,  int firstRow, int firstCol, std::vector<AnsiString> *DataFormat)
{
    Variant ArrayRowsCount = VarArrayHighBound(ArrayData, 1) - VarArrayLowBound(ArrayData, 1)+1;
    Variant ArrayColsCount = VarArrayHighBound(ArrayData, 2) - VarArrayLowBound(ArrayData, 2)+1;

	int lastRow = firstRow + ArrayRowsCount - 1; // firstRow .. lastRow
	int lastCol = firstCol + ArrayColsCount - 1; // firstCol  .. lastCol

    Variant sell_left_top = worksheet.OlePropertyGet("Cells", firstRow, firstCol);
	Variant sell_right_bottom = worksheet.OlePropertyGet("Cells", lastRow, lastCol);
	Variant range = worksheet.OlePropertyGet("Range", sell_left_top, sell_right_bottom);

    return WriteTableToRange(range, ArrayData, firstRow, firstCol, DataFormat);

/*    if (DataFormat != NULL) {
        SetRangeColumnsFormat(range, *DataFormat);
    }

    range.OlePropertySet("Value", ArrayData);		// ����� ������ � ��������. ��� ����� ���� ����� ��� ������� ���������� ������
    return range; */
}

//----------------------------------------------------------------------------
// ������� ���������� ������� � ��������� (�����������) �� ����� �������
Variant __fastcall MSExcelWorks::WriteTable(Variant& worksheet, const Variant &ArrayData, AnsiString CellName, std::vector<AnsiString> *DataFormat)
{
    Variant vNames = worksheet.OlePropertyGet("Names");

    //int nNameCount = vNames.OlePropertyGet("Count");

    AnsiString strName;

    Variant range = GetRangeByName(worksheet, CellName);

    int nRow = range.OlePropertyGet("Row");
    int nCol = range.OlePropertyGet("Column");

    return WriteTable(worksheet, ArrayData,  nRow, nCol, DataFormat);
}

//------------------------------------------------------------------------------
// ���������� � �������� ������ ������� ���� � �����
void __fastcall MSExcelWorks::DateTimeCreateDoc(Variant& wst, int Row, int Col)
{
	AnsiString txt = "���� �������� ���������: "+DateTimeToStr(Now());
	WriteToCell(wst, txt, Row, Col);
}

//------------------------------------------------------------------------------
// ������� � ������ ������ � �������� �������
Variant __fastcall MSExcelWorks::WriteToRange(const AnsiString &txt, Variant range, AnsiString format)
{
    if (range.IsEmpty())
        return range;
	range.OlePropertySet("Value", txt);
    if (format != "")
        range.OlePropertySet("NumberFormat", format);   // ������������� ������ ������ ��� ������
	return range;
}

/*//------------------------------------------------------------------------------
// ������� � ������ ������ � �������� �������   2
Variant __fastcall MSExcelWorks::WriteToRange(const AnsiString& txt, const AnsiString& sRangeName, AnsiString format)
{
    Variant range = GetRangeByName(sRangeName);
    return WriteToRange(txt, range, format);
}   */

//------------------------------------------------------------------------------
// ������� � ������ ������ � �������� �������
Variant __fastcall MSExcelWorks::WriteToCell(Variant& worksheet, const AnsiString &txt, int Row, int Col, AnsiString format)
{
	Variant range = worksheet.OlePropertyGet("Cells").OlePropertyGet("Item", Row, Col);
    return WriteToRange(txt, range, format);
}

//------------------------------------------------------------------------------
// ������� � ������ � �������� ������ ������ � �������� �������
Variant __fastcall MSExcelWorks::WriteToCell(Variant& worksheet, const AnsiString &txt, AnsiString CellName, AnsiString format)
{
    Variant range = GetRangeByName(worksheet, CellName);

    return WriteToRange(txt, range, format);
}

//------------------------------------------------------------------------------
// ������� ������� � ������
Variant __fastcall MSExcelWorks::WriteFormulaToCell(Variant& wst, const AnsiString &txt, int Row, int Col, bool fBold)
{
	Variant range = wst.OlePropertyGet("Cells").OlePropertyGet("Item", Row, Col);
	range.OlePropertySet("FormulaR1C1", txt);
    range.OlePropertyGet("Font").OlePropertySet("Bold",fBold);

	return range;
}

//------------------------------------------------------------------------------
// ������� ������� � ������
Variant __fastcall MSExcelWorks::WriteFormula(Variant& worksheet, const AnsiString &txt, int Row, int Col, int countRow, int countCol,  bool fBold)
{
    Variant sell_left_top = worksheet.OlePropertyGet("Cells", Row, Col);
	Variant sell_right_bottom = worksheet.OlePropertyGet("Cells", Row+countRow-1, Col+countCol-1);
	Variant range = worksheet.OlePropertyGet("Range", sell_left_top, sell_right_bottom);
	//Variant range = GetRange(wst, Row, Col, Row+countRow-1, Col+countCol-1);
    // wst.OlePropertyGet("Cells").OlePropertyGet("Item", Row, Col);
	range.OlePropertySet("FormulaR1C1", txt);
    range.OlePropertyGet("Font").OlePropertySet("Bold",fBold);

	return range;
}

/*//------------------------------------------------------------------------------
// ������� ������� � ������
Variant __fastcall MSExcelWorks::WriteFormula(Variant wst, int firstRow, int firstCol, int lastRow = 0, int lastCol = 0);
*/

//---------------------------------------------------------------------------
// ��������� ������ �� ������ ������� Excel
AnsiString __fastcall MSExcelWorks::ReadCell(Variant& worksheet, int Row, int Col)
{
    return worksheet.OlePropertyGet("Cells").OlePropertyGet("Item", Row, Col);
}

//---------------------------------------------------------------------------
// ��������� ������ �� ������ ������� Excel
Variant __fastcall MSExcelWorks::ReadCellFormula(Variant& worksheet, int Row, int Col)
{
    return worksheet.OlePropertyGet("Cells").OlePropertyGet("Item", Row, Col).OlePropertyGet("Formula");
}

//------------------------------------------------------------------------------
// ������� ������ ���� varVariant
Variant MSExcelWorks::CreateVariantArray(int RowCount, int ColCount)
{
    int Bounds[4] = {1, RowCount, 1, ColCount};
//    return VarArrayCreate(Bounds, varString);
    return VarArrayCreate(Bounds, 3, varVariant);
}

//------------------------------------------------------------------------------
// Increase the length of the variant array
void MSExcelWorks::RedimVariantArray(Variant &ArrayData, int RowCount)
{
    VarArrayRedim(ArrayData, RowCount);
}

//------------------------------------------------------------------------------
// ���������� ������ ������ � �������
std::vector<AnsiString> __fastcall MSExcelWorks::GetDataFormat(const Variant &ArrayData, int RowIndex)
{
    std::vector<AnsiString> formats;

    int firstCol = VarArrayLowBound(ArrayData, 2);
    int lastCol = VarArrayHighBound(ArrayData, 2);

    formats.reserve(lastCol - firstCol + 1);
    for (int i = firstCol; i <= lastCol; i++) {         // ��� ������� ������� �� ������ ������
        AnsiString s = ArrayData.GetElement(RowIndex, i);
        AnsiString format;

        // ������������� ������ ����� ��� ������
        if ( IsDate(s.c_str()) )
            format = "��.��.����"; //"m/d/yyyy";    //"��.��.����";
        else if ( IsFloat(s.c_str()) )
            format = "0.00";
        else if ( IsInt(s.c_str()) )   // �������, ��� ��� ������� ����� ����� �������� �� 0
             format = "0";
        else
            format = "@";   // "General"

        formats.push_back(format);
    }
    return formats;
}

//------------------------------------------------------------------------------
// ������������� �������� ����� Excel
void __fastcall MSExcelWorks::SetActiveWorkbook(Variant& Workbook)
{
    Workbook.OleProcedure("Activate");
}

//------------------------------------------------------------------------------
// ������������� �������� ���� � ����� Excel
void __fastcall MSExcelWorks::SetActiveWorksheet(Variant& worksheet)
{
    worksheet.OleProcedure("Activate");
}

//------------------------------------------------------------------------------
// ������������� ������ �� ������ ������ Excel
void __fastcall MSExcelWorks::SetActiveRange(Variant& Worksheet, int firstRow, int firstCol, int lastRow, int lastCol)
{
    GetRange(Worksheet, firstRow, firstCol, lastRow, lastCol).OleProcedure("Select");
}

//------------------------------------------------------------------------------
// ����������� ��������� Excel
// � ���������� ������ �� ������
void __fastcall MSExcelWorks::SetVisibleExcel(bool fVisible, bool fForeground)
{
    String ExcelCaption = ExcelApp.OlePropertyGet("Caption");
    HWND hExcelWindow = FindWindow("XLMAIN", ExcelCaption.c_str());

	if(!ExcelApp.IsEmpty()) {
    	ExcelApp.OlePropertySet("Visible", fVisible);
        //SetWindowPos(hExcelWindow, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE || SWP_NOSIZE);
    	// ExcelApp.OlePropertySet("Activate", true);
        ExcelApp.OlePropertySet("DisplayAlerts", true);         // ���������� ��������������

        if (fVisible && fForeground) {
            // ���� �� �������� ����
            ExcelApp.OlePropertySet("UserControl", true);
            //SetForegroundWindow(hExcelWindow);
            //SetWindowPos(hExcelWindow, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE || SWP_NOSIZE);
            //SetWindowPos(hExcelWindow, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE || SWP_NOSIZE);
            SendMessage(hExcelWindow,WM_SYSCOMMAND,SC_MAXIMIZE,0);
        }
    }
}

//------------------------------------------------------------------------------
// ����������� ��������� Excel
// � ���������� ������ �� ������
void __fastcall MSExcelWorks::SetVisible(Variant Workbook, bool fVisible, bool fForeground)
{
    ExcelApp = Workbook.OlePropertyGet("Application");
    String ExcelCaption = ExcelApp.OlePropertyGet("Caption");
    HWND hExcelWindow = FindWindow("XLMAIN", ExcelCaption.c_str());

	if(!ExcelApp.IsEmpty()) {
    	ExcelApp.OlePropertySet("Visible", fVisible);
        //SetWindowPos(hExcelWindow, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE || SWP_NOSIZE);
    	// ExcelApp.OlePropertySet("Activate", true);
        ExcelApp.OlePropertySet("DisplayAlerts", true);         // ���������� ��������������

        if (fVisible && fForeground) {
            // ���� �� �������� ����
            ExcelApp.OlePropertySet("UserControl", true);
            //SetForegroundWindow(hExcelWindow);
            //SetWindowPos(hExcelWindow, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE || SWP_NOSIZE);
            //SetWindowPos(hExcelWindow, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE || SWP_NOSIZE);
            SendMessage(hExcelWindow,WM_SYSCOMMAND,SC_MAXIMIZE,0);
        }
    }
}


//---------------------------------------------------------------------------
//������ Excel � ����������� ������� �����
Variant __fastcall MSExcelWorks::OpenExcel(AnsiString TemplateName)
{
    if(ExcelApp.IsEmpty())
    {
        try {
    	    ExcelApp = CreateOleObject("Excel.Application");
            ExcelApp.OlePropertySet("DisplayAlerts", false);
            WorkBooks = ExcelApp.OlePropertyGet("Workbooks");
            ExcelApp.OlePropertySet("Visible", false);
        } catch (Exception &exception) {
            //Application->MessageBox("�� ������� ������� �������� Excel","��������",MB_ICONSTOP + MB_OK);
            VariantClear(ExcelApp);
            throw Exception("�� ������� ������� �������� Excel");
            //throw Exception(exception);
            //return Unassigned;
        }
  	}
 	return AddBook(TemplateName);
}

//---------------------------------------------------------------------------
// ���������� (������ ��� �������� �� �������) ������� �����
Variant __fastcall MSExcelWorks::AddBook(AnsiString TemplateName)
{
    Variant Book;
    if (TemplateName != "") {   // ������� �� ������ �������
	    try {
            Book = WorkBooks.OlePropertyGet("Open", TemplateName);
  	    } catch (Exception &exception)	{
            throw Exception("������ ��� �������� �����: " + TemplateName + ".");
        }
    }
    else {                      // ������� ������ ��������
        try {
            Book = WorkBooks.OleFunction("Add");
        } catch (Exception &exception) {
            VariantClear(ExcelApp);
            throw Exception("������ �������� ��������� Excel.");
        }
    }
  	return Book;
}

//---------------------------------------------------------------------------
// ���������� �����
Variant __fastcall MSExcelWorks::AddSheet(Variant& Book, AnsiString& SheetName, int SheetIndex)
{
    Variant Sheets = Book.OlePropertyGet("Worksheets");
    Variant Sheet;

    Variant position;
    if (SheetIndex <= 0) {  // ���������� � ����� ����� (�� ���������)
        position = Sheets.OlePropertyGet("Count");
        Variant After = Sheets.OlePropertyGet("Item", position);

        Sheet = Sheets.OleFunction("Add", VT_EMPTY, After);
        //Sheet = Sheets.OleFunction("Add", EmptyParam, After);
    } else if (SheetIndex == 1){    // ���������� � ������ �����
        Sheet = Sheets.OleFunction("Add");
    } else {                        // ���������� � ������� SheetIndex
        position = SheetIndex - 1;
        Variant After = Sheets.OlePropertyGet("Item", position);
        Sheet = Sheets.OleFunction("Add", VT_EMPTY, After);
        //Sheet = Sheets.OleFunction("Add", EmptyParam, After);
    }

    Sheet.OlePropertySet("Name", SheetName);
    return Sheet;
}

//---------------------------------------------------------------------------
// ���������� �������� �������� Excel
Variant __fastcall MSExcelWorks::GetActiveSheet()
{
    if (!VarIsNull(ExcelApp)) {  // VarIsEmpty
        return ExcelApp.OlePropertyGet("ActiveSheet");
    } else {
        return NULL;
    }
}

//---------------------------------------------------------------------------
// �������� �� ������                               // ����������, ����� ��������� ����� �����
Variant __fastcall MSExcelWorks::GetSheet(Variant& Workbook, int SheetIndex)
{
    if (!VarIsNull(ExcelApp)) {  // VarIsEmpty
        Variant Worksheet = Workbook.OlePropertyGet("Worksheets", SheetIndex);
        return  Worksheet;
        //return  WorkSheets.OlePropertyGet("Item", SheetIndex);
    } else {
        return NULL;
    }
}

//----------------------------------------------------------------------------
// ������� �������� Excel
void __fastcall MSExcelWorks::CloseExcel()
{
    if(!ExcelApp.IsEmpty())
    {
/*        try{
            int WorkbooksCount = WorkBooks.OlePropertyGet("Count");
            for (int i = 1; i <= WorkbooksCount; i++)
                ExcelApp.OlePropertyGet("WorkBooks", i).OleProcedure("Close");
                //WorkBooks.OlePropertyGet("WorkBooks", i).OleProcedure("Close");
        } catch(...){ }   */
        

        WorkBooks.OleProcedure("Close");     // ��� ���������� ���������
        ExcelApp.OleProcedure("Quit");

        /*String ExcelCaption = ExcelApp.OlePropertyGet("Caption");
        HWND hExcelWindow = FindWindow("XLMAIN", ExcelCaption.c_str());
        if ( hExcelWindow != NULL)
            SendMessage(hExcelWindow, WM_DESTROY, 0, 0);  */

        VariantClear(WorkBooks);
        VariantClear(ExcelApp);
    }
}

//----------------------------------------------------------------------------
// ������� �������� ������� ����� Excel
void __fastcall MSExcelWorks::CloseWorkbook(Variant& Book)
{
    if(!Book.IsEmpty())
    {
        Book.OleProcedure("Close", false);
        VariantClear(Book);
    }
}

//------------------------------------------------------------------------------
// ������������� ������� ������� ������� ������ ����� ��������� ���������
void __fastcall MSExcelWorks::DrawBorders(Variant& range, bool r7, bool r8, bool r9, bool r10, bool r11, bool r12)
{
	// r7  - ������� �����
  	// r8  - ������� ������
  	// r9  - ������� �����
  	// r10 - ������� ������
  	// r11 - ����� ������
  	// r12 - ����� �������

  	if (r7) range.OlePropertyGet("Borders",7).OlePropertySet("LineStyle", 1);
  	if (r8) range.OlePropertyGet("Borders",8).OlePropertySet("LineStyle", 1);
  	if (r9) range.OlePropertyGet("Borders",9).OlePropertySet("LineStyle", 1);
  	if (r10) range.OlePropertyGet("Borders",10).OlePropertySet("LineStyle", 1);
  	if (r11) try { range.OlePropertyGet("Borders",11).OlePropertySet("LineStyle", 1); } catch(...) {}
  	if (r12) try { range.OlePropertyGet("Borders",12).OlePropertySet("LineStyle", 1);} catch(...) {}
}

//------------------------------------------------------------------------------
// ���������
void __fastcall MSExcelWorks::RangeShtrich(Variant& wst, int firstRow, int firstCol, int CountRow, int lastCol, int Shtrich)
{
  int lastRow = firstRow + CountRow - 1; // ����� �������� ������
  Variant sell_left_top = wst.OlePropertyGet("Cells", firstRow, firstCol),
          sell_right_bottom = wst.OlePropertyGet("Cells", lastRow, lastCol), diap;
  diap = wst.OlePropertyGet("Range", sell_left_top, sell_right_bottom);
  diap.OlePropertyGet("Interior").OlePropertySet("Pattern", Shtrich);
}

//------------------------------------------------------------------------------
// ����������� �����
Variant __fastcall MSExcelWorks::MergeCells(Variant& worksheet, int firstRow, int firstCol, int lastRow, int lastCol)
{
    //int lastRow = firstRow + CountRow - 1; // ����� �������� ������
    Variant sell_left_top = worksheet.OlePropertyGet("Cells", firstRow, firstCol);
    Variant sell_right_bottom = worksheet.OlePropertyGet("Cells", lastRow, lastCol);
    Variant range = worksheet.OlePropertyGet("Range", sell_left_top, sell_right_bottom);
    range.OlePropertySet("MergeCells", true);
    // Across - True, ����� ���������� ������ � ������ ������ ���������� ��������� ��� ��������� �������. �������� �� ��������� False.

    return range;
}

//------------------------------------------------------------------------------
//
Variant __fastcall MSExcelWorks::OpenWorksheetFromFile(AnsiString& FileName)
{
        Variant workbook;
        Variant worksheet;
        workbook = OpenExcel(FileName);
        worksheet = GetSheet(workbook, 1);
        return worksheet;
}

//------------------------------------------------------------------------------
//
inline int MSExcelWorks::GetRangeFirstRow(Variant range)
{
    return range.OlePropertyGet("Row");
}

//------------------------------------------------------------------------------
//
inline int MSExcelWorks::GetRangeFirstColumn(Variant range)
{
    return range.OlePropertyGet("Column");
}

//---------------------------------------------------------------------------
//
Variant __fastcall MSExcelWorks::CopyRange(Variant& worksheet, const Variant& range, int Row, int Column, bool fCopyData)
{
    //Variant worksheet = range.OlePropertyGet("Worksheet");

    int rowsCount = GetRangeRowsCount(range);
    int colsCount = GetRangeColumnsCount(range);

    int rangeFirstRow = GetRangeFirstRow(range);
    int rangeFirstColumn = GetRangeFirstColumn(range);
    int firstRow = rangeFirstRow + rowsCount + Row;
    int firstCol = rangeFirstColumn + Column;
//    int firstRow = rangeFirstRow + rowsCount + Row;
//   int firstCol = GetRangeFirstColumn(range) + Column;

    int lastRow = firstRow + rowsCount-1;
    int lastCol = firstCol + colsCount-1;


    Variant sell_left_top = worksheet.OlePropertyGet("Cells", firstRow, firstCol);
	Variant sell_right_bottom = worksheet.OlePropertyGet("Cells", lastRow, lastCol);
	Variant range_new = worksheet.OlePropertyGet("Range", sell_left_top, sell_right_bottom);

    // ��������! ��������� ������ ����������� ���-�� ����� = Row
    //if (firstRow != rangeFirstRow)
    //if (Row != 0)
    range_new.OleProcedure("Insert", -4121);  // xlDown

    sell_left_top = worksheet.OlePropertyGet("Cells", firstRow, firstCol);
	sell_right_bottom = worksheet.OlePropertyGet("Cells", lastRow, lastCol);
	range_new = worksheet.OlePropertyGet("Range", sell_left_top, sell_right_bottom);

    range.OleProcedure("Copy", range_new);

    if (!fCopyData)
        range_new.OleProcedure("ClearContents");

    return range_new;
}

//---------------------------------------------------------------------------
//
Variant __fastcall MSExcelWorks::CopyRangeTo(Variant& worksheet, const Variant& range, int Row, int Column, bool fCopyData)
{
    //Variant worksheet = range.OlePropertyGet("Worksheet");

    int rowsCount = GetRangeRowsCount(range);
    int colsCount = GetRangeColumnsCount(range);

    //int rangeFirstRow = GetRangeFirstRow(range);
    //int rangeFirstColumn = GetRangeFirstColumn(range);
    int firstRow = Row;
    int firstCol = Column;
    //int firstRow = rangeFirstRow + rowsCount + Row;
    //int firstCol = GetRangeFirstColumn(range) + Column;

    int lastRow = firstRow + rowsCount-1;
    int lastCol = firstCol + colsCount-1;


    Variant sell_left_top = worksheet.OlePropertyGet("Cells", firstRow, firstCol);
	Variant sell_right_bottom = worksheet.OlePropertyGet("Cells", lastRow, lastCol);
	Variant range_new = worksheet.OlePropertyGet("Range", sell_left_top, sell_right_bottom);

    // ��������! ��������� ������ ����������� ���-�� ����� = Row
    //if (firstRow != rangeFirstRow)
    //if (Row != 0)
        range_new.OleProcedure("Insert", -4121);  // xlDown

    sell_left_top = worksheet.OlePropertyGet("Cells", firstRow, firstCol);
	sell_right_bottom = worksheet.OlePropertyGet("Cells", lastRow, lastCol);
	range_new = worksheet.OlePropertyGet("Range", sell_left_top, sell_right_bottom);

    range.OleProcedure("Copy", range_new);

    if (!fCopyData)
        range_new.OleProcedure("ClearContents");

    return range_new;
}

//---------------------------------------------------------------------------
//
Variant __fastcall MSExcelWorks::CopyRange(Variant& worksheet, AnsiString sRangeName, int Row, int Column, bool fCopyData)
{
    Variant range = GetRangeByName(worksheet, sRangeName);
    return CopyRange(worksheet, range, Row, Column, fCopyData);
}



//---------------------------------------------------------------------------
// ��������� ����������� ������� � ������������ ���������
// ���������������� ���������� �� TOraQuery
//void TForm1::ExportToExcel(String Template, TOraQuery* QTable, String RangeName)
Variant MSExcelWorks::ExportToExcelTable(TOraQuery* QTable, Variant Worksheet, String RangeName, bool fUnbounded)
{
    // �������� range, ���� ���������� �������� �������,
    // �������� ������ � ������� �������� � ���� range
    Variant range_body = GetRangeByName(Worksheet, RangeName);
    int RangeColumnsCount = GetRangeColumnsCount(range_body);
    int RangeRowsCount = GetRangeRowsCount(range_body);
    int RangeFirstRow = GetRangeFirstRow(range_body);
    int RangeFirstColumn = GetRangeFirstColumn(range_body);

    int RecordCount = QTable->RecordCount;


    // �������� �� ��, ��� ���������� ����� � ��������� �� ��������� ���-�� ����� � ��������� Range
    if (RecordCount > RangeRowsCount) {
        if (fUnbounded) {
            //InsertRows(Worksheet, RangeFirstRow+1, RangeFirstColumn-1);
            InsertRows(Worksheet, RangeFirstRow+1, RecordCount - RangeRowsCount);
        } else {
            throw Exception("Error. The source dataset contains too much records.");
        }
    }


    // ��������� ������� vector<String>
    // ��� ������ �������� - ������ ������ � range_body?
    // �������� �������� - ��� ���� � TQuery (� ��� ������ � range_body)

    // ���� �� ��������� range_body � ��������� ������ �� Names
    // ��������� ������ ������� ����� (= ������ ����� � range_body)
    std::vector<String> vs;
    vs.reserve(RangeColumnsCount);

    Variant Cells = range_body.OlePropertyGet("Cells");
    for (int i = 1; i <= RangeColumnsCount; i++) {
        Variant Cell = Cells.OlePropertyGet("Item", 1, i);

        String CellName;
        try {
            CellName = Cell.OlePropertyGet("Name").OlePropertyGet("Name");
        } catch (...) {
            throw Exception("Error. The range is not solid. Cell in column " + IntToStr(i) + " is not named."); // �� ��������� ��� ������ � Range
        }

        if (QTable->Fields->FindField(CellName)) {    // ���������, ���������� �� ���� � ���� ������ � ����� ���������
            vs.push_back(CellName);
        } else {
            throw Exception("Error. Not enough field " + CellName + ".");
        }
    }

    // ��������� ������ data_body ���������� �� QTable
    Variant data_body = CreateVariantArray(RecordCount, RangeColumnsCount);
    VarArrayLock(data_body);
    int j = 1;
    for (QTable->First(); !QTable->Eof; QTable->Next() ) {
        for (int i = 1; i <= RangeColumnsCount; i++) {
            data_body.PutElement(QTable->FieldByName(vs[i-1])->AsString, j, i);
        }
        j++;
    }
    VarArrayUnlock(data_body);


    // ������� �� ����, � ������� rPos, cPos
    Variant range = WriteTable(Worksheet, data_body, RangeFirstRow, RangeFirstColumn);

    // ������������ ������
    VarClear(data_body);

    return range;



//try
//{
    /*if (QTable->RecordCount <= 0)
        MessageBoxInf("��� ������!");


    // ��������� ������
    Variant workbook;
    Variant worksheet;
    try {
        workbook = msexcel.OpenExcel(Template);
        worksheet = msexcel.GetSheet(workbook, 1);
    } catch (...) {
        //fDone = true;
    }  */

    //Variant sss;
    //sss = GetRangeByName(Worksheet, RangeName);
    //Variant vNames = sss.OlePropertyGet("Names");
    //int nNamesCount = vNames.OlePropertyGet("Count");
    //std::vector<AnsiString> vFields;
    //vFields.reserve(nNamesCount);
    //AnsiString sName = vNames.OleFunction("Item", 1).OlePropertyGet("Name");

    /*for(int i=1; i < nNamesCount + 1; i++) {
        AnsiString sName = vNames.OleFunction("Item", i).OlePropertyGet("Name");
        int n = sName.Pos("!");  */


    //Variant sell_left_top = Worksheet.OlePropertyGet("Cells", firstRow, firstCol);
    //Variant sell_right_bottom = Worksheet.OlePropertyGet("Cells", firstRow+countRow-1, firstCol+countCol-1);

    //Variant range = Worksheet.OlePropertyGet("Cells", RangeName);       // ������� ����� ��������� �� �����
    //Variant range = Worksheet.OlePropertyGet("Range", sell_left_top, sell_right_bottom);
    //Variant range = Worksheet.OlePropertyGet("Range", RangeName);       // ������� ����� ��������� �� �����


    //Variant Workbook = Worksheet.OlePropertyGet("Parent");
    //Variant Names = Workbook.OlePropertyGet("Names"); // �������� ����� ����� �� �������� range_body
    //int NamesCount = Names.OlePropertyGet("Count");



    //excelApp.OleFunction("Transpose", data_body);

    // ������� ������ �� ���� � �������� range_body
    //msexcel.WriteTableToRange(range_body, data_body, 1, 1);
    //range_body = msexcel.WriteTable(worksheet, data_body, 1, 1, &DataFormat);

    // ����������� ����������� �������!!!!!!!!!!!!!!
    //Variant range_body_ext = msexcel.GetRange(Worksheet, rPos, cPos, RecordCount, FieldCount);
}








//---------------------------------------------------------------------------
// ��������� ����������� ������� � ������������ ���������
// ���������������� ���������� �� TOraQuery
//void TForm1::ExportToExcel(String Template, TOraQuery* QTable, String RangeName)
/* ���������� �� ������� ������ 2016-04-01
Variant MSExcelWorks::ExportToExcelTable(TOraQuery* QTable, Variant Worksheet, String RangeName)
{

    //try
    //{
        QTable->First();

        // �������� range, ���� ���������� �������� �������,
        // �������� ������ � ������� �������� � ���� range
        Variant range_body = GetRangeByName(Worksheet, RangeName);
        int FieldCount = GetRangeColumnsCount(range_body);
        int RecordCount = QTable->RecordCount;
        Variant data_body = CreateVariantArray(RecordCount, FieldCount);


        Variant Workbook = Worksheet.OlePropertyGet("Parent");
        Variant Names = Workbook.OlePropertyGet("Names"); // �������� ����� ����� �� �������� range_body
        int NamesCount = Names.OlePropertyGet("Count");
        Variant Cells = range_body.OlePropertyGet("Cells");
        std::vector<String> vs;


        // ��������� ������� vector<String>
        // ��� ������ �������� - ������ ������ � range_body?
        // �������� �������� - ��� ���� � TQuery (� ��� ������ � range_body)


        // ���� �� ��������� range_body � ��������� ������ �� Names
        // ��������� ������ ������� ����� (= ������ ����� � range_body)
        for (int i = 1; i <= FieldCount; i++) {
            Variant Cell = Cells.OlePropertyGet("Item", 1, i);
            String cell_addr1;
            try {
                cell_addr1 = Cell.OlePropertyGet("Name");
            } catch (...) {
                throw ES_ERROR_RANGE_IS_NOT_SOLID;  // �� ��������� ��� ������ � Range
            }
            String sFieldName = "";
            for(int j = 1; j <= NamesCount; j++) {
                Variant Name = Names.OleFunction("Item", j);
                String cell_addr2 = Name.OlePropertyGet("Value");   // ���������, ����������� �� ������ ��������� Range. ������?
                if (cell_addr1 == cell_addr2) {         // ���� ������ ����� ���������
                    sFieldName = Name.OlePropertyGet("Name");
                    break;
                }
            }

            if (QTable->Fields->FindField(sFieldName)) {    // ���������, ���������� �� ���� � ���� ������ � ����� ���������
                vs.push_back(sFieldName);
            } else {
                throw ES_ERROR_NOT_ENOUGTH_FIELD;
            }
        }









        // ��������� ������ data_body
        VarArrayLock(data_body);
        int j = 1;
        while (!QTable->Eof) {
            for (int i = 1; i <= FieldCount; i++) {
                data_body.PutElement(QTable->FieldByName(vs[i-1])->AsString, j, i);
            }
            QTable->Next();
            j++;
        }
        VarArrayUnlock(data_body);

        // ������� ������ �� ���� � �������� range_body
        //msexcel.WriteTableToRange(range_body, data_body, 1, 1);
        //range_body = msexcel.WriteTable(worksheet, data_body, 1, 1, &DataFormat);

        int rPos = GetRangeFirstRow(range_body);
        int cPos = GetRangeFirstColumn(range_body);

        InsertRows(Worksheet, rPos+1, RecordCount-1);


        // ����������� ����������� �������!!!!!!!!!!!!!!
        //Variant range_body_ext = msexcel.GetRange(Worksheet, rPos, cPos, RecordCount, FieldCount);

        Variant range = WriteTable(Worksheet, data_body, rPos, cPos);

        // ������������ ������
        VarClear(data_body);

        return range;
        //data_head = NULL;


        //range_body_ext.OlePropertySet("Borders") =
        //range_body.OlePropertyGet("Borders");
/*        Variant t = range_body.OlePropertyGet("Borders").OlePropertyGet("LineStyle");

        int k = range_body.OlePropertyGet("Borders").OlePropertyGet("Count");

        for (int i=7; i <=10; i++) {
            Variant ls = range_body.OlePropertyGet("Borders").OlePropertyGet("Item",i).OlePropertyGet("LineStyle");
            //Variant ls = range_body.OlePropertyGet("Borders", i).OlePropertyGet("LineStyle");
            range_body_ext.OlePropertyGet("Borders").OlePropertyGet("Item",i).OlePropertySet("LineStyle", ls);
            //range_body_ext.OlePropertyGet("Borders", i).OlePropertySet("LineStyle", ls);
        }
*/        //range_body_ext.OlePropertyGet("Borders").OlePropertySet("LineStyle", 1);


        //range_body_ext.OlePropertyGet("Borders").OlePropertySet("LineStyle", range_body.OlePropertyGet("Borders").OlePropertyGet("LineStyle"));
        //range_body.OlePropertyGet("Borders").OlePropertySet("LineStyle", range_body_ext.OlePropertyGet("Borders").OlePropertyGet("LineStyle"));

        //range.OlePropertyGet("Borders", i).OlePropertySet("LineStyle", cf.BorderStyle);


    /*// ��������!!!!!
    if (cf.BorderStyle >= 0) {
        Variant borders = range.OlePropertyGet("Borders");
        //borders.OlePropertySet("LineStyle", cf.BorderStyle);
        for (int i = CellFormat::xlEdgeLeft; i <= CellFormat::xlInsideVertical; i++)
        {
            if (cf.BordeLine.Contains(i))
      	        //try {range.OlePropertyGet("Borders", i).OlePropertySet("LineStyle", cf.BorderStyle);} catch(...) {};
      	        range.OlePropertyGet("Borders", i).OlePropertySet("LineStyle", cf.BorderStyle);
        }
    }*/



    /*Variant ArrayRowsCount = VarArrayHighBound(ArrayData, 1) - VarArrayLowBound(ArrayData, 1)+1;
    Variant ArrayColsCount = VarArrayHighBound(ArrayData, 2) - VarArrayLowBound(ArrayData, 2)+1;

	int lastRow = firstRow + ArrayRowsCount - 1; // firstRow .. lastRow
	int lastCol = firstCol + ArrayColsCount - 1; // firstCol  .. lastCol

    Variant sell_left_top = worksheet.OlePropertyGet("Cells", firstRow, firstCol);
	Variant sell_right_bottom = worksheet.OlePropertyGet("Cells", lastRow, lastCol);
	Variant range = worksheet.OlePropertyGet("Range", sell_left_top, sell_right_bottom);
    */

    //Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
    //SkipBlanks:=False, Transpose:=False


    //} catch(...) {
    //}
//}


//---------------------------------------------------------------------------
// �������� �������� ����������� ����������
// �� ��������������� �������� �� TOraQuery
void MSExcelWorks::ExportToExcelFields(TOraQuery* QTable, Variant Worksheet)
{
    MSExcelWorks msexcel;
    try
    {
        // �������� range, ���� ���������� �������� �������,
        // �������� ������ � ������� �������� � ���� range
        //Variant range_body = msexcel.GetRangeByName(Worksheet, RangeName);
        //int FieldCount = msexcel.GetRangeColumnsCount(range_body);

        if ( !(QTable != NULL && QTable->RecordCount > 0) )
            return;

        //QTable->First();      // ���������������� 2016-03-31
        //int FieldCount = QTable->Fields->Count;  // ���������������� 2016-03-31

        Variant Workbook = Worksheet.OlePropertyGet("Parent");
        std::vector<AnsiString> vExcelFields = msexcel.GetNamesFromWorkbook(Workbook);

        for (std::vector<AnsiString>::iterator itExcelField = vExcelFields.begin(); itExcelField < vExcelFields.end(); itExcelField++) {
            TField* pField = QTable->Fields->FindField(*itExcelField);
            if (pField) {
                Variant range = msexcel.GetRangeByName(Worksheet, *itExcelField);
                msexcel.WriteToRange(QTable->FieldByName(*itExcelField)->AsString, range);
            }
        }


       /*
        // ���������������� 2016-03-31
        
        Variant Workbook = Worksheet.OlePropertyGet("Parent");
        Variant Names = Workbook.OlePropertyGet("Names"); // �������� ����� ����� �� �������� range_body
        int NamesCount = Names.OlePropertyGet("Count");

        // ���� �� ��������� range_body � ��������� ������ �� Names
        // ��������� ������ ������� ����� (= ������ ����� � range_body)
        for (int i = 1; i <= NamesCount; i++) {
            Variant Name = Names.OleFunction("Item", i);
            String sFieldName = UpperCase(Name.OlePropertyGet("Name"));

            for(int j = 1; j <= FieldCount; j++) {
                TField* field = QTable->Fields->FieldByNumber(j);
                if (sFieldName == field->DisplayName) {         // ���� ������ ����� ���������
                    Variant range = msexcel.GetRangeByName(Worksheet, sFieldName);
                    msexcel.WriteToRange(QTable->FieldByName(sFieldName)->AsString, range);
                    break;
                }
            }
        }*/
    } catch(...){
    }
}


/*
//---------------------------------------------------------------------------
// ������ ������ �� ������ ������� Excell
AnsiString __fastcall ReadCell(Variant wst, int Row, int Col)
{
  return(Trim(wst.OlePropertyGet("Cells").OlePropertyGet("Item", Row, Col)));
}

//---------------------------------------------------------------------------
// �������� ��������
Variant __fastcall actPage(Variant appl)
{
  return(appl.OlePropertyGet("ActiveSheet"));
}



    //Variant Application = WorksheetOrWorkbook.OlePropertyGet("Application");
    //Variant Parent = WorksheetOrWorkbook.OlePropertyGet("Parent");
    //if (VarType(Parent) == VarType(Application)) {      // ���� �������� Workbook


*/


/*

// �������� ������ ��� ������ ������
excelApp.OlePropertySet("ScreenUpdating", true);

*/

//---------------------------------------------------------------------------
#endif
