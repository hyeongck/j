using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
namespace EXCEL_Class
{
    public class Excel_Editing
    {
        public class FCM_Automation_EXCEL : INT
        {
            public Excel.Application xlApp { get; set; }
            public Excel.Workbook xlWorkbook { get; set; }
            public Excel._Worksheet xlWorksheet { get; set; }
            public Excel.Range xlRange { get; set; }

            public Excel.Application xlApp2 { get; set; }
            public Excel.Workbook xlWorkbook2 { get; set; }
            public Excel._Worksheet xlWorksheet2 { get; set; }
            public Excel.Range xlRange2 { get; set; }


            public Excel.Application xlApp3 { get; set; }
            public Excel.Workbook xlWorkbook3 { get; set; }
            public Excel._Worksheet xlWorksheet3 { get; set; }
            public Excel.Range xlRange3 { get; set; }


            public int RowCount { get; set; }
            public int ColumnCount { get; set; }
            public object[,] Data { get; set; }

            public void Open_Session1(string Filepath, string PW, bool Visible)
            {
                xlApp = new Excel.Application();
                Object pwd = PW;
                if (Visible == true) xlApp.Visible = true;
                Object MissingValue = System.Reflection.Missing.Value;
                xlWorkbook = xlApp.Workbooks.Open(Filepath, MissingValue, true, MissingValue, PW);
            }

            public void Open_Session2(string Filepath, string PW, bool Visible)
            {
                xlApp2 = new Excel.Application();
                Object pwd = PW;
                if (Visible == true) xlApp2.Visible = true;
                Object MissingValue = System.Reflection.Missing.Value;
                xlWorkbook2 = xlApp2.Workbooks.Open(Filepath, MissingValue, true, MissingValue, PW);
            }

            public void Clear_Data(string SheetName)
            {
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[SheetName];
                xlRange = xlWorksheet.UsedRange;
                xlRange.Clear();
            }

            public void MakeSheet(string SheetName)
            {
                try
                {
                    xlWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkbook.Sheets["Spec_Sheet_Band"];
                }
                catch
                {
                    int totalSheets = xlApp.ActiveWorkbook.Sheets.Count;
                    Object MissingValue = System.Reflection.Missing.Value;
                    xlWorksheet = (Excel.Worksheet)xlApp.Worksheets.Add();
                    xlWorksheet.Name = "Spec_Sheet_Band";
                    ((Excel.Worksheet)xlApp.ActiveSheet).Move(MissingValue, xlApp.Worksheets[totalSheets + 1]);
                }
            }

            public void MakeSheet_For_Report(string SheetName, int Nb)
            {
                if (Nb == 0)
                {
                    xlWorksheet3 = (Excel.Worksheet)xlApp3.Worksheets.Add();
                    int totalSheets = xlApp3.ActiveWorkbook.Sheets.Count;
                    Object MissingValue = System.Reflection.Missing.Value;

                    xlWorksheet3.Name = SheetName;

                    xlWorksheet3 = (Excel.Worksheet)xlWorkbook3.Worksheets.get_Item(2);

                    xlWorksheet3.Delete();

                    //  ((Excel.Worksheet)xlApp.ActiveSheet).Move(MissingValue, xlApp.Worksheets[totalSheets + 1]);

                }
                else
                {
                    int totalSheets = xlApp3.ActiveWorkbook.Sheets.Count;
                    Object MissingValue = System.Reflection.Missing.Value;
                    xlWorksheet3 = (Excel.Worksheet)xlApp3.Worksheets.Add();
                    xlWorksheet3.Name = SheetName;
                    ((Excel.Worksheet)xlApp3.ActiveSheet).Move(MissingValue, xlApp3.Worksheets[totalSheets + 1]);

                }

            }

            public int Get_Row_Count(string SheetName)
            {
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[SheetName];
                xlRange = xlWorksheet.UsedRange;
                RowCount = xlRange.Rows.Count;

                return RowCount;
            }

            public int Get_Row_Count2(string SheetName)
            {
                xlWorksheet2 = (Excel.Worksheet)xlWorkbook2.Sheets[SheetName];
                xlRange2 = xlWorksheet2.UsedRange;
                RowCount = xlRange2.Rows.Count;

                return RowCount;
            }

            public int Get_Column_Count(string SheetName)
            {
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[SheetName];
                xlRange = xlWorksheet.UsedRange;
                ColumnCount = xlRange.Columns.Count;

                return ColumnCount;
            }

            public int Get_Column_Count2(string SheetName)
            {
                xlWorksheet2 = (Excel.Worksheet)xlWorkbook2.Sheets[SheetName];
                xlRange2 = xlWorksheet2.UsedRange;
                ColumnCount = xlRange2.Columns.Count;

                return ColumnCount;
            }

            public object[,] Read(string SheetName, int Row)
            {
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[SheetName];
                xlRange = xlWorksheet.UsedRange;
                RowCount = xlRange.Rows.Count;
                ColumnCount = xlRange.Columns.Count;

                string columnLetter = ColumnIndexToColumnLetter(ColumnCount); // returns CV
                xlRange = xlWorksheet.get_Range("A" + Row, columnLetter + Row);

                return (object[,])xlRange.Value;
            }

            public object[,] Read2(string SheetName, int Row)
            {
                xlWorksheet2 = (Excel.Worksheet)xlWorkbook2.Sheets[SheetName];
                xlRange2 = xlWorksheet2.UsedRange;
                RowCount = xlRange2.Rows.Count;
                ColumnCount = xlRange.Columns.Count;

                string columnLetter = ColumnIndexToColumnLetter(ColumnCount); // returns CV
                xlRange = xlWorksheet.get_Range("A" + Row, columnLetter + Row);

                return (object[,])xlRange.Value;
            }

            public object Selected_RowandColumn_Read(string SheetName, int Row, int Column)
            {
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[SheetName];
                xlRange = xlWorksheet.UsedRange;
                RowCount = xlRange.Rows.Count;
                ColumnCount = xlRange.Columns.Count;

                string columnLetter = ColumnIndexToColumnLetter(Column); // returns CV
                xlRange = xlWorksheet.get_Range(columnLetter + Row);
                return (object)xlRange.Value;
            }

            public object[,] Read_ColumnbyColumn(string SheetName, int Row, int EndRow, int Column)
            {
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[SheetName];
                xlRange = xlWorksheet.UsedRange;
                RowCount = xlRange.Rows.Count;
                ColumnCount = xlRange.Columns.Count;

                string StartColumn_columnLetter = ColumnIndexToColumnLetter(Column); // returns CV
                xlRange = xlWorksheet.get_Range(StartColumn_columnLetter + Row, StartColumn_columnLetter + EndRow);

                return (object[,])xlRange.Value;
            }

            public object[,] Read2_ColumnbyColumn(string SheetName, int Row, int EndRow, int Column)
            {
                xlWorksheet2 = (Excel.Worksheet)xlWorkbook2.Sheets[SheetName];
                xlRange2 = xlWorksheet2.UsedRange;
                RowCount = xlRange2.Rows.Count;
                ColumnCount = xlRange2.Columns.Count;

                string columnLetter = ColumnIndexToColumnLetter(Column); // returns CV
                xlRange2 = xlWorksheet2.get_Range("A" + Row, columnLetter + EndRow);

                return (object[,])xlRange2.Value;
            }

            public void SelectSheet(string SheetName)
            {

                xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[SheetName];
                xlRange = xlWorksheet.UsedRange;


            }

            public void SelectSheet_For_Report(string SheetName)
            {

                xlWorksheet3 = (Excel.Worksheet)xlWorkbook3.Sheets[SheetName];
                xlRange3 = xlWorksheet3.UsedRange;

                xlWorksheet3.Rows.RowHeight = 17;


            }

            public void SelectSheet2(string SheetName)
            {

                xlWorksheet2 = (Excel.Worksheet)xlWorkbook2.Sheets[SheetName];
                xlRange2 = xlWorksheet2.UsedRange;

            }

            public void Write_Array(int StartRow, int StartColumn, int EndRow, int EndColumn, object[,] Write_Array_Data)
            {
                //try
                //{
                    var StartCell = xlWorksheet.Cells[StartRow, StartColumn];
                    var EndCell = xlWorksheet.Cells[EndRow, EndColumn];
                    var WriteRange = xlWorksheet.Range[StartCell, EndCell];


                    WriteRange.Value2 = Write_Array_Data;
                //}
                //catch (Exception e)
                //{


                //}
            }


            public void Write_Array(int StartRow, int StartColumn, long EndRow, int EndColumn, object[,] Write_Array_Data)
            {
                var StartCell = xlWorksheet.Cells[StartRow, StartColumn];
                var EndCell = xlWorksheet.Cells[EndRow, EndColumn];
                var WriteRange = xlWorksheet.Range[StartCell, EndCell];

                try
                {
                    WriteRange.Value2 = Write_Array_Data;
                }
                catch (Exception e)
                {


                }
            }

            public void Write_Array_For_Report(int StartRow, int StartColumn, int EndRow, int EndColumn, object[,] Write_Array_Data)
            {
                var StartCell = xlWorksheet3.Cells[StartRow, StartColumn];
                var EndCell = xlWorksheet3.Cells[EndRow, EndColumn];
                var WriteRange = xlWorksheet3.Range[StartCell, EndCell];

                WriteRange.Value2 = Write_Array_Data;
            }

            public void Write_Array2(int StartRow, int StartColumn, int EndRow, int EndColumn, object[,] Write_Array_Data)
            {
                var StartCell = xlWorksheet2.Cells[StartRow, StartColumn];
                var EndCell = xlWorksheet2.Cells[EndRow, EndColumn];
                var WriteRange = xlWorksheet2.Range[StartCell, EndCell];

                WriteRange.Value2 = Write_Array_Data;

            }

            public void Write_Array_Formula(int StartRow, int StartColumn, int EndRow, int EndColumn, object[,] Write_Array_Data)
            {
                var StartCell = xlWorksheet.Cells[StartRow, StartColumn];
                var EndCell = xlWorksheet.Cells[EndRow, EndColumn];
                var WriteRange = xlWorksheet.Range[StartCell, EndCell];

                WriteRange.Formula = Write_Array_Data;

            }

            public object[,] Read_Range(int StartRow, int StartColumn, int EndRow, int EndColumn, string SheetName)
            {

                string columnLetter = ColumnIndexToColumnLetter(EndColumn); // returns CV
                xlRange = xlWorksheet.get_Range("A" + StartRow, columnLetter + EndRow);


                return (object[,])xlRange.Value;
            }

            public object[,] Read2_Range(int StartRow, int StartColumn, int EndRow, int EndColumn, string SheetName)
            {

                string columnLetter = ColumnIndexToColumnLetter(EndColumn); // returns CV
                xlRange2 = xlWorksheet2.get_Range("A" + StartRow, columnLetter + EndRow);



                return (object[,])xlRange2.Value;
            }

            public void AddValidation(int StartRow, int StartRColumn, int EndRow, int EndColumn, string Validation, string SheetName)
            {

                xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[SheetName];
                xlRange = xlWorksheet.UsedRange;

                string StartRcolumnLetter = ColumnIndexToColumnLetter(StartRColumn); // returns CV
                string EndRcolumnLetter = ColumnIndexToColumnLetter(EndColumn); // returns CV

                xlRange = xlWorksheet.get_Range(StartRcolumnLetter + StartRow, EndRcolumnLetter + EndRow);

                xlRange.Validation.Delete();
                xlRange.Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, Validation, Type.Missing);

                //string columnLetter = ColumnIndexToColumnLetter(Column); // returns CV

                //var Cell = (Excel.Range) xlWorksheet.Cells[Row, Column];
                //Cell.Validation.Delete();
                //Cell.Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, Validation, Type.Missing);

            }

            public void Insert_Image(string Path, float Left, float Top, float Width, float Heigh)
            {

                Object MissingValue = System.Reflection.Missing.Value;
                Excel.Pictures excelPictures = xlWorksheet3.Pictures(Type.Missing) as Excel.Pictures;


                //   excelPictures.Insert(Path);

                // xlApp3.ActiveCell.Offset.
                xlWorksheet3.Shapes.AddPicture(Path, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Left, Top, Width, Heigh);



            }
            public void MakeBook(bool Visible)
            {
                if (xlApp == null) xlApp = new Excel.Application();

                if (Visible == true) xlApp.Visible = true;
                else xlApp.Visible = false;
                Object MissingValue = System.Reflection.Missing.Value;
                xlWorkbook = xlApp.Workbooks.Add(MissingValue);
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);
                xlApp.DisplayAlerts = true;
            }

            public string MakeBook_For_Report(bool Visible)
            {
                if (xlApp3 == null) xlApp3 = new Excel.Application();

                if (Visible == true) xlApp3.Visible = true;
                else xlApp3.Visible = false;
                Object MissingValue = System.Reflection.Missing.Value;
                xlWorkbook3 = xlApp3.Workbooks.Add(MissingValue);
                xlWorksheet3 = (Excel.Worksheet)xlWorkbook3.Worksheets.get_Item(1);

                xlWorksheet3.Delete();
                xlWorksheet3 = (Excel.Worksheet)xlWorkbook3.Worksheets.get_Item(1);
                xlWorksheet3.Delete();




                return xlWorkbook3.Name;
            }

            public void Interior(string SheetName, int Row, int Column)
            {

                xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[SheetName];
                xlRange = xlWorksheet.UsedRange;

                string columnLetter = ColumnIndexToColumnLetter(Column); // returns CV
                xlRange = xlWorksheet.get_Range(columnLetter + Row);
                xlRange.Interior.ColorIndex = 3;



            }

            public void SaveAs(string File)
            {

                xlWorkbook.SaveAs(File, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                                false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            }

            public void Close()
            {


                xlWorkbook.Close();
            }

            public void Close2()
            {
                xlWorkbook2.Close();
            }

        }

        public class GETSPEC : INT
        {
            public Excel.Application xlApp { get; set; }
            public Excel.Workbook xlWorkbook { get; set; }
            public Excel._Worksheet xlWorksheet { get; set; }
            public Excel.Range xlRange { get; set; }

            public Excel.Application xlApp2 { get; set; }
            public Excel.Workbook xlWorkbook2 { get; set; }
            public Excel._Worksheet xlWorksheet2 { get; set; }
            public Excel.Range xlRange2 { get; set; }

            public Excel.Application xlApp3 { get; set; }
            public Excel.Workbook xlWorkbook3 { get; set; }
            public Excel._Worksheet xlWorksheet3 { get; set; }
            public Excel.Range xlRange3 { get; set; }

            public int RowCount { get; set; }
            public int ColumnCount { get; set; }
            public object[,] Data { get; set; }

            public void Open_Session1(string Filepath, string PW, bool Visible)
            {
                xlApp = new Excel.Application();
                Object pwd = PW;
                if (Visible == true) xlApp.Visible = true;
                Object MissingValue = System.Reflection.Missing.Value;
                xlWorkbook = xlApp.Workbooks.Open(Filepath, MissingValue, true, MissingValue, PW);
            }

            public void Open_Session2(string Filepath, string PW, bool Visible)
            {
                xlApp2 = new Excel.Application();
                Object pwd = PW;
                if (Visible == true) xlApp2.Visible = true;
                Object MissingValue = System.Reflection.Missing.Value;
                xlWorkbook2 = xlApp2.Workbooks.Open(Filepath, MissingValue, true, MissingValue, PW);
            }

            public void Clear_Data(string SheetName)
            {
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[SheetName];
                xlRange = xlWorksheet.UsedRange;
                xlRange.Clear();
            }

            public void MakeSheet(string SheetName)
            {
                try
                {
                    xlWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkbook.Sheets["Spec_Sheet_Band"];
                }
                catch
                {
                    int totalSheets = xlApp.ActiveWorkbook.Sheets.Count;
                    Object MissingValue = System.Reflection.Missing.Value;
                    xlWorksheet = (Excel.Worksheet)xlApp.Worksheets.Add();
                    xlWorksheet.Name = "Spec_Sheet_Band";
                    ((Excel.Worksheet)xlApp.ActiveSheet).Move(MissingValue, xlApp.Worksheets[totalSheets + 1]);
                }
            }

            public void MakeSheet_For_Report(string SheetName, int Nb)
            {
                try
                {
                    xlWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkbook.Sheets["Spec_Sheet_Band"];
                }
                catch
                {
                    int totalSheets = xlApp.ActiveWorkbook.Sheets.Count;
                    Object MissingValue = System.Reflection.Missing.Value;
                    xlWorksheet = (Excel.Worksheet)xlApp.Worksheets.Add();
                    xlWorksheet.Name = "Spec_Sheet_Band";
                    ((Excel.Worksheet)xlApp.ActiveSheet).Move(MissingValue, xlApp.Worksheets[totalSheets + 1]);
                }
            }

            public int Get_Row_Count(string SheetName)
            {
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[SheetName];
                xlRange = xlWorksheet.UsedRange;
                RowCount = xlRange.Rows.Count;

                return RowCount;
            }

            public int Get_Row_Count2(string SheetName)
            {
                xlWorksheet2 = (Excel.Worksheet)xlWorkbook2.Sheets[SheetName];
                xlRange2 = xlWorksheet2.UsedRange;
                RowCount = xlRange2.Rows.Count;

                return RowCount;
            }

            public int Get_Column_Count(string SheetName)
            {
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[SheetName];
                xlRange = xlWorksheet.UsedRange;
                ColumnCount = xlRange.Columns.Count;

                return ColumnCount;
            }

            public int Get_Column_Count2(string SheetName)
            {
                xlWorksheet2 = (Excel.Worksheet)xlWorkbook2.Sheets[SheetName];
                xlRange2 = xlWorksheet2.UsedRange;
                ColumnCount = xlRange2.Columns.Count;

                return ColumnCount;
            }

            public object[,] Read(string SheetName, int Row)
            {
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[SheetName];
                xlRange = xlWorksheet.UsedRange;
                RowCount = xlRange.Rows.Count;
                ColumnCount = xlRange.Columns.Count;

                string columnLetter = ColumnIndexToColumnLetter(ColumnCount); // returns CV
                xlRange = xlWorksheet.get_Range("A" + Row, columnLetter + Row);

                return (object[,])xlRange.Value;
            }

            public object[,] Read2(string SheetName, int Row)
            {
                xlWorksheet2 = (Excel.Worksheet)xlWorkbook2.Sheets[SheetName];
                xlRange2 = xlWorksheet2.UsedRange;
                RowCount = xlRange2.Rows.Count;
                ColumnCount = xlRange.Columns.Count;

                string columnLetter = ColumnIndexToColumnLetter(ColumnCount); // returns CV
                xlRange = xlWorksheet.get_Range("A" + Row, columnLetter + Row);

                return (object[,])xlRange.Value;
            }

            public object Selected_RowandColumn_Read(string SheetName, int Row, int Column)
            {
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[SheetName];
                xlRange = xlWorksheet.UsedRange;
                RowCount = xlRange.Rows.Count;
                ColumnCount = xlRange.Columns.Count;

                string columnLetter = ColumnIndexToColumnLetter(Column); // returns CV
                xlRange = xlWorksheet.get_Range(columnLetter + Row);
                return (object)xlRange.Value;
            }

            public object[,] Read_ColumnbyColumn(string SheetName, int Row, int EndRow, int Column)
            {
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[SheetName];
                xlRange = xlWorksheet.UsedRange;
                RowCount = xlRange.Rows.Count;
                ColumnCount = xlRange.Columns.Count;

                string StartColumn_columnLetter = ColumnIndexToColumnLetter(Column); // returns CV
                xlRange = xlWorksheet.get_Range(StartColumn_columnLetter + Row, StartColumn_columnLetter + EndRow);

                return (object[,])xlRange.Value;
            }

            public object[,] Read2_ColumnbyColumn(string SheetName, int Row, int EndRow, int Column)
            {
                xlWorksheet2 = (Excel.Worksheet)xlWorkbook2.Sheets[SheetName];
                xlRange2 = xlWorksheet2.UsedRange;
                RowCount = xlRange2.Rows.Count;
                ColumnCount = xlRange2.Columns.Count;

                string columnLetter = ColumnIndexToColumnLetter(Column); // returns CV
                xlRange2 = xlWorksheet2.get_Range("A" + Row, columnLetter + EndRow);

                return (object[,])xlRange2.Value;
            }

            public void SelectSheet(string SheetName)
            {

                xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[SheetName];
                xlRange = xlWorksheet.UsedRange;

            }

            public void SelectSheet_For_Report(string SheetName)
            {

                xlWorksheet3 = (Excel.Worksheet)xlWorkbook3.Sheets[SheetName];
                xlRange3 = xlWorksheet3.UsedRange;

            }

            public void SelectSheet2(string SheetName)
            {

                xlWorksheet2 = (Excel.Worksheet)xlWorkbook2.Sheets[SheetName];
                xlRange2 = xlWorksheet2.UsedRange;

            }

            public void Write_Array(int StartRow, int StartColumn, int EndRow, int EndColumn, object[,] Write_Array_Data)
            {
                var StartCell = xlWorksheet.Cells[StartRow, StartColumn];
                var EndCell = xlWorksheet.Cells[EndRow, EndColumn];
                var WriteRange = xlWorksheet.Range[StartCell, EndCell];

                WriteRange.Value2 = Write_Array_Data;
            }
            public void Write_Array(int StartRow, int StartColumn, long EndRow, int EndColumn, object[,] Write_Array_Data)
            {
                var StartCell = xlWorksheet.Cells[StartRow, StartColumn];
                var EndCell = xlWorksheet.Cells[EndRow, EndColumn];
                var WriteRange = xlWorksheet.Range[StartCell, EndCell];

                WriteRange.Value2 = Write_Array_Data;
            }

            public void Write_Array_For_Report(int StartRow, int StartColumn, int EndRow, int EndColumn, object[,] Write_Array_Data)
            {
                var StartCell = xlWorksheet3.Cells[StartRow, StartColumn];
                var EndCell = xlWorksheet3.Cells[EndRow, EndColumn];
                var WriteRange = xlWorksheet3.Range[StartCell, EndCell];

                WriteRange.Value2 = Write_Array_Data;
            }

            public void Write_Array2(int StartRow, int StartColumn, int EndRow, int EndColumn, object[,] Write_Array_Data)
            {
                var StartCell = xlWorksheet2.Cells[StartRow, StartColumn];
                var EndCell = xlWorksheet2.Cells[EndRow, EndColumn];
                var WriteRange = xlWorksheet2.Range[StartCell, EndCell];

                WriteRange.Value2 = Write_Array_Data;

            }

            public void Write_Array_Formula(int StartRow, int StartColumn, int EndRow, int EndColumn, object[,] Write_Array_Data)
            {
                var StartCell = xlWorksheet.Cells[StartRow, StartColumn];
                var EndCell = xlWorksheet.Cells[EndRow, EndColumn];
                var WriteRange = xlWorksheet.Range[StartCell, EndCell];

                WriteRange.Formula = Write_Array_Data;

            }

            public object[,] Read_Range(int StartRow, int StartColumn, int EndRow, int EndColumn, string SheetName)
            {

                string columnLetter = ColumnIndexToColumnLetter(EndColumn); // returns CV
                xlRange = xlWorksheet.get_Range("A" + StartRow, columnLetter + EndRow);


                return (object[,])xlRange.Value;
            }

            public object[,] Read2_Range(int StartRow, int StartColumn, int EndRow, int EndColumn, string SheetName)
            {

                string columnLetter = ColumnIndexToColumnLetter(EndColumn); // returns CV
                xlRange2 = xlWorksheet2.get_Range("A" + StartRow, columnLetter + EndRow);



                return (object[,])xlRange2.Value;
            }

            public void AddValidation(int StartRow, int StartRColumn, int EndRow, int EndColumn, string Validation, string SheetName)
            {

                xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[SheetName];
                xlRange = xlWorksheet.UsedRange;

                string StartRcolumnLetter = ColumnIndexToColumnLetter(StartRColumn); // returns CV
                string EndRcolumnLetter = ColumnIndexToColumnLetter(EndColumn); // returns CV

                xlRange = xlWorksheet.get_Range(StartRcolumnLetter + StartRow, EndRcolumnLetter + EndRow);

                xlRange.Validation.Delete();
                xlRange.Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, Validation, Type.Missing);

                //string columnLetter = ColumnIndexToColumnLetter(Column); // returns CV

                //var Cell = (Excel.Range) xlWorksheet.Cells[Row, Column];
                //Cell.Validation.Delete();
                //Cell.Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, Validation, Type.Missing);

            }

            public void Insert_Image(string Path, float Left, float Top, float Width, float Heigh)
            {

                Object MissingValue = System.Reflection.Missing.Value;
                Excel.Pictures excelPictures = xlWorksheet3.Pictures(Type.Missing) as Excel.Pictures;


                excelPictures.Insert(Path);
                excelPictures.ShapeRange.IncrementLeft(200);
                excelPictures.ShapeRange.IncrementTop(50);

            }
            public void MakeBook(bool Visible)
            {
                if (xlApp == null) xlApp = new Excel.Application();

                if (Visible == true) xlApp.Visible = true;
                else xlApp.Visible = false;
                Object MissingValue = System.Reflection.Missing.Value;
                xlWorkbook = xlApp.Workbooks.Add(MissingValue);
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);

            }

            public string MakeBook_For_Report(bool Visible)
            {
                if (xlApp == null) xlApp = new Excel.Application();

                if (Visible == true) xlApp.Visible = true;
                else xlApp.Visible = false;
                Object MissingValue = System.Reflection.Missing.Value;
                xlWorkbook = xlApp.Workbooks.Add(MissingValue);
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);
                return xlWorkbook.Name;
            }

            public void Interior(string SheetName, int Row, int Column)
            {

                xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[SheetName];
                xlRange = xlWorksheet.UsedRange;

                string columnLetter = ColumnIndexToColumnLetter(Column); // returns CV
                xlRange = xlWorksheet.get_Range(columnLetter + Row);
                xlRange.Interior.ColorIndex = 3;



            }

            public void SaveAs(string File)
            {

                xlWorkbook.SaveAs(File, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                                false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            }

            public void Close()
            {
                xlWorkbook.Close();
            }

            public void Close2()
            {
                xlWorkbook2.Close();
            }

        }

        public interface INT
        {
            Excel.Application xlApp { get; set; }
            Excel.Workbook xlWorkbook { get; set; }
            Excel._Worksheet xlWorksheet { get; set; }
            Excel.Range xlRange { get; set; }

            Excel.Application xlApp2 { get; set; }
            Excel.Workbook xlWorkbook2 { get; set; }
            Excel._Worksheet xlWorksheet2 { get; set; }
            Excel.Range xlRange2 { get; set; }

            Excel.Application xlApp3 { get; set; }
            Excel.Workbook xlWorkbook3 { get; set; }
            Excel._Worksheet xlWorksheet3 { get; set; }
            Excel.Range xlRange3 { get; set; }

            int RowCount { get; set; }
            int ColumnCount { get; set; }
            object[,] Data { get; set; }

            void Open_Session1(string Files, string PW, bool Visible);
            void Open_Session2(string Files, string PW, bool Visible);
            void Clear_Data(string SheetName);
            void MakeSheet(string SheetName);
            void MakeSheet_For_Report(string SheetName, int Nb);
            int Get_Row_Count(string SheetName);
            int Get_Row_Count2(string SheetName);
            int Get_Column_Count(string SheetName);
            int Get_Column_Count2(string SheetName);
            object[,] Read(string SheetName, int Row);
            object[,] Read2(string SheetName, int Row);
            object Selected_RowandColumn_Read(string SheetName, int Row, int Column);
            object[,] Read_ColumnbyColumn(string SheetName, int Row, int EndRow, int Column);
            object[,] Read2_ColumnbyColumn(string SheetName, int Row, int EndRow, int Column);
            void SelectSheet(string SheetName);
            void SelectSheet_For_Report(string SheetName);
            void SelectSheet2(string SheetName);
            void Write_Array(int StartRow, int StartColumn, int EndRow, int EndColumn, object[,] Write_Array_Data);
            void Write_Array(int StartRow, int StartColumn, long EndRow, int EndColumn, object[,] Write_Array_Data);
            void Write_Array_For_Report(int StartRow, int StartColumn, int EndRow, int EndColumn, object[,] Write_Array_Data);
            void Write_Array2(int StartRow, int StartColumn, int EndRow, int EndColumn, object[,] Write_Array_Data);
            void Write_Array_Formula(int StartRow, int StartColumn, int EndRow, int EndColumn, object[,] Write_Array_Data);
            object[,] Read_Range(int StartRow, int StartColumn, int EndRow, int EndColumn, string SheetName);
            object[,] Read2_Range(int StartRow, int StartColumn, int EndRow, int EndColumn, string SheetName);
            void AddValidation(int StartRow, int StartRColumn, int EndRow, int EndColumn, string Validation, string SheetName);
            void MakeBook(bool Visible);
            string MakeBook_For_Report(bool Visible);
            void Insert_Image(string Path, float Left, float Top, float Width, float Heigh);
            void Interior(string SheetName, int Row, int Column);
            void SaveAs(string File);
            void Close();
            void Close2();
        }

        public INT Open(string Key)
        {
            INT Int = null;
            switch (Key.ToUpper().Trim())
            {
                case "FCM":
                    Int = new FCM_Automation_EXCEL();
                    break;

                case "GETSPEC":
                    Int = new GETSPEC();
                    break;
            }
            return Int;
        }

        public static string ColumnIndexToColumnLetter(int colIndex)
        {
            int div = colIndex;
            string colLetter = String.Empty;
            int mod = 0;

            while (div > 0)
            {
                mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (int)((div - mod) / 26);
            }
            return colLetter;
        }
    }
}
