using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Excel=Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace PivotTableInMVC.Utility
{
    public class Export
    {
        //this method will create pivot table in excel file
        public string OfficeDll()
        {
            string filepath = System.Web.HttpContext.Current.Server.MapPath("~/Content/ProductReport.xlsx");
            int rows = 0;
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkBook = excelApp.Workbooks.Open(filepath);
            Excel.Worksheet excelworksheet = excelWorkBook.ActiveSheet;
            Excel.Worksheet sheet2 = excelWorkBook.Sheets.Add();
            try
            {
                sheet2.Name = "Pivot Table";
                excelApp.ActiveWindow.DisplayGridlines = false;
                Excel.Range oRange = excelworksheet.UsedRange;
                Excel.PivotCache oPivotCache = excelWorkBook.PivotCaches().Add(Excel.XlPivotTableSourceType.xlDatabase, oRange);  // Set the Source data range from First sheet
                Excel.PivotCaches pch = excelWorkBook.PivotCaches();
                pch.Add(Excel.XlPivotTableSourceType.xlDatabase, oRange).CreatePivotTable(sheet2.Cells[3, 3], "PivTbl_2", Type.Missing, Type.Missing);// Create Pivot table
                Excel.PivotTable pvt = sheet2.PivotTables("PivTbl_2");
                pvt.ShowDrillIndicators = true;
                pvt.InGridDropZones = false;
                Excel.PivotField fld = ((Excel.PivotField)pvt.PivotFields("CATEGORY"));
                fld.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                fld.set_Subtotals(1, false);

                fld = ((Excel.PivotField)pvt.PivotFields("PLACE"));
                fld.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                fld.set_Subtotals(1, false);

                fld = ((Excel.PivotField)pvt.PivotFields("NAME"));
                fld.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                fld.set_Subtotals(1, false);

                fld = ((Excel.PivotField)pvt.PivotFields("PRICE"));
                fld.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                fld.set_Subtotals(1, false);

                fld = ((Excel.PivotField)pvt.PivotFields("NoOfUnits"));
                fld.Orientation = Excel.XlPivotFieldOrientation.xlDataField;

                sheet2.UsedRange.Columns.AutoFit();
                pvt.ColumnGrand = true;
                pvt.RowGrand = true;
                excelApp.DisplayAlerts = false;
                excelworksheet.Delete();
                sheet2.Activate();
                sheet2.get_Range("B1", "B1").Select();
                excelWorkBook.SaveAs(filepath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                excelApp.DisplayAlerts = false;
                excelWorkBook.Close(0);
                excelApp.Quit();
                Marshal.ReleaseComObject(excelWorkBook);
                Marshal.ReleaseComObject(excelApp);
            }
            catch (Exception ex)
            {
                excelWorkBook.Close(0);
                excelApp.Quit();
                Marshal.ReleaseComObject(excelWorkBook);
                Marshal.ReleaseComObject(excelApp);

                return ex.Message;
            }
            return filepath;
        }


    }    
}