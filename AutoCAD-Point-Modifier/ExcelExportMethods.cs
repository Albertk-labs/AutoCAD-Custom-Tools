using System;
using System.Collections.Generic;
using System.IO;
using System.Drawing;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using OfficeOpenXml;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Runtime.InteropServices;

public static class ExcelExportMethods
{
    public static void ExportPointsToExcel(List<(string Label, Point3d Original, Point3d Modified)> points, List<double> distances, Editor ed)
    {
        ed.WriteMessage("\nStarting export to Excel...");

        Excel.Application excelApp = null;
        Excel.Workbook excelWorkbook = null;
        Excel.Worksheet excelWorksheet = null;

        try
        {
            excelApp = new Excel.Application();
            excelWorkbook = excelApp.Workbooks.Add();
            excelWorksheet = excelWorkbook.Worksheets[1];

            // Headers
            excelWorksheet.Cells[1, 1] = "Point Label";
            excelWorksheet.Cells[1, 2] = "Original X";
            excelWorksheet.Cells[1, 3] = "Original Y";
            excelWorksheet.Cells[1, 4] = "Original Z";
            excelWorksheet.Cells[1, 5] = "Modified X";
            excelWorksheet.Cells[1, 6] = "Modified Y";
            excelWorksheet.Cells[1, 7] = "Modified Z";
            excelWorksheet.Cells[1, 8] = "Translation Distance";

            // Data
            for (int i = 0; i < points.Count; i++)
            {
                excelWorksheet.Cells[i + 2, 1] = points[i].Label;
                excelWorksheet.Cells[i + 2, 2] = points[i].Original.X;
                excelWorksheet.Cells[i + 2, 3] = points[i].Original.Y;
                excelWorksheet.Cells[i + 2, 4] = points[i].Original.Z;
                excelWorksheet.Cells[i + 2, 5] = points[i].Modified.X;
                excelWorksheet.Cells[i + 2, 6] = points[i].Modified.Y;
                excelWorksheet.Cells[i + 2, 7] = points[i].Modified.Z;
                excelWorksheet.Cells[i + 2, 8] = distances[i];
            }

            // Format numeric data
            Excel.Range dataRange = excelWorksheet.Range[excelWorksheet.Cells[2, 2], excelWorksheet.Cells[points.Count + 1, 8]];
            try
            {
                dataRange.NumberFormatLocal = "0.000";
            }
            catch (Exception ex)
            {
                ed.WriteMessage($"\nError setting number format: {ex.Message}");
            }

            // Autofit columns
            excelWorksheet.Columns.AutoFit();

            // Save file
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string fileName = $"ExportedPoints_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            string fullPath = Path.Combine(desktopPath, fileName);
            excelWorkbook.SaveAs(fullPath);

            ed.WriteMessage($"\nExcel file saved: {fullPath}");
        }
        catch (Exception ex)
        {
            ed.WriteMessage($"\nAn error occurred while exporting to Excel: {ex.Message}");
            ed.WriteMessage($"\nError details: {ex.StackTrace}");
        }
        finally
        {
            if (excelWorksheet != null) Marshal.ReleaseComObject(excelWorksheet);
            if (excelWorkbook != null)
            {
                excelWorkbook.Close();
                Marshal.ReleaseComObject(excelWorkbook);
            }
            if (excelApp != null)
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
        }
    }
}
