using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using Excel = Microsoft.Office.Interop.Excel;


namespace DataTranfer
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;

            string filePath = @"C:\Users\Nimap\Downloads\backups\Daily sales - Copy.xlsx";
            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);

            try
            {
                string sourceSheetName = "2023";
                Excel.Worksheet sourceSheet = null;

                foreach (Excel.Worksheet sheet in workbook.Sheets)
                {
                    if (sheet.Name == sourceSheetName)
                    {
                        sourceSheet = sheet;
                        break; 
                    }
                }

                sourceSheet.Activate();

                string targetHeader = "Dist";
                Excel.Range headerRow = sourceSheet.Rows[2]; 
                int columnIndex = -1;
                foreach (Excel.Range cell in headerRow.Cells)
                {
                    if (cell.Value2 != null && cell.Value2.ToString() == targetHeader)
                    {
                        columnIndex = cell.Column;
                        break; 
                    }
                }

                Excel.Application targetexcelApp = new Excel.Application();
                targetexcelApp.Visible = true;

                string targetfilePath = @"C:\Users\Nimap\Downloads\backups\Daily Transactions 2023 - Copy.xlsx";
                Excel.Workbook targetworkbook = excelApp.Workbooks.Open(targetfilePath);

                string targetsheetSouthName = "South 23";
                string targetsheetNorthName = "North 23";
                Excel.Worksheet targetSheetSouth = null;
                Excel.Worksheet targetSheetNorth = null;

                foreach (Excel.Worksheet sheet in targetworkbook.Sheets)
                {
                    if (sheet.Name == targetsheetSouthName)
                    {
                        targetSheetSouth = sheet;
                        
                        
                    }
                    else if (sheet.Name == targetsheetNorthName) 
                    {
                        targetSheetNorth = sheet;
                        break;
                        
                    }

                }

                targetSheetSouth.Activate();
                targetSheetNorth.Activate();

     
                int targetSouthRow = targetSheetSouth.UsedRange.Rows.Count + 1;
                int targetNorthRow = targetSheetNorth.UsedRange.Rows.Count + 1;
                int lastRow = sourceSheet.UsedRange.Rows.Count;

                for (int sourceRow = 3; sourceRow <= lastRow; sourceRow++) 
                {
                    string cellValue = sourceSheet.Cells[sourceRow, columnIndex].Value2?.ToString();
                    if (cellValue != null && cellValue.StartsWith("D"))
                    {
                        sourceSheet.Rows[sourceRow].Copy(targetSheetSouth.Rows[targetSouthRow]);
                        targetSouthRow++;
                    }
                    else if (cellValue != null && cellValue.StartsWith("C"))
                    {
                        sourceSheet.Rows[sourceRow].Copy(targetSheetNorth.Rows[targetNorthRow]);
                        targetNorthRow++;
                        
                    }
                }
                targetworkbook.Save();

            }
            finally
            {
                workbook.Close(false, Type.Missing, Type.Missing);
                Marshal.ReleaseComObject(workbook);
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
        }
    }
}


