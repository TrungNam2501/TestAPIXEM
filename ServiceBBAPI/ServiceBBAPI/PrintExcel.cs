using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web;
using Microsoft.Office.Interop.Excel;

namespace ServiceBBAPI
{
    public class PrintExcel
    {
        public static void Print_xls_file(string printername, string pathFile, ref bool flag, ref string messager)
        {
            bool flag1 = true;
            string messager1 = "";
            try
            {
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

                FileInfo newFile = new FileInfo(pathFile);

                // Open the Workbook:
                Microsoft.Office.Interop.Excel.Workbook wb = excelApp.Workbooks.Open(newFile.ToString(),
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);


                // Get the first worksheet.
                // (Excel uses base 1 indexing, not base 0.)
                //Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];

                // Print out 1 copy to the default printer:
                //ws.PrintOut(
                //    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                //    printername, Type.Missing, Type.Missing, Type.Missing);

                //In tat ca cac sheet trong file
                wb.Worksheets.PrintOut(Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    printername, Type.Missing, Type.Missing, Type.Missing);

                // Cleanup:
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //Marshal.FinalReleaseComObject(ws);

                wb.Close(false, Type.Missing, Type.Missing);
                Marshal.FinalReleaseComObject(wb);

                excelApp.Quit();
                Marshal.FinalReleaseComObject(excelApp);
            }
            catch (Exception ex)
            {
                flag = false;
                messager = ex.ToString();
            }
        }
    }
}