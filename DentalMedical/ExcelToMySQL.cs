using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace DentalMedical
{
    class ExcelToMySQL
    {
        Excel.Application xlApp;

        public ExcelToMySQL()
        {
            xlApp  = new Excel.Application();   // In case we need to import multiple Excel files only 1 Application for all of them.
        }

        /*
         * Read an Excel file (full path) and return the ArrayList describing its headers
         * Prerequisite: 
         *  -  Assume the first header
         *  
         *  @return
         *  - null: Either couldn't open the sheet/workbook, or I skipped more than 20 lines and couldn't find header.
         */
        public ArrayList ExportExcelHeaders(string excelFilePath, int sheetName, string csvExportPath = "", object password = null)
        {
            ArrayList headers = new ArrayList();
            Excel.Workbook xlWkbook;
            Excel.Worksheet xlWksht;

            if (password == null)
                password = "";
            
            try
            {
                xlWkbook = xlApp.Workbooks.Open(excelFilePath, Password: password, ReadOnly: true);
                xlWksht = (Excel.Worksheet) xlWkbook.Worksheets[sheetName];
            }
            catch (System.NullReferenceException e)
            {
                
                throw e; 
            }
            catch(System.Runtime.InteropServices.COMException e)
            {
                throw new System.Runtime.InteropServices.COMException("Bad password");
            }
            
            // Skip any unnecessary space at the top.
            Excel.Range xlCell = xlWksht.Range["A1"];
            int skipped = 1;
            while (xlCell.Value != "First Name" || xlCell.Value != "LName")
            {
                xlCell = xlWksht.Range["A1"];
                skipped++;
                if (skipped > 20)
                    return null;
            }
            MessageBox.Show("Lines skipped" + skipped);
            // FName, LName, BirthDate, Email, HPhone, MPHone, Last Visit, Appt Date, Appt Time
            
            // TODO: Make this into a while loop
            
            //for (int i = 1; i < 17; i++)
            //{
            //    MessageBox.Show("Adding column: " + xlWksht.Cells[skipped, i].Value);
            //    headers.Add(xlWksht.Cells[skipped, i].Value);
            //} 

            var r = xlWksht.Range["A1"].Resize[1, 20];
            var array = r.Value;

            for(int i = 1; i <= 20; i++)
            {
                var text = array[skipped, i] as string;
                MessageBox.Show(text);
            }

            var array2 = Array.CreateInstance(
                   typeof(object),
                   new int[] { 20, 1 },
                   new int[] { 1, 1 }) as object[,];

            for (int i = 1; i <= 20; i++)
            {
                array2[i, 1] = string.Format("Text{0}", i);
            }
            r.Value2 = array2;
            xlWksht.SaveAs(csvExportPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSVWindows, Type.Missing, Type.Missing, true, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);

            xlWkbook.Close(SaveChanges: true);
            return headers;

            // TODO: Error catching in case we can't find the headers and end up with all empties. None should be empty.
        }

        /*
         * Read a csv and use an Arraylist of CSV headers to read in each column.
         */
        public void ImportCSVToMySQL(string csvFilePath, ArrayList csvHeaders)
        {
            
        }

        public void OpenExtractClose(string path)
        {
            Excel.Application 
        }
        public void CloseExcel()
        {
            xlApp.Quit();
        }
        
    }
}
