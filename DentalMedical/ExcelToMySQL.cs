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


using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DentalMedical
{
    class ExcelToMySQL
    {
        Excel.Application xlApp;

        public ExcelToMySQL()
        {
            xlApp  = new Excel.Application();   // In case we need to import multiple Excel files only 1 Application for all of them.
        }

        public ExcelCampaign OpenCampaign(string filePath, string password)
        {
            ExcelCampaign newExcel = null; ;
            try
            {
                newExcel = new ExcelCampaign(xlApp, filePath, password);
            }
            catch (Exception e)
            {
                Debug.WriteLine("Couldn't open campaign");
            }

            return newExcel;
        }


        public ArrayList GetAllWorksheetNames(string excelFilePath, string sheetName, string csvExportPath = "", object password = null)
        {
            ArrayList worksheets = new ArrayList();

            Excel.Workbook xlWkbook;
            Excel.Worksheet xlWksht;

            if (password == null)
                password = "";

            try
            {
                xlWkbook = xlApp.Workbooks.Open(excelFilePath, Password: password, ReadOnly: true);
                xlWksht = (Excel.Worksheet)xlWkbook.Worksheets.get_Item(sheetName);
            }
            catch (System.NullReferenceException e)
            {
                Debug.WriteLine("Failed to open");
                throw e;
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                throw new System.Runtime.InteropServices.COMException("Bad password");
            }

            return worksheets;
        }
        /*
         * Read an Excel file (full path) and return the ArrayList describing its headers
         * Prerequisite: 
         *  -  Assume the first header
         *  
         *  @return
         *  - null: Either couldn't open the sheet/workbook, or I skipped more than 20 lines and couldn't find header.
         */
        public ArrayList ExportExcelHeaders(string excelFilePath, string sheetName, string csvExportPath = "", object password = null)
        {
            ArrayList headers = new ArrayList();
            
            
            // Skip any unnecessary space at the top.
            //Excel.Range xlCell = xlWksht.Range["A1"];
            /**
            int skipped = 1;
            while (xlCell.Value != "First Name" || xlCell.Value != "LName")
            {
                //xlCell = xlWksht.Range["A1"];
                skipped++;
                if (skipped > 20)
                    return null;
            }
            MessageBox.Show("Lines skipped" + skipped);
            **/
            // FName, LName, BirthDate, Email, HPhone, MPHone, Last Visit, Appt Date, Appt Time
            
            // TODO: Make this into a while loop
            
            //for (int i = 1; i < 17; i++)
            //{
            //    MessageBox.Show("Adding column: " + xlWksht.Cells[skipped, i].Value);
            //    headers.Add(xlWksht.Cells[skipped, i].Value);
            //} 

            //xlWksht.SaveAs(csvExportPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSVWindows, Type.Missing, Type.Missing, true, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
            
            //xlWkbook.Close(SaveChanges: true);
            return headers;

            // TODO: Error catching in case we can't find the headers and end up with all empties. None should be empty.
        }

        public void ExportExcel(string xlFilePath, string xlPassword)
        {

        }
        /*
         * Read a csv and use an Arraylist of CSV headers to read in each column.
         */
        public void ImportCSVToMySQL(string csvFilePath, ArrayList csvHeaders)
        {
            
        }
        
        public void RemovePassword(string filePath, string password)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlCBPTemplate = xlApp.Workbooks.Open(filePath, Password: password);
            xlCBPTemplate.Password = "";
            xlCBPTemplate.Save();
            xlCBPTemplate.Close();
            xlApp.Quit();
        }
        public void CloseExcel()
        {
            xlApp.Quit();
        }
        
    }
}
