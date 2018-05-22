using Microsoft.Office.Interop.Excel;
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

        public ExcelCampaign OpenCampaign(string filePath, string password, string title, string monthSheet)
        {
            ExcelCampaign newExcel = null; ;
            try
            {
                newExcel = new ExcelCampaign(xlApp, filePath, password, title, monthSheet);
            }
            catch (Exception e)
            {
                Debug.WriteLine("Couldn't open campaign");
            }

            return newExcel;
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
