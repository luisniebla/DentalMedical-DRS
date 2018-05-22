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
        Excel.Application xlApp = new Excel.Application();
        ExcelCampaign selectedCampaign;
        CSVToMySQL sqlHandler = new CSVToMySQL();

        public ExcelToMySQL()
        {
            ;
        }

        public bool OpenCampaign(string filePath, string password, string title, string monthSheet)
        {
            // Let's try opening this workbook
            try
            {
                selectedCampaign = new ExcelCampaign(xlApp, filePath, password, title, "May 2018");
            }
            catch (Exception ex)
            {
                Debug.Write("Could not open excel workbook " + title + "\n" + ex.ToString());
                xlApp.Quit();
                return false;
            }

            return true;
        }
        
        /*
         * Read a csv and use an Arraylist of CSV headers to read in each column.
         */
        public bool ExcelToCSVToMySQL(string csvFilePath)
        {
            // Our Excel workbook opened correctly. Let's try exporting to csv
            try
            {
                selectedCampaign.ExportHeaders(csvFilePath);
            }
            catch (Exception ez)
            {
                Debug.Write("Could not write headers " + "\n" + ez.ToString());
                selectedCampaign.close();
                xlApp.Quit();
                return false;
            }

            try
            {
               
                sqlHandler.Initialize();
            }
            catch (Exception e)
            {
                return false;
            }
            return true;
        }


        
        public bool ConnectToSQL()
        {
            sqlHandler.Initialize();
            return true;
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
