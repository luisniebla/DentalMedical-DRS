using Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
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
    class ExcelToMySQL : DBConnection
    {
        Excel.Application xlApp = new Excel.Application();
        ExcelCampaign selectedCampaign;

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
        
        public string LoadCSV(string csvPath, string sqlTable, string delimitter = ",")
        {
            string output = "";
            MySqlDataReader reader = QueryDB("LOAD DATA LOW_PRIORITY LOCAL INFILE '" + csvPath  + "' IGNORE INTO TABLE test." + sqlTable + " FIELDS TERMINATED BY '" + delimitter + "' LINES TERMINATED BY '\r\n' IGNORE 1 LINES;");

            while (reader.Read())
            {
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    output += reader.GetString(i) + "\n";
                }
            }

            reader.Close();

            return output;
        }
        public string[] GetMonthHeader()
        {
            ArrayList headers = selectedCampaign.ExportHeaders()[0];
            return ConvertArrayListToStringList(headers);
        }

        public string[] ConvertArrayListToStringList(ArrayList input)
        {
            string[] headerStrings = new string[input.Count];
            for (int i = 0; i < input.Count; i++)
            {
                headerStrings[i] = (string)input[i];
            }

            return headerStrings;
        }
        public string[] GetMasterHeaders()
        {
            ArrayList headers =  selectedCampaign.ExportHeaders()[1];
            return ConvertArrayListToStringList(headers);
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

                IsConnect();
            }
            catch (Exception e)
            {
                return false;
            }
            return true;
        }
        
        
        public bool ConnectToSQL()
        {
            IsConnect();
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
