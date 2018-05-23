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
        string sqlTableNameMonth { get; set; }
        string sqlTableNameMaster { get; set; }

        public ExcelToMySQL()
        {
            sqlTableNameMaster = "";
            sqlTableNameMonth = "";
        }

        public bool OpenCampaign(string filePath, string password, string title, string monthSheet)
        {
            // Let's try opening this workbook
            try
            {
                if (selectedCampaign == null)
                    selectedCampaign = new ExcelCampaign(xlApp, filePath, password, title, monthSheet);
                else
                {
                    selectedCampaign.CloseWorkbook();
                    sqlTableNameMonth = "";
                    sqlTableNameMaster = "";
                    selectedCampaign = new ExcelCampaign(xlApp, filePath, password, title, monthSheet);
                }
                    
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
            ArrayList headers = selectedCampaign.ExportHeaders()[1];
            return ConvertArrayListToStringList(headers);
        }


        /// <summary>
        /// Convert the Excel Campaign (Master and Month tabs) to an SQL table using the headers for each sheet
        /// Only creates all strings for the columns
        /// TODO: Create an index after the table is created.
        /// TODO: Error checking if the table was already made
        /// </summary>
        /// <returns>
        /// True if the creation was successful.
        /// </returns>
        public bool CreateCampaignTables()
        {
            sqlTableNameMonth = selectedCampaign.Title + "_" + selectedCampaign.Month;
            sqlTableNameMaster = selectedCampaign.Title + "_" + selectedCampaign.Master;

            try
            {
                if (CreateStringTable(sqlTableNameMonth, GetMonthHeader()) == null || CreateStringTable(sqlTableNameMaster, GetMasterHeaders()) == null)
                {
                    throw new Exception("Could not connect to SQL server to create table");
                }
            }
            catch(Exception e)
            {
                Debug.WriteLine("Could not create Campaign tables in SQL " + e.ToString());
                return false;
            }
            return true;
        }

        
        /*
         * Read a csv and use an Arraylist of CSV headers to read in each column.
         */
         /// <summary>
         /// Conver the campaign to two seperate CSV files, then import those into SQL
         /// </summary>
         /// <param name="csvFilePath"></param>
         /// <returns></returns>
        public bool ExcelToCSVToMySQL(string csvFilePath = "")
        {
            
            // Our Excel workbook opened correctly. Let's try exporting to csv
            try
            {
                if (sqlTableNameMaster == "" || sqlTableNameMonth == "")
                    CreateCampaignTables();

                if (csvFilePath == "")
                    csvFilePath = @"C:\Users\data\Desktop\";


                selectedCampaign.ExportHeaders(csvFilePath);    // Export to csv
                LoadCSV(string.Format("{0}{1}{2}.csv", csvFilePath, selectedCampaign.Title, selectedCampaign.Month), sqlTableNameMonth);    // Import 
                LoadCSV(string.Format("{0}{1}{2}.csv", csvFilePath, selectedCampaign.Title, selectedCampaign.Master), sqlTableNameMaster);    // Import csv
            }
            catch (Exception e)
            {
                Debug.Write("Could not write headers/export/import to sql " + "\n" + e.ToString());
                selectedCampaign.CloseWorkbook();
                xlApp.Quit();
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
