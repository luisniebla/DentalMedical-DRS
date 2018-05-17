using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;

namespace DentalMedical
{
    public abstract class ExcelHandler 
    {
        protected string Password { get; set; }
        public Workbook xlWorkbook;   // Can only have 1 workbook open at a time for each campaign, but can have multiple sheets
        protected Dictionary<string, Worksheet> dict = new Dictionary<string, Worksheet>();

        public ExcelHandler()
        {
            Password = "";
            xlWorkbook = null;
        }

        public ExcelHandler(Application xlApplication, string filePath, string password)
        {
            try
            {
                xlWorkbook = xlApplication.Workbooks.Open(filePath, Password: password, ReadOnly: true);
            }
            catch (Exception e)
            {
                Debug.WriteLine("Failed to open workbook");
                throw e;
            }
        }

        /**
         * @return: A dictionary containing the worksheet name, along with the worksheet object. To use: dict["Sheet1"].Name
         * @error: return null if the workbook is not open
         */
        public Dictionary<string, Worksheet> GetWorksheets()
        {
            if (xlWorkbook == null)
                return null;

            if (dict == null)
            {
                foreach (Worksheet worksheet in xlWorkbook.Worksheets)
                {
                    dict.Add(worksheet.Name, worksheet); // Associate the name of every worksheet with its worksheet object
                    
                }
            }
            
            return dict;
        }

        /// <summary>
        /// Return the worksheet
        /// </summary>
        public Worksheet GetSheet(string sheetName)
        {
            if (dict == null)
            {
                try
                {
                    return GetWorksheets()[sheetName];
                }
                catch (Exception e)
                {
                    Debug.WriteLine(e.ToString());
                    return null;
                }
            }


            return dict[sheetName];
        }

        // I want to extend the worksheet class rather than make a new one, but i don't know how or if it's possible
        public int GetLastRow(string sheet)
        {
            Range lastRange = GetSheet(sheet).Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);

            int lastUsedRow = lastRange.Row;

            return lastUsedRow;
        }

        public void close()
        {
            xlWorkbook.Close();
        }
    }
}
