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
        protected Dictionary<string, Worksheet> dict;

        public ExcelHandler()
        {
            Password = "";
            xlWorkbook = null;
            dict = new Dictionary<string, Worksheet>();
        }

        public ExcelHandler(Application xlApplication, string filePath, string password)
        {
            try
            {
                dict = new Dictionary<string, Worksheet>();
                xlWorkbook = xlApplication.Workbooks.Open(filePath, Password: password, ReadOnly: false);
                xlWorkbook.Unprotect();
                
                foreach (Worksheet worksheet in xlWorkbook.Worksheets)
                {
                    dict.Add(worksheet.Name, worksheet); // Associate the name of every worksheet with its worksheet object
                }
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
        /// Return the first line of the worksheet in the first column that matches the parameter.
        /// <para>
        /// columnAHeaderString: if empty, then use the first non-empty row as the headers. Else, delete everything until columnAHeaderString is found
        /// deletionLimit: If we delete more than this number of rows, return null arraylist
        /// </para>
        /// </summary>
        /// <returns></returns>
        public ArrayList GetHeaders(Worksheet sheet, string columnAHeaderString = "", int lastColIndex = 13 ,int deletionLimit = 20)
        {
            int headerRow = 1;
            
            if (columnAHeaderString == "")
            {
                while (sheet.Range["A" + headerRow].Value == "" && headerRow < deletionLimit)
                {
                    headerRow++;
                }
            }
            else
            {
                while (sheet.Range["A" + headerRow].Value != columnAHeaderString && headerRow < deletionLimit)
                {
                    headerRow++;
                }
            }

            // Don't delete the rows if we exceed the deletion limit.
            // Now that I think about it, deletion isn't really necessary... but I guess we might as well
            if (headerRow < deletionLimit)
            {
                    for (int j = 1; j < headerRow; j++)
                    {
                        sheet.Range["A" + 1].EntireRow.Delete();
                    }

                    return ReadRow(sheet, 1, lastColIndex);
             }
            else
            {
                return null;
            }
        }

        // TODO: Use Sheet naem instead of WOrksheet
        public ArrayList ReadRow(Worksheet sheet, int row, int lastCol)
        {
            ArrayList values = new ArrayList();
            
            for(int colIndex = 1; colIndex <= lastCol; colIndex++)
            {
                if (sheet.Cells[row, colIndex].Value == "")
                    Debug.WriteLine("WARNING: Could not find a column header for colIndex " + colIndex);
                values.Add(sheet.Cells[row, colIndex].Value);
            }
            
            return values;
        }
        /// <summary>
        /// Return the worksheet given a string
        /// </summary
        /// <returns>the Worksheet object. throws Exception if there was an error in getting worksheet</returns>
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
                    throw e;
                }
            }
            
            return dict[sheetName];
        }

        // I want to extend the worksheet class rather than make a new one, but i don't know how or if it's possible
        public int GetLastRow(string sheet)
        {
            // Grab the range for the last used cell
            Range lastRange = GetSheet(sheet).Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);

            // Convert range to its row value
            int lastUsedRow = lastRange.Row;

            return lastUsedRow;
        }
        
        /// <summary>
        /// Close the Excel workbook.
        /// </summary>
        /// <param name="identifier">Give the option  </param>
        public void close(string identifier = "")
        {
            if (identifier == "")
                xlWorkbook.Close(SaveChanges:XlSaveAction.xlDoNotSaveChanges);
            else
                xlWorkbook.Close(SaveChanges: XlSaveAction.xlSaveChanges, Filename: @"U:\Call Campaigns\ZZ_Merge & CBP RawData\CallBackProof Data\060118\PIMA_THC\" + xlWorkbook.Name + identifier + "." + xlWorkbook.FileFormat);
        }
    }
}
