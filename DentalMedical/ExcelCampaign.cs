using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DentalMedical
{
    public class ExcelCampaign : ExcelHandler
    {
        public string headerFlag;
        public int numberOfColumns;

        public string Title { get; set; }
        public string Month { get; set; }
        public ArrayList monthHeaders { get; set; }
        public ArrayList masterHeaders { get; set; }
        public Worksheet masterSheet { get; set; }
        public Worksheet monthSheet { get; set; }
        
        
        public ExcelCampaign()
        {
            ;
        }
        public ExcelCampaign(Application xlApp, string password, string path, string title, string month) : base(xlApp, path, password)
        {
            Title = title;
            Month = month;
            headerFlag = "First Name";
            numberOfColumns = 13;
            try
            {
                masterSheet = GetSheet("Master");
                monthSheet = GetSheet(Month);
            }
            catch (Exception e) 
            {
                Debug.WriteLine("Failed to load in sheets");
                throw e;
            }
        }


        /**
         * TODO: This is unnecessary and confusing since it does two things at one time. Needs to be split up
         * and the Headers need to get their own individual return functions with error handling */
        public ArrayList[] ExportHeaders(string firstHeader = "First Name", int lastColIndex = 30, string exportPath = "")
        {
            ArrayList[] headers = new ArrayList[2];

            try
            {
                if (masterHeaders == null)
                    masterHeaders = GetHeaders(masterSheet, firstHeader, lastColIndex);
                if (monthHeaders == null)
                    monthHeaders = GetHeaders(monthSheet, firstHeader, lastColIndex);
            }
            catch (IndexOutOfRangeException ex)
            {
                throw ex;
            }
            

            headers[0] = monthHeaders;
            headers[1] = masterHeaders;

            if (exportPath == "")
                return headers;
            else
            {
                // TODO: Verify that this works across all databases
                // Works on: Diablo
                
                
                return headers;
            }
        }

        public void ExportCampaign(string exportPath)
        {
            // Grab the size of the datasets
            int lastMasterRow = GetLastRow("Master");
            int lastMonthRow = GetLastRow("May 2018");

            masterSheet.Range["A1:Z" + lastMasterRow].Replace(",", ""); // It is very important that we not have commas anywhere in our data due to csv comma delimitter
            masterSheet.Select();
            xlWorkbook.SaveAs(string.Format("{0}{1}{2}.csv", exportPath, Title, "Master"), XlFileFormat.xlCSVWindows, XlSaveAsAccessMode.xlNoChange);
            monthSheet.Range["A1:Z" + lastMonthRow].Replace(",", "");
            monthSheet.Select();
            xlWorkbook.SaveAs(string.Format("{0}{1}{2}.csv", exportPath, Title, "Month"), XlFileFormat.xlCSVWindows, XlSaveAsAccessMode.xlNoChange);
        }
        /// <summary>
        /// Given a column header, attempt to 
        /// </summary>
        /// <param name="header"></param>
        /// <returns></returns>
        public int FindMonthColumnIndexByHeader(string header)
        {
            if (monthHeaders == null)
            {
                Debug.WriteLine("WARNING: Did not call ExportHeaders. Initial settings may be incorrect.");
                monthHeaders = GetHeaders(monthSheet, "First Name", 22);
            }
            foreach (string column in monthHeaders)
            {
                if (column != null) // Sometimes the DIAL column doesn't have a header, it's fine.
                    if (column.Contains(header))
                         monthHeaders.IndexOf(column);
            }
            return -1;
        }

        // New Philosophy: We should never delete things without explicit permission.
        public void CleanSheet(string sheetName, string firstCol, string firstColHeader) 
        {
            ;
        }

        /// <summary>
        /// The Master sheet can sometimes contain a Do Not Call Above Line.
        /// </summary>
        /// <returns>The Worksheet row number of the DNC line. -1 if it is not found.</returns>
        public int GetDNCRow()
        {
            Range dncLine = masterSheet.Range["A1", "A10"].Find("DO NOT CALL ABOVE LINE", MatchCase: false);

            if (dncLine == null)
            {
                Debug.WriteLine("WARNING: Could not find DNC line in Master");
                return -1;
            }
            else
            {
                return dncLine.Row;
            }
        }

        
    }
}
