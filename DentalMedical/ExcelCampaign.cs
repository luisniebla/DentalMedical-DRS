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
        public ArrayList[] ExportHeaders(string firstHeader = "First Name", int lastColIndex = 13, string exportPath = "")
        {
            ArrayList[] headers = new ArrayList[2];

            if (masterHeaders == null)
                masterHeaders = GetHeaders(masterSheet, firstHeader, lastColIndex);
            if (monthHeaders == null)
                monthHeaders = GetHeaders(monthSheet, firstHeader, lastColIndex);

            headers[0] = monthHeaders;
            headers[1] = masterHeaders;

            if (exportPath == "")
                return headers;
            else
            {
                // TODO: Verify that this works across all databases
                // Works on: Diablo
                int lastMasterRow = GetLastRow("Master");
                int lastMonthRow = GetLastRow("May 2018");

                masterSheet.Range["A1:Z" + lastMasterRow].Replace(",","");
                masterSheet.Select();
                xlWorkbook.SaveAs(string.Format("{0}{1}{2}.csv", exportPath, Title, "Master"), XlFileFormat.xlCSVWindows, XlSaveAsAccessMode.xlNoChange);
                masterSheet.Range["A1:Z" + lastMonthRow].Replace(",", "");
                monthSheet.Select();
                xlWorkbook.SaveAs(string.Format("{0}{1}{2}.csv", exportPath, Title, "Month"), XlFileFormat.xlCSVWindows, XlSaveAsAccessMode.xlNoChange);
                
                return headers;
            }
            
        }

        // TODO: Use SQL 
        public void CleanSheet(string sheetName, string firstCol, string firstColHeader) 
        {
            Range dncLine = masterSheet.Range["A1", "A10"].Find("DO NOT CALL ABOVE LINE", MatchCase: false);

            while (masterSheet.Range[firstCol + "1"].Value != "First Name")
            {
                masterSheet.Range[firstCol + "1"].EntireRow.Delete();
            }
        }

        public int GetDNCRow(string sheetName)
        {
            return 0;
        }

        
    }
}
