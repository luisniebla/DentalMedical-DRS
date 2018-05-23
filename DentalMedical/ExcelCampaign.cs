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
        private string Title { get; set; }
        private string Month { get; set; }
        ArrayList monthHeaders { get; set; }
        ArrayList masterHeaders { get; set; }
        Worksheet masterSheet { get; set; }
        Worksheet monthSheet { get; set; }

        public ExcelCampaign(Application xlApp, string password, string path, string title, string month) : base(xlApp, password, path)
        {
            Title = title;
            Month = month;

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

        public ArrayList[] ExportHeaders(string exportPath = "")
        {
            ArrayList[] headers = new ArrayList[2];

            if (masterHeaders == null)
                masterHeaders = GetHeaders(masterSheet, "First Name");
            if (monthHeaders == null)
                monthHeaders = GetHeaders(masterSheet, "First Name");

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
