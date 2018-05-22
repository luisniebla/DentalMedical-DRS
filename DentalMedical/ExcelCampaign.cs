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
                masterHeaders = GetHeaders(masterSheet, "First Name");
                monthHeaders = GetHeaders(monthSheet, "First Name");
            }
            catch(Exception e) 
            {
                Debug.WriteLine("Failed to load in sheets");
                throw e;
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
