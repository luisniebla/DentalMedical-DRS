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
            catch(Exception e) 
            {
                Debug.WriteLine("Failed to load in sheets");
                throw e;
            }
        }

        public void CleanSheet(string sheetName, string firstCol, string firstColHeader) 
        {

            
            
        }

        public int GetDNCRow(string sheetName)
        {
            return 0;
        }

    }
}
