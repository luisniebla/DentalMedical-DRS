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
    class ExcelCampaign
    {
        public string Password { get; set; }
        public Workbook xlWorkbook;   // Can only have 1 workbook open at a time for each campaign, but can have multiple sheets
        Dictionary<string, Worksheet> dict = new Dictionary<string, Worksheet>();

        public ExcelCampaign(Application xlApplication, string filePath, string password)
        {
            try
            {
                xlWorkbook = xlApplication.Workbooks.Open(filePath, Password: password, ReadOnly: true);
            }
            catch (System.NullReferenceException e)
            {
                Debug.WriteLine("Failed to open");
                throw e;
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                throw new System.Runtime.InteropServices.COMException("Bad password");
            }
            
        }


        
        /**
         * @return: A dictionary containing the worksheet name, along with the worksheet object. To use: dict["Sheet1"].Name
         * @error: return null if the workbook is not open
         */
        public ArrayList GetWorksheets()
        {
            if (xlWorkbook == null)
                return null;

            ArrayList sheetNames = new ArrayList();

            foreach(Worksheet worksheet in xlWorkbook.Worksheets)
            {
                dict.Add(worksheet.Name,worksheet); // Associate the name of every worksheet with its worksheet object
                sheetNames.Add(worksheet.Name);
            }

            return sheetNames;
        }

        
        public Worksheet GetSheet(string sheetName)
        {
            if (dict == null)
                return null;

            return dict[sheetName];
        }

        public void close()
        {
            xlWorkbook.Close();
        }
    }
}
