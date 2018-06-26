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
        Application xlApp = new Application();

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
        
        /// <summary>
        /// Find and grab the header column for a campaign
        /// </summary>
        /// <param name="firstHeader">Defaults to "First Name"</param>
        /// <param name="lastColIndex">The size of the header array</param>
        /// <returns>An ArrayList consisting of the month headers for a campaign</returns>
        public ArrayList GetMonthHeaderList(string firstHeader = "First Name", int lastColIndex = 30)
        {
            if (monthHeaders == null)
            {
                try
                {
                    monthHeaders = GetHeaders(monthSheet, firstHeader, lastColIndex);
                }
                catch (IndexOutOfRangeException e)
                {
                    monthHeaders = null;
                }
            }
                monthHeaders = GetHeaders(monthSheet, firstHeader, lastColIndex);

            return monthHeaders;
        }

        public void cbp()
        {
            /**
            string lastNameCol = "B";
            string firstNameCol = "A";
            string dobCol = "G";
            string updateCol = "S";
            **/
            // Call back proofs don't need all the new insurance data... maybe?
            /**
            for (int i = 2; i < 2618; i++)
            {
                Worksheet newData = bancroft.GetWorksheets()["NewData"];
                string lastName = newData.Range[lastNameCol + i].Value;

                int findMatch = bancroft.FindItemInMonthColumn(lastNameCol, lastName);

                if (findMatch > 0)
                {
                    string firstName = newData.Range[firstNameCol + i].Value;
                    string DOB = newData.Range[dobCol + i].Value.ToString();

                    if (bancroft.monthSheet.Range[dobCol + findMatch].Value.ToString() == DOB && bancroft.monthSheet.Range[firstNameCol + findMatch].Value == firstName)
                    {
                        newData.Range[updateCol + i].Value = "MONTH MATCH";
                    }
                    else
                    {
                        newData.Range[firstNameCol + i].EntireRow.Delete();
                        i = i - 1;
                    }
                }
                else
                {
                    newData.Range[firstNameCol + i].EntireRow.Delete();
                    i = i - 1;
                }
            **/
        }
        
        public string MonthHeadersToString()
        {
            string monthHeaderString = "";

            if (monthHeaders == null)
                monthHeaders = GetMonthHeaderList();


            foreach (object s in monthHeaders)
            {
                if (s == null)
                    monthHeaderString += "|";
                else
                    monthHeaderString += s.ToString() + "|";
            }

            return monthHeaderString;
        }

        public ArrayList GetMasterHeaders(string firstHeader, int lastColIndex)
        {
            if (masterHeaders == null)
            {
                try
                {
                    monthHeaders = GetHeaders(monthSheet, firstHeader, lastColIndex);
                }
                catch (IndexOutOfRangeException e)
                {
                    monthHeaders = null;
                }
            }
            masterHeaders = GetHeaders(monthSheet, firstHeader, lastColIndex);

            return monthHeaders;
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
        /// Given a column header, attempt to get the Excel-based index for it
        /// </summary>
        /// <param name="header"></param>
        /// <returns>The Excel-based index depending on the</returns>
        public int FindMonthColumnIndexByHeader(string header)
        {
            if (monthHeaders == null)
            {
                Debug.WriteLine("WARNING: Did not call ExportHeaders. Initial settings may be incorrect.");
                monthHeaders = GetMonthHeaderList();
            }

            int colIndex = monthHeaders.IndexOf(header);

            if (colIndex == null)
                return 0;
            else
                return colIndex + 1;
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

        /**
         * Assume data stored in NewData very prettily
         */
        public int CallBackProof( )
        {
            return 0;   
        }
        
        /// <summary>
        /// Return the row of the first instance of text in a certain column using Excel Find command.
        /// NOTE THIS IS NOTE MATCHING THE ENTIRE CELL
        /// </summary>
        /// <param name="column">The letter of the column to search</param>
        /// <param name="text">Text to find</param>
        /// <returns>
        /// The row if the text in found within the column
        /// 0 if the text is not found
        /// -1 if there is a duplicate in the data, which we don't know how to handle right now
        /// </returns>
        public int FindItemInMonthColumn(string column, string text)
        {

            Range rng = monthSheet.Range[column + ":" + column].Find(text, LookAt: XlLookAt.xlPart);
            
            if (rng == null)
            {
                return 0;
            }
            else
            {
                int firstRow = rng.Row;

                Range dupe = rng.FindNext();
                int nextRow = dupe.Row;

                if( nextRow == dupe.Row)
                {
                    return firstRow;
                }
                else
                {
                    return -1;
                }
            }
        }
    }
}
