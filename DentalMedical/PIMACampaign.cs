using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Office.Interop.Excel;

namespace DentalMedical
{ 
    /// <summary>
    /// TODO: There's a lot of potential for abstraction here
    /// THat is to say, we have the lower-level methods for comparing data, changing data up, which columns which we change up
    /// as parameters, and then have the higher order classes decide on those magic values, rather than making them magic 
    /// values across the board.
    /// 
    /// TODO: It'd be nice if we could get some sort of algorithm in place to determine CBP. It's actually not that complex, 90% of situations are the same.
    /// i.e. If all there are, are notes, then it's obviously a pend/previous.
    /// </summary>
    class PIMACampaign : ExcelCampaign
    {
        private int numberOfColumns;
        private System.Data.DataTable dt;
        
        public PIMACampaign(Microsoft.Office.Interop.Excel.Application xlApp, string password, string path, string title, string month ) : base( xlApp,  password,  path,  title,  month)
        {
            numberOfColumns = 26;
            dt = new System.Data.DataTable();
            headerFlag = "Provider";
        }

        /// <summary>
        /// Convert the PIMA campaign header to a string
        /// </summary>
        /// <returns>The PIMA column headers delimitted by a |</returns>
        public string HeadersToString()
        {
            string headerString = "";
            foreach (object header in ExportHeaders("Provider", numberOfColumns)[0])
            {
                if (header != null)
                    headerString += header.ToString() + "|";
            }
            return headerString;
        }


        /// <summary>
        /// Query the database to select the entire table, sqlTableName
        /// </summary>
        /// <param name="sqlTableName">The SQL table to query</param>
        /// <returns>
        /// The DataView object representing the result from the SQL query
        /// null if the database could not be connected to. Should I throw an error here? Probably...
        /// </returns>
        public DataView GetCBPDataView(string sqlTableName)
        {
            DBConnection db = new DBConnection();
            if (db.IsConnect())
            {
                dt.Load(db.QueryDB("SELECT * FROM " + sqlTableName + ";" ));
                return dt.DefaultView;
            }
            else
            {
                
                 new Exception("Could not connect to SQL");
                return null;
            }
        }
        /// <summary>
        /// Attempt to do a call back proof
        /// Prerequisites:
        /// - The found_appts tables (master and month) have been created on sql server
        ///     - This is just so we can easily read through it... not totally necessary i suppose, but one step at a time.
        /// 
        /// Goals:
        /// - A user should be able to cross-proof the existing Excel database with the found matches.
        /// - A user should be able to update the excel database by telling it whether it's an appointment or not.
        /// 
        /// </summary>
        public int AttemptCallBackProof()
        {
            dt.AcceptChanges();

            if (monthHeaders == null)
                HeadersToString();  // Make sure to export those headers. Useful for later on when we access them directly.

            int resColIndex = 15; //FindMonthColumnIndexByHeader("Resolution");
            int notesColIndex = 20; //FindMonthColumnIndexByHeader("Notes");
            int updateColIndex = 24; //FindMonthColumnIndexByHeader("Update");
            int apptColIndex = 16;  // PIMA is special in that the appt date isn't unique, and there wouldbe an error if we searched like this.
            int newApptColorIndex = 50;
            string updateStr = DateTime.Today.ToShortDateString() + " LNR";
            string PtIdCol = "B";

            int apptCount = 0;
            int pendCount = 0;
            int keepCount = 0;
            int dataTableRowCount = 0;
            DataRow[] tbl = dt.Select();
            int totalMatches = 0;
            // TODO: For each row, find the corresponding range location in Excel, and update the Excel sheet accordingly.
            foreach (DataRow row in dt.Rows)
            {
                dataTableRowCount++;

                string personNumber = row.Field<string>(1);

                int lastRow = GetLastRow(Month);
                Range found = monthSheet.Range[PtIdCol + "1", PtIdCol + lastRow].Find(personNumber, LookAt: XlLookAt.xlWhole);
                if (found != null)
                {
                    string flag = row.Field<string>(8);
                    string resolution = row.Field<string>(7);
                    string appt = row.Field<string>(0);

                    
                    // Process the patient
                    // TODO: Move the indexs and resolution to global variables accesss/whatever, and move these repetitive stuff into their own functions
                    switch (flag)
                    {
                        // Process the appointment case
                        case "A":
                            monthSheet.Cells[found.Row, resColIndex] = "Appointment";
                            monthSheet.Cells[found.Row, resColIndex].Interior.ColorIndex = 0;   // Clear the queue
                            monthSheet.Cells[found.Row, updateColIndex] = updateStr;
                            monthSheet.Cells[found.Row, notesColIndex].Interior.ColorIndex = newApptColorIndex;
                            monthSheet.Cells[found.Row, apptColIndex] = appt;
                            apptCount++;
                            break;
                        // Process pend/previous
                        case "":
                            monthSheet.Cells[found.Row, resColIndex] = "Pend/Previous Appt";
                            monthSheet.Cells[found.Row, resColIndex].Interior.ColorIndex = 0;
                            monthSheet.Cells[found.Row, updateColIndex] = updateStr;
                            monthSheet.Cells[found.Row, apptColIndex] = appt;
                            pendCount++;
                            break;
                        // Don't change the resolutions, just add new appt
                        case "K":
                            monthSheet.Cells[found.Row, updateColIndex] = updateStr;
                            monthSheet.Cells[found.Row, resColIndex].Interior.ColorIndex = 0;
                            monthSheet.Cells[found.Row, apptColIndex] = appt;
                            keepCount++;
                            break;
                    }
                    totalMatches++;
                }
                else
                {
                    MessageBox.Show("WARNING: Could not find " + personNumber + " in Worksheet");
                }
            }
            MessageBox.Show("Number of rows in Data Table " + dataTableRowCount
                + "\n Number of matches found in Excel: " + totalMatches
                + "\n Number of changed Appts: " + apptCount
                + "\n Number of changed Pend/Prev: " + pendCount
                + "\n Number of kept resolutions: " + keepCount
                );

            
            return totalMatches;
        }


        public void AttemptCBPRegex()
        {
            // TODO: The program will do its best to try and do the callback proof, then the user just proofs it.
            //string note = row.Field<string>(9);
            Regex FindRelevantNote = new Regex(@".*[(][0-9]{4}[1][8][A-Z][A-Z][)]\s+[^(].*");
           // if (FindRelevantNote.IsMatch(note))
                //row.SetField<string>(8, "MATCH");
            //else
                //row.SetField<string>(8, "NO MATCH");
        }
        
    }
}
