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
    class PIMACampaign : ExcelCampaign
    {
        private int numberOfColumns;
        private string firstHeaderString;
        private System.Data.DataTable dt;
        public PIMACampaign()
        {

        }
        public PIMACampaign(Microsoft.Office.Interop.Excel.Application xlApp, string password, string path, string title, string month ) : base( xlApp,  password,  path,  title,  month)
        {
            numberOfColumns = 26;
            firstHeaderString = "Provider";
            dt = new System.Data.DataTable();
            headerFlag = "Provider";
        }

        public string HeadersToString()
        {
            string headerString = "";
            foreach (object header in ExportHeaders("Provider", 25)[0])
            {
                if (header != null)
                    headerString += header.ToString() + "|";
            }
            return headerString;
        }

        public DataView GetCBPDataView()
        {
            DBConnection db = new DBConnection();
            if (db.IsConnect())
            {
                dt.Load(db.QueryDB("SELECT * FROM pima_westside_cbp_may_results_52918;"));
                return dt.DefaultView;
            }
            else
            {
                Debug.WriteLine("Could not read SQL");
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
                DataRow[] tbl = dt.Select();
                int totalMatches = 0;
                // TODO: For each row, find the corresponding range location in Excel, and update the Excel sheet accordingly.
                foreach (DataRow row in dt.Rows)
                {
                    string personNumber = row.Field<string>(1);
                    
                    int lastRow = GetLastRow("May-Merge 2018");
                    Range found = monthSheet.Range["B1", "B" + lastRow].Find(personNumber, LookAt:XlLookAt.xlWhole);
                    if (found != null)
                    {
                        
                        totalMatches++;
                    }

                    // TODO: The program will do its best to try and do the callback proof, then the user just proofs it.
                    string note = row.Field<string>(9);
                    Regex FindRelevantNote = new Regex(@".*[(][0-9]{4}[1][8][A-Z][A-Z][)]\s+[^(].*");
                    if (FindRelevantNote.IsMatch(note))
                        row.SetField<string>(8, "MATCH");
                    else
                        row.SetField<string>(8, "NO MATCH");
                }


            return totalMatches;
        }

        
    }
}
