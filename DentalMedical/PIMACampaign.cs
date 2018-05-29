using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace DentalMedical
{
    class PIMACampaign : ExcelCampaign
    {
        private int numberOfColumns;
        private string firstHeaderString;

        public PIMACampaign()
        {

        }
        public PIMACampaign(Application xlApp, string password, string path, string title, string month ) : base( xlApp,  password,  path,  title,  month)
        {
            numberOfColumns = 26;
            firstHeaderString = "Provider";

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
        public void AttemptCallBackProof()
        {
            // TODO: Logmeins
            CallBackProof cbp = new CallBackProof();
            cbp.Show();

            DBConnection db = new DBConnection();
            var dt = new System.Data.DataTable();
            if (db.IsConnect())
            {
                
                dt.Load(db.QueryDB("SELECT * FROM pima_westside_cbp_may_results_52918;"));
                cbp.DataGridCBP.DataContext = dt.DefaultView;
                cbp.DataGridCBP.UpdateLayout();
                DataRow[] tbl = dt.Select();
                foreach (DataRow row in dt.Rows)
                {
                    string personNumber = row.Field<string>(1);
                    int lastRow = GetLastRow("May-Merge 2018");
                    Range found = monthSheet.Range["B1", "B" + lastRow].Find(personNumber, LookAt:XlLookAt.xlWhole);
                    if (found != null)
                        cbp.TextBlockExcel.Text += found.Value + " | ";

                    
                }

            }
            else
            {
                Debug.WriteLine("Could not connect to DB");
                throw new Exception("Could not connec to DB");
            }
            
            
        }
    }
}
