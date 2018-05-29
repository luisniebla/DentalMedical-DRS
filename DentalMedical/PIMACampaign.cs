using System;
using System.Collections;
using System.Collections.Generic;
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

        public PIMACampaign(Application xlApp, string password, string path, string title, string month ) : base( xlApp,  password,  path,  title,  month)
        {
            numberOfColumns = 26;
            firstHeaderString = "Provider";

            headerFlag = "Provider";
        }

        public string HeadersToString()
        {
            string headerString = "";
            foreach (object header in ExportHeaders("Provider", 26)[0])
            {
                headerString += header.ToString() + "|";
            }
            return headerString;
        }

        /// <summary>
        /// Attempt to do a call back proof
        /// Prerequisites:
        /// - The found_appts table has been matched
        /// - 
        /// </summary>
        public void AttemptCallBackProof()
        {

        }
    }
}
