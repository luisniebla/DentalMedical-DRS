using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DentalMedical
{
    class Blackford
    {
        public string apptTable;
        public string monthTable;
        public string masterTable;

        DBConnection dbc;
        ExcelCampaign excelCampaign;
        string campaign = "6th_Street";
        string[] blackfordDataReportHeaders = {"First Name","Last Name","Telephone","DIAL","Alt #","Email","DOB/Age","Last Visit","Resolution","Date","Time","Notes","Call Back Date","CC Type","Provider","Insurance","CBP", "#Dials"  };
        string[] bfCBPHeaders = { "LName", "FName",	"MI","Status", "BirthDate", "HPhone", "MPhone", "Billing_Type", "Prov_Name", "LastVisit", "Appt_Date", "Appt_Time", "CC_TypeName", "CC_TypeDesc", "CC_Note" };
    
        public Blackford()
        {
        dbc = new DBConnection();
        }

        public void SetApptTable(string campaign, string fileLoc)
        {
            string tblName = campaign + "_appts";
            if (dbc.IsConnect())
            {
                dbc.CreateStringTable(tblName, blackfordDataReportHeaders);
                apptTable = tblName;
            }
        }

        public void SetMonthTable(string campaign, string fileLoc)
        {

        }
       
        public void AttemptTableCreation(string tableName, string fileLoc)
        {
            if (dbc.IsConnect())
            {
                dbc.CreateStringTable(tableName, blackfordDataReportHeaders);
            }
        }

        public void AttemptCallBackProof()
        {

        }
    }
}
