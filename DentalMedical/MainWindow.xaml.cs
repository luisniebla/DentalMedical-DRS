using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using System.IO;
using System.Text.RegularExpressions;
using System.Collections;

using System.Data.OleDb;
using System.Data;

using Excel = Microsoft.Office.Interop.Excel;
using ExcelDataReader;
using System.Data;
using MySql.Data.MySqlClient;
using System.Runtime.InteropServices;
using System.Data.Common;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;



namespace DentalMedical
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public object Application { get; private set; }
        //public static Excel.Application xlApp = new Excel.Application();
        //public static Excel.Workbook xlCBPTemplate = xlApp.Workbooks.Open(@"//serv-az/drs_client_admin/Call Campaigns/ZZ_CBPProccess/CBP Template April (Blackford).xlsx", ReadOnly:true);
        //Excel.Workbook xlCBPData = xlApp.Workbooks.Open(@"\\serv-az\drs_client_admin\Call Campaigns\ZZ_Merge & CBP RawData\CallBackProof Data\042118\Blackford Lifetime CBP.xlsx");

        //Excel.Worksheet xlFoundAppts = xlCBPTemplate.Worksheets[3];
        public MainWindow()
        {
            InitializeComponent();
            //xlApp.DisplayAlerts = false;

            //ReadExcelHeaders(xlCBPData.Worksheets[4], "CBPData");
            //xlCBPData.Close();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            listBoxSearchResults.Items.Clear();
            listBoxSearchResults.UpdateLayout();

            IO myIO = new IO();

            string campaign = TextBoxSearchCriteria.Text;
            string searchCriteria = "*" + campaign + "*Data*.xlsm";

            ArrayList foundFiles = myIO.getFiles(@"//serv-az/drs_ofc/05-DENTAL-MEDICAL/", searchCriteria);

            foreach (string file in foundFiles.ToArray())
            {
                listBoxSearchResults.Items.Add(file);
            }

            listBoxSearchResults.UpdateLayout();

            LabelSearchNotification.Content = "DONE";
        }

        PIMACampaign thc;
        Excel.Application xlApp;

        private void BtnOpenSelected_Click(object sender, RoutedEventArgs e)
        {
            //Excel.Application xlApp = new Excel.Application();

            string password = TextBoxPassword.Text;
            string campaign = TextBoxSearchCriteria.Text;
            string filePath = "";
            try
            {
                filePath = listBoxSearchResults.SelectedItem.ToString();
            }catch(NullReferenceException)
            {
                MessageBox.Show("Please select a listbox item");
                return;
            }
            // Excel.Application is too slow. Let's try something else.

            xlApp = new Excel.Application();

            
        }

        
        public void pimaCampaignCBP(string password, string filePath, string campaign)
        {
            try
            {
                thc = new PIMACampaign(xlApp, password, filePath, campaign, "May 2018");

                Debug.WriteLine(thc.HeadersToString());


                DGCBP.DataContext = thc.GetCBPDataView("pima_greenvalley_cbp_may_results_53018");
                DGCBP.UpdateLayout();
            }
            catch (MySql.Data.MySqlClient.MySqlException mysqle)
            {
                MessageBox.Show("Error during SQL transactions " + mysqle.ToString());
                thc = null;
                xlApp.Quit();
            }
            catch (IndexOutOfRangeException ex)
            {
                Debug.WriteLine(ex.ToString());
                MessageBox.Show("Could not find column headers");
                thc = null;
                xlApp.Quit();
            }
        }
        public void oledb()
        {
            using (OleDbConnection conn = new OleDbConnection())
            {
                System.Data.DataTable dt = new System.Data.DataTable();
                string Import_FileName = @"C:\Users\data\Desktop\exceloutput.xlsx";
                string fileExtension = System.IO.Path.GetExtension(Import_FileName);
                if (fileExtension == ".xls")
                    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 8.0;HDR=YES;'";
                if (fileExtension == ".xlsx")
                    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'";


                using (OleDbCommand comm = new OleDbCommand())
                {
                    comm.CommandText = "Select * from [" + "Master" + "$]";

                    comm.Connection = conn;

                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    {
                        da.SelectCommand = comm;
                        da.Fill(dt);
                    }

                }
            }
        }
        private void BtnImport_Click(object sender, RoutedEventArgs e)
        { 
            // Append found to here.
        }

        private void listBoxSQLTables_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // This isn't being used currently but I do think it will be used in the future.
        }

        private void App_Close(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // DOn't leave any excel processes open
            if(thc != null) 
                thc.close();    
            if(xlApp != null)
                xlApp.Quit();

        }

        private void BtnCBP_Click(object sender, RoutedEventArgs e)
        {
            int rsults = thc.AttemptCallBackProof();
            Debug.WriteLine("DONE WITH CBP");
            thc.close("Post_CBP");
            thc = null;
            // Don't quite out of the excel app just yet
        }
    }
}
