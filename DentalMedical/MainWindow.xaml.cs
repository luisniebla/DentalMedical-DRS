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

        private void BtnOpenSelected_Click(object sender, RoutedEventArgs e)
        {
            ExcelToMySQL handler = new ExcelToMySQL();
            
            string password = TextBoxPassword.Text;
            string campaign = TextBoxSearchCriteria.Text;
            string filePath = listBoxSearchResults.SelectedItem.ToString();

            ExcelCampaign selectedCampaign = handler.OpenCampaign(filePath, password);

            try
            {
                //ArrayList selectedCampaignSheets = selectedCampaign.GetWorksheets();
                
                
                

                
                //Range dncLine = xlMaster.Range["A1", "A10"].Find("DO NOT CALL ABOVE LINE", MatchCase: false);


                //MessageBox.Show("FOUND: " + dncLine.Row.ToString());

                string ext = System.IO.Path.GetExtension(filePath);
                selectedCampaign.xlWorkbook.SaveAs(@"C:\Users\Data\Desktop\pretty." + ext);
                selectedCampaign.close();
            }
            catch(Exception ex)
            {
                Debug.Write("Could not open excel workbook " + campaign + "\n" + ex.ToString());
            }
            
            handler.CloseExcel();
            //OpenXMLSDK readXL = new OpenXMLSDK(listBoxSearchResults.SelectedItem.ToString());

            //MessageBox.Show(readXL.SAXRead());
            //readXL.LoopRows();
            //readXL.SAXRead();
            //handler.ExportExcelHeaders(listBoxSearchResults.SelectedItem.ToString(),"May 2018", @"C:\Users\data\Desktop\" + campaign + "month.csv", password);
            //handler.ExportExcelHeaders(listBoxSearchResults.SelectedItem.ToString(), "Master", @"C:\Users\data\Desktop\" + campaign + "master.csv", password);

            //handler.CloseExcel();
        }
       
        private void BtnImport_Click(object sender, RoutedEventArgs e)
        {
            

            // Append found to here.
        }

        private void listBoxSQLTables_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            
        }

        private void App_Close(object sender, System.ComponentModel.CancelEventArgs e)
        {
            
        }
    }
}
