﻿using System;
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
            
            ExcelCampaign thisCampaign = new ExcelCampaign(xlApp, password, filePath, campaign, "June 2018");

            MessageBox.Show(thisCampaign.FindMonthColumnIndexByHeader("Status").ToString());

            thisCampaign.Close();
            xlApp.Quit();
            
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
            if (thc != null)
                thc.Close();
            if(xlApp != null)
                xlApp.Quit();

        }

        private void BtnCBP_Click(object sender, RoutedEventArgs e)
        {
            int rsults = thc.AttemptCallBackProof();
            Debug.WriteLine("DONE WITH CBP");
            thc.Close();
            thc = null;
            // Don't quite out of the excel app just yet
        }
    }
}
