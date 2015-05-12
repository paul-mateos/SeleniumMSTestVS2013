using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Input;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;
using RTA.Automation.AX.Utils;
using RTA.Automation.AX.Pages;
using System.IO;
using RTA.Automation.AX.UI;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using RTA.Automation.CRM.DataSource;



namespace RTA.Automation.AX.Tests
{
   
    [CodedUITest]
    public class AXClientTests : TestBase
    {

       
        
        public AXClientTests()
        {
        }


        #region TestInitialize
        //Use TestInitialize to run code before running each test 
        [TestInitialize()]
        public override void TestInitialize()
        {
            // Call a common routine to set up the test
            PlayBackSettings.StartTest();
            base.TestInitialize();

            //open excel
            MyApp = new Excel.Application();
            MyApp.Visible = false;

        
        }
        #endregion

        [TestMethod]
        [TestProperty("TestCaseID", "6362")]
        public void ATC6362b_E2ESingleBPAYCancelTenancyRequest()
        {
         
            #region Start Up Excel
            MyBook = MyApp.Workbooks.Open(DatasourceDir + @"\TenancyRequests.xlsx", 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            MySheet = (Excel.Worksheet)MyBook.Sheets[Properties.Settings.Default.ENVIRONMENT.ToString()];
            MyRange = MySheet.UsedRange;
            //Get specific row for the data
            int testDataRows = MyRange.Rows.Count;
            int MyRow = 0;
            for (int i = 2; i <= testDataRows; i++)
            {
                if (MyRange.Cells[i, 1].Value.ToString() == "6362")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion


            string managingParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();

            Homepage homePage = new Homepage();
            homePage.ClickHomeTab();
            homePage.ClickBondManagementTab();
            homePage.ClickBondClientLink();
            BondClientPage bondClientPage = new BondClientPage();
            Table table = new Table(bondClientPage.GetClientOverviewTable());
            Assert.IsFalse(table.GetCellValueExists("Name", managingParty));


            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion


        }

        [TestMethod]
        [TestProperty("TestCaseID", "4434")]
        public void ATC4434_ClientDetailsfromCRMOrganizations()
        {

            #region Start Up Excel
            MyBook = MyApp.Workbooks.Open(DatasourceDir + @"\TenancyRequests.xlsx", 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            MySheet = (Excel.Worksheet)MyBook.Sheets[Properties.Settings.Default.ENVIRONMENT.ToString()];
            MyRange = MySheet.UsedRange;
            //Get specific row for the data
            int testDataRows = MyRange.Rows.Count;
            int MyRow = 0;
            for (int i = 2; i <= testDataRows; i++)
            {
                if (MyRange.Cells[i, 1].Value.ToString() == "4434")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion


            string managingParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();

            Homepage homePage = new Homepage();
            homePage.ClickHomeTab();
            homePage.ClickBondManagementTab();
            homePage.ClickBondClientLink();
            BondClientPage bondClientPage = new BondClientPage();
            Table table = new Table(bondClientPage.GetClientOverviewTable());
            Assert.IsFalse(table.GetCellValueExists("Name", managingParty));


            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion


        }
        
        
        
        //[TestMethod]
        //[TestProperty("TestCaseID", "6371")]
        //public void ATC6371_DeactivatingClientsIDClientwithtransactions()
        //{

        //    #region Start Up Excel
        //    MyBook = MyApp.Workbooks.Open(DatasourceDir + @"\Clients.xlsx", 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
        //    MySheet = (Excel.Worksheet)MyBook.Sheets[Properties.Settings.Default.ENVIRONMENT.ToString()];
        //    MyRange = MySheet.UsedRange;
        //    //Get specific row for the data
        //    int testDataRows = MyRange.Rows.Count;
        //    int MyRow = 0;
        //    for (int i = 2; i <= testDataRows; i++)
        //    {
        //        if (MyRange.Cells[i, 1].Value.ToString() == "6371")
        //        {
        //            MyRow = i;
        //            break;
        //        }
        //    }
        //    #endregion


        //    string client = MyRange.Cells[MyRow, ClientsSchema.GetColumnIndex("CLIENT_NAME")].Value.ToString();

        //    Homepage homePage = new Homepage();
        //    homePage.ClickHomeTab();
        //    homePage.ClickBondManagementTab();
        //    homePage.ClickBondClientLink();
        //    BondClientPage bondClientPage = new BondClientPage();
        //    Table table = new Table(bondClientPage.GetClientOverviewTable());
        //    table.ClickCellValue("Name", client, "Name");
        //    bondClientPage.ClickSetup();
        //    bondClientPage.ClickDeactivateMenuItem();

        //    MicrosoftDynamicsAXPage microsoftDynamicsAXPage = new MicrosoftDynamicsAXPage();
        //    microsoftDynamicsAXPage.SetDeactivationReasonText("Test reason");
        //    microsoftDynamicsAXPage.ClickOkButton();

        //    InfoLogPage infoLogPage = new InfoLogPage();
        //    Assert.IsTrue(infoLogPage.GetTreeItemExists("Only clients with zero balance can be deactivated!", "Message"));
           
        //    infoLogPage.ClickCloseButton();
        //    bondClientPage.ClickCloseButton();


        //    #region Shut down Excel
        //    MyBook.Save();
        //    MyBook.Close();
        //    MyApp.Quit();
        //    #endregion


        //}

        [TestMethod]
        [TestProperty("TestCaseID", "4434")]
        public void ATC4434_ClientsDetailsFromCRM()
        {
            Homepage homePage = new Homepage();
            homePage.ClickHomeTab();
            homePage.ClickBondManagementTab();
            homePage.ClickBondClientLink();

            BondClientPage bondClientPage = new BondClientPage();
            Table table = new Table(bondClientPage.GetClientOverviewTable());

            // Verify there are fields to display CRM client number, Name, Client group and none of these fields can be updated
            Assert.IsFalse(table.SetCellValue("CRM client number", 1, "Test 4434"), "CRM client number row value is not editable");
            Assert.IsFalse(table.SetCellValue("Name", 1, "Test 4434"), "Name row value is not editable");
            Assert.IsFalse(table.SetCellValue("Client group", 1, "Test 4434"), "Client groupr row value is not editable");

            // Verify there is a field to display Method of Payment and this field can not be updated
            bondClientPage.ClickGeneralTab();
            Assert.IsTrue(bondClientPage.IsMethodOfPaymentEditable() , "Method Of Payment edit box is not read only!!!");
        }

        #region Test Clean Up
        [TestCleanup()]
        public override void TestCleanup()
        {
            base.TestCleanup();
        }
        #endregion   
    }
}
