using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.TestTools.UITesting;
using RTA.Automation.CRM.UI;
using RTA.Automation.CRM.Environment;
using OpenQA.Selenium;
using RTA.Automation.CRM.Pages;
using System.Collections.ObjectModel;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reflection;
using RTA.Automation.CRM.DataSource;
using System.Windows.Forms;
using System.Threading;

namespace RTA.Automation.CRM.Tests
{
    [TestClass]
    public class CRMNewTenancyRequestBPayFileCreation : BaseTest
    {
       

        [TestInitialize]
        public void DataSourceSetup()
        {
            //open excel
            MyApp = new Excel.Application();
            MyApp.Visible = false;
        }

        [TestMethod]
        [TestProperty("TestcaseID", "7120")]
        public void ATC7120a_CRMAC1Fileformatverification()
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
                if (MyRange.Cells[i, 1].Value.ToString()== "7120")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion


            string managepingParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();
            string amountPaidLodgement = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString();

            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();

            TenancyRequestsSearchPage tenancyRequestsSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestsSearchPage.ClickNewTenancyRequestButton();

            TenancyRequestPage tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            tenancyRequestPage.PopulateTenancyRequestValidationSuccessful(
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("REQUEST_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString(),
                managepingParty, 
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGEMENT_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("NO_ROOMS")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("WEEKLY_RENT")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString(),
                amountPaidLodgement,
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_START")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("ANTICIPATED_END")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString());

            tenancyRequestPage.ClickSaveButton();
            string tenancyrequest;
            tenancyrequest = tenancyRequestPage.GetRequestNumber();
            StringAssert.Equals(tenancyrequest, "TR-BL-");
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value = tenancyrequest;

            tenancyRequestPage.SetStatusReason("Ready for validation");
            tenancyRequestPage.ClickSaveButton();
            StringAssert.Contains(tenancyRequestPage.GetValidationStatusReason(), "Validation successful");
            tenancyRequestPage.ClickSaveCloseButton();
            
            tenancyRequestsSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestsSearchPage.SetTenancyRequestSearchText(tenancyrequest);
            Table table = new Table(tenancyRequestsSearchPage.GetSearchResultTable());
            table.ClickCellValue("Name", tenancyrequest, "Name");
            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();
            StringAssert.Contains(tenancyRequestPage.GetStatusReason(), "Pending Financials");
            StringAssert.Contains(tenancyRequestPage.GetFundedStatus(), "Payment pending"); 
            table = new Table(tenancyRequestPage.GetPaymentSummaryResultTable());
            StringAssert.Contains(table.GetCellValue("Tenancy Request", tenancyrequest, "Tenancy Request"), tenancyrequest);
            StringAssert.Contains(table.GetCellValue("Tenancy Request", tenancyrequest, "Payment Type"), "BPay");
            StringAssert.Contains(table.GetCellValue("Tenancy Request", tenancyrequest, "Amount"),
                ((double)Int32.Parse(amountPaidLodgement)).ToString("C"));
            StringAssert.Contains(table.GetCellValue("Tenancy Request", tenancyrequest, "Client"), managepingParty);
            string referencenumber = table.GetCellValue("Tenancy Request", tenancyrequest, "Reference Number");
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAY_REF_NUMBER")].Value = referencenumber;
            

            //Create BPay file with new reference number
            string dateValue = DateTime.Today.ToString("yyyyMMdd");
            string fileLocation = Utils.BPayFileCreator.bPayFileCreator(referencenumber, dateValue, tenancyrequest, MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString() + "00");
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("OUTFILE")].Value = fileLocation;

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }

        [TestMethod]
        [TestProperty("TestcaseID", "7125")]
        public void ATC7125a_CRMTESTINGEndtoEndSingleFormBPayAXSuccess()
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
                if (MyRange.Cells[i, 1].Value.ToString()== "7125")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion
            string managingParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();
            string initialContribAmount = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString();
            string lodgementAmount = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString();
            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            //read excel file and get test data for this test
            // store test data in a class so that it can be used later

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();

            TenancyRequestsSearchPage tenancyRequestsSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestsSearchPage.ClickNewTenancyRequestButton();

            TenancyRequestPage tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();
            tenancyRequestPage.PopulateTenancyRequestValidationSuccessful(
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("REQUEST_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString(),
               managingParty, 
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGEMENT_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("NO_ROOMS")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("WEEKLY_RENT")].Value.ToString(),
               initialContribAmount,
               lodgementAmount,
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_START")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("ANTICIPATED_END")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString());

            tenancyRequestPage.ClickSaveButton();
            string tenancyrequest = tenancyRequestPage.GetRequestNumber();
            StringAssert.Equals(tenancyrequest, "TR-BL-");

            tenancyRequestPage.SetStatusReason("Ready for validation");
            tenancyRequestPage.ClickSaveButton();
            StringAssert.Contains(tenancyRequestPage.GetValidationStatusReason(), "Validation successful");
            tenancyRequestPage.ClickSaveCloseButton();

            tenancyRequestsSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestsSearchPage.SetTenancyRequestSearchText(tenancyrequest);

            Assert.IsTrue(tenancyRequestsSearchPage.GetPaymentRefernceRefreshTable(tenancyrequest));

            Table table = new Table(tenancyRequestsSearchPage.GetSearchResultTable());
            table.ClickCellValue("Name", tenancyrequest, "Name");

            tenancyRequestPage = new TenancyRequestPage(driver);
            table = new Table(tenancyRequestPage.GetPaymentSummaryResultTable());

            StringAssert.Contains(table.GetCellValue("Tenancy Request", tenancyrequest, "Tenancy Request"), tenancyrequest);
            StringAssert.Contains(table.GetCellValue("Tenancy Request", tenancyrequest, "Request Batch"), "");
            StringAssert.Contains(table.GetCellValue("Tenancy Request", tenancyrequest, "Payment Type"), "BPay");
            StringAssert.Contains(table.GetCellValue("Tenancy Request", tenancyrequest, "Amount"), ((double)Int32.Parse(lodgementAmount)).ToString("C"));
            StringAssert.Contains(table.GetCellValue("Tenancy Request", tenancyrequest, "Client"), managingParty);
            string referencenumber = table.GetCellValue("Tenancy Request", tenancyrequest, "Reference Number");
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAY_REF_NUMBER")].Value = referencenumber;
            StringAssert.Contains(tenancyRequestPage.GetStatusReason(), "Pending Financials");
            StringAssert.Contains(tenancyRequestPage.GetFundedStatus(), "Payment pending");


            string dateValue = DateTime.Today.AddDays(-1).ToString("yyyyMMdd");
            string fileLocation = Utils.BPayFileCreator.bPayFileCreator(referencenumber, dateValue, tenancyrequest, lodgementAmount + "00");
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("OUTFILE")].Value = fileLocation;
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value = tenancyrequest;
            
            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion


        }

        [TestMethod]
        [TestProperty("TestcaseID", "7125")]
        public void ATC7125c_CRMTESTINGEndtoEndSingleFormBPayAXSuccess()
        {
            #region Start Up Excel
            MyApp = new Excel.Application();
            MyApp.Visible = false;
            MyBook = MyApp.Workbooks.Open(DatasourceDir + @"\TenancyRequests.xlsx", 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            MySheet = (Excel.Worksheet)MyBook.Sheets[Properties.Settings.Default.ENVIRONMENT.ToString()];
            MyRange = MySheet.UsedRange;
            
            //Get specific row for the data
            int testDataRows = MyRange.Rows.Count;
            int MyRow = 0;
            for(int i = 2; i <= testDataRows; i++)
            {
                if (MyRange.Cells[i, 1].Value.ToString()== "7125")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string tenancyrequest = MyRange.Cells[MyRow,TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value.ToString();
            string bond = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("BOND_REF")].Value.ToString();
            string initialContribution = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString();

            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRibbonButton();

            TenancySearchPage tenancySearchPage = new TenancySearchPage(driver);
            tenancySearchPage.SetTenancySearchText(bond);

            Table table = new Table(tenancySearchPage.GetSearchResultTable());
            table.SelectTableRow("Bond Number", bond);

            TenancyPage tenancyPage = new TenancyPage(driver);
            tenancyPage.HoverBondPropertyRibbonTab();
            tenancyPage.ClickBondTenancyRequestRibbonButton();
            tenancyPage.ClickSelectViewButton();
            tenancyPage.SetViewList("All Tenancy Requests");

            table = new Table(tenancyPage.GetSearchResultTable());
            StringAssert.Contains(table.GetCellValue("Name", tenancyrequest, "Name"), tenancyrequest);
            table.ClickCellValue("Name", tenancyrequest, "Name");



            TenancyRequestPage tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            StringAssert.Contains(tenancyRequestPage.GetFundedStatus(), "Financials processing successful");
            StringAssert.Contains(tenancyRequestPage.GetPropertyDataControlModeRTAFundedStatus(), "deactivated");
            StringAssert.Contains(tenancyRequestPage.GetAmountMatched(), ((double)Int32.Parse(initialContribution)).ToString("C"));
            StringAssert.Contains(tenancyRequestPage.GetStatusReason(), "Completed");




            //reveiced date?? step 19


            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }

        [TestMethod]
        [TestProperty("TestcaseID", "7121")]
        public void ATC7121a_CRMDuplicateBPAYpayment()
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
                if (MyRange.Cells[i, 1].Value.ToString()== "7121")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string initialRequestParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString();
            string initialContribution = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString();
            string managepingParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();
            string lodgementAmount = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString();

            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();

            TenancyRequestsSearchPage tenancyRequestsSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestsSearchPage.ClickNewTenancyRequestButton();

            TenancyRequestPage tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();
           
            tenancyRequestPage.PopulateTenancyRequestValidationSuccessful(
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("REQUEST_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString(),
                managepingParty, 
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGEMENT_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("NO_ROOMS")].Value.ToString(),
                initialRequestParty,
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("WEEKLY_RENT")].Value.ToString(),
                initialContribution,
                lodgementAmount,
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_START")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("ANTICIPATED_END")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString());
           
            tenancyRequestPage.ClickSaveButton();
            string tenancyrequest = tenancyRequestPage.GetRequestNumber();
            StringAssert.Equals(tenancyrequest, "TR-BL-");
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value = tenancyrequest;

            tenancyRequestPage.SetStatusReason("Ready for validation");
            tenancyRequestPage.ClickSaveButton();
            StringAssert.Contains(tenancyRequestPage.GetValidationStatusReason(), "Validation successful");
            tenancyRequestPage.ClickSaveCloseButton();

            tenancyRequestsSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestsSearchPage.SetTenancyRequestSearchText(tenancyrequest);

            Assert.IsTrue(tenancyRequestsSearchPage.GetPaymentRefernceRefreshTable(tenancyrequest));

            Table table = new Table(tenancyRequestsSearchPage.GetSearchResultTable());
            table.ClickCellValue("Name", tenancyrequest, "Name");

            tenancyRequestPage = new TenancyRequestPage(driver);
            table = new Table(tenancyRequestPage.GetPaymentSummaryResultTable());

            StringAssert.Contains(table.GetCellValue("Tenancy Request", tenancyrequest, "Tenancy Request"), tenancyrequest);
            StringAssert.Contains(table.GetCellValue("Tenancy Request", tenancyrequest, "Payment Type"), "BPay");
            StringAssert.Contains(table.GetCellValue("Tenancy Request", tenancyrequest, "Amount"), ((double)Int32.Parse(lodgementAmount)).ToString("C"));
            StringAssert.Contains(table.GetCellValue("Tenancy Request", tenancyrequest, "Client"), managepingParty);
            string referencenumber = table.GetCellValue("Tenancy Request", tenancyrequest, "Reference Number");
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAY_REF_NUMBER")].Value = referencenumber;
            StringAssert.Contains(tenancyRequestPage.GetStatusReason(), "Pending Financials");
            StringAssert.Contains(tenancyRequestPage.GetFundedStatus(), "Payment pending");
          

            string dateValue = DateTime.Today.ToString("yyyyMMdd");
            string filelocation1 = Utils.BPayFileCreator.bPayFileCreator(referencenumber, dateValue, tenancyrequest, initialContribution + "00");
            //Create another file to test duplicate scenario
            dateValue = DateTime.Today.AddDays(-1).ToString("yyyyMMdd");
            string filelocation2 = Utils.BPayFileCreator.bPayFileCreator(referencenumber, dateValue, tenancyrequest, initialContribution + "00");
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("OUTFILE")].Value = filelocation1;
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MISC1")].Value = filelocation2;
        
            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }


        [TestMethod]
        [TestProperty("TestcaseID", "6629")]
        public void ATC6629a_CRMTESTINGEndtoEndSingleFormBPayCRM()
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
                if (MyRange.Cells[i, 1].Value.ToString()== "6629")
                {
                    MyRow = i;
                    break;
                }

            }
            #endregion

            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            //read excel file and get test data for this test
            // store test data in a class so that it can be used later

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();

            TenancyRequestsSearchPage tenancyRequestsSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestsSearchPage.ClickNewTenancyRequestButton();

            TenancyRequestPage tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            tenancyRequestPage.PopulateTenancyRequestValidationSuccessful(
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("REQUEST_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString(), 
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGEMENT_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("NO_ROOMS")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("WEEKLY_RENT")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_START")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("ANTICIPATED_END")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString());

            tenancyRequestPage.ClickSaveButton();
            string tenancyrequest = tenancyRequestPage.GetRequestNumber();
            StringAssert.Equals(tenancyrequest, "TR-BL-");

            //Add a new request party
            //tenancyRequestPage.ClickRequestPartyAssociated();

            string BaseWindow = driver.CurrentWindowHandle; //Records the current window handle

            tenancyRequestPage.ClickAddNewRequestPartyImage();

            //Enter Request Party details

            driver = tenancyRequestPage.SwitchNewBrowser(driver, BaseWindow);
           
            TenancyRequestPartyPage tenancyRequestPartyPage = new TenancyRequestPartyPage(driver);
            tenancyRequestPartyPage.SetClientNameValue(
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MISC1")].Value.ToString());
            tenancyRequestPartyPage.SetAmountValue("500.00");
            tenancyRequestPartyPage.ClickSaveCloseButton();

            string amountOtherParty = (MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value + 500).ToString();
            driver = driver.SwitchTo().Window(BaseWindow);
            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.SetStatusReason("Ready for validation");
            tenancyRequestPage.ClickSaveButton();
            StringAssert.Contains(tenancyRequestPage.GetValidationStatusReason(), "Validation successful");
            tenancyRequestPage.ClickSaveCloseButton();

            tenancyRequestsSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestsSearchPage.SetTenancyRequestSearchText(tenancyrequest);

            Assert.IsTrue(tenancyRequestsSearchPage.GetPaymentRefernceRefreshTable(tenancyrequest));

            Table table = new Table(tenancyRequestsSearchPage.GetSearchResultTable());
            table.ClickCellValue("Name", tenancyrequest, "Name");

            tenancyRequestPage = new TenancyRequestPage(driver);
            table = new Table(tenancyRequestPage.GetPaymentSummaryResultTable());

            StringAssert.Contains(table.GetCellValue("Tenancy Request", tenancyrequest, "Tenancy Request"), tenancyrequest);
            StringAssert.Contains(table.GetCellValue("Tenancy Request", tenancyrequest, "Request Batch"), "");
            StringAssert.Contains(table.GetCellValue("Tenancy Request", tenancyrequest, "Payment Type"), "BPay");
            StringAssert.Contains(table.GetCellValue("Tenancy Request", tenancyrequest, "Amount"), "$2,000.00");
            StringAssert.Contains(table.GetCellValue("Tenancy Request", tenancyrequest, "Client"), "BLAIR TEST");
            string referencenumber = table.GetCellValue("Tenancy Request", tenancyrequest, "Reference Number");
            StringAssert.Contains(tenancyRequestPage.GetStatusReason(), "Pending Financials");
            StringAssert.Contains(tenancyRequestPage.GetFundedStatus(), "Payment pending");
            string dateValue = DateTime.Today.AddDays(-1).ToString("yyyyMMdd");
            string fileLocation = Utils.BPayFileCreator.bPayFileCreator(referencenumber, dateValue, tenancyrequest, "200000");
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("OUTFILE")].Value = fileLocation;
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value = tenancyrequest;
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAY_REF_NUMBER")].Value = referencenumber;
            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion


        }

        [TestMethod]
        [TestProperty("TestcaseID", "6629")]
        public void ATC6629c_CRMTESTINGEndtoEndSingleFormBPayCRM()
        {
            #region Start Up Excel
            MyApp = new Excel.Application();
            MyApp.Visible = false;
            MyBook = MyApp.Workbooks.Open(DatasourceDir + @"\TenancyRequests.xlsx", 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            MySheet = (Excel.Worksheet)MyBook.Sheets[Properties.Settings.Default.ENVIRONMENT.ToString()];
            MyRange = MySheet.UsedRange;

            //Get specific row for the data
            int testDataRows = MyRange.Rows.Count;
            int MyRow = 0;
            for (int i = 2; i <= testDataRows; i++)
            {
                if (MyRange.Cells[i, 1].Value.ToString()== "6629")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string tenancyrequest = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value.ToString();
            string bond = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("BOND_REF")].Value.ToString();

            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRibbonButton();

            TenancySearchPage tenancySearchPage = new TenancySearchPage(driver);
            tenancySearchPage.SetTenancySearchText(bond);

            Table table = new Table(tenancySearchPage.GetSearchResultTable());
            table.SelectTableRow("Bond Number", bond);

            TenancyPage tenancyPage = new TenancyPage(driver);
            tenancyPage.HoverBondPropertyRibbonTab();
            tenancyPage.ClickBondTenancyRequestRibbonButton();
            tenancyPage.ClickSelectViewButton();
            tenancyPage.SetViewList("All Tenancy Requests");

            table = new Table(tenancyPage.GetSearchResultTable());
            StringAssert.Contains(table.GetCellValue("Name", tenancyrequest, "Name"), tenancyrequest);
            table.ClickCellValue("Name", tenancyrequest, "Name");



            TenancyRequestPage tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            StringAssert.Contains(tenancyRequestPage.GetFundedStatus(), "Financials processing successful");
            StringAssert.Contains(tenancyRequestPage.GetPropertyDataControlModeRTAFundedStatus(), "deactivated");
            StringAssert.Contains(tenancyRequestPage.GetAmountMatched(), "$1,000.00");
            StringAssert.Contains(tenancyRequestPage.GetStatusReason(), "Completed");




            //reveiced date?? step 19


            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }

        [TestMethod]
        [TestProperty("TestcaseID", "6365a")]
        public void ATC6365a_CRMTESTINGEndtoEndSingleFormBPayAXSuccess()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "6365")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string managepingParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();

            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();


            TenancyRequestsSearchPage tenancyRequestsSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestsSearchPage.ClickNewTenancyRequestButton();

            TenancyRequestPage tenancyRequestPage = new TenancyRequestPage(driver);
            Thread.Sleep(3000);
            tenancyRequestPage.ClickPageTitle();
            tenancyRequestPage.PopulateTenancyRequestValidationSuccessful(
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("REQUEST_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString(),
                managepingParty,
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGEMENT_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("NO_ROOMS")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("WEEKLY_RENT")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_START")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("ANTICIPATED_END")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString());

            tenancyRequestPage.ClickSaveButton();
            string tenancyrequest = tenancyRequestPage.GetRequestNumber();
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value = tenancyrequest;
            string amount = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString();
            tenancyRequestPage.SetStatusReason("Ready for validation");
            tenancyRequestPage.ClickSaveButton();
            StringAssert.Contains(tenancyRequestPage.GetValidationStatusReason(), "Validation successful");
            tenancyRequestPage.ClickSaveCloseButton();

            tenancyRequestsSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestsSearchPage.SetTenancyRequestSearchText(tenancyrequest);
            Assert.IsTrue(tenancyRequestsSearchPage.GetPaymentRefernceRefreshTable(tenancyrequest));

            Table table = new Table(tenancyRequestsSearchPage.GetSearchResultTable());
            table.ClickCellValue("Name", tenancyrequest, "Name");

            tenancyRequestPage = new TenancyRequestPage(driver);
            table = new Table(tenancyRequestPage.GetPaymentSummaryResultTable());

            string referencenumber = table.GetCellValue("Tenancy Request", tenancyrequest, "Reference Number");
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAY_REF_NUMBER")].Value = referencenumber;
            StringAssert.Contains(tenancyRequestPage.GetStatusReason(), "Pending Financials");
            StringAssert.Contains(tenancyRequestPage.GetFundedStatus(), "Payment pending");

            //Create BPay file with new reference number
            string dateValue = DateTime.Today.ToString("yyyyMMdd");
            string fileLocation = Utils.BPayFileCreator.bPayFileCreator(referencenumber, dateValue, tenancyrequest, MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString() + "00");
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("OUTFILE")].Value = fileLocation;

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

        [TestMethod]
        [TestProperty("TestcaseID", "6365c")]
        public void ATC6365c_CRMTESTINGEndtoEndSingleFormBPayAXSuccess()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "6365")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string tenancyRequest = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value.ToString();
            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();

            TenancyRequestsSearchPage tenancyRequestsSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestsSearchPage.SetPageFilterList("All Tenancy Requests");

            Table table = new Table(tenancyRequestsSearchPage.GetSearchResultTableColHeadings());
            table.ClickTableColumnHeader("Created On");
            table.ClickTableColumnHeader("Created On");

            table = new Table(tenancyRequestsSearchPage.GetSearchResultTable());
            StringAssert.Contains(table.GetCellValue("Name", tenancyRequest, "Status Reason"), "Completed"); 
            StringAssert.Contains(table.GetCellValue("Name", tenancyRequest, "Processing Status"), "Financials processing successful");

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

        [TestMethod]
        [TestProperty("TestcaseID", "6837a")]
        public void ATC6837a_CRMCancelBPayRequestBatchWithTopup()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "Topup_6837")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string managepingParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();

            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();


            TenancyRequestsSearchPage tenancyRequestsSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestsSearchPage.ClickNewTenancyRequestButton();

            TenancyRequestPage tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();
            tenancyRequestPage.PopulateTenancyRequestValidationSuccessful(
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("REQUEST_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString(),
                managepingParty,
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGEMENT_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("NO_ROOMS")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("WEEKLY_RENT")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_START")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("ANTICIPATED_END")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString());

            tenancyRequestPage.SetDwellingTypeListValue(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("DWELLING_TYPE")].Value.ToString());

            tenancyRequestPage.ClickSaveButton();
            string tenancyrequest = tenancyRequestPage.GetRequestNumber();
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value = tenancyrequest;
            string amount = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString();
            tenancyRequestPage.SetStatusReason("Ready for validation");
            tenancyRequestPage.ClickSaveButton();
            StringAssert.Contains(tenancyRequestPage.GetValidationStatusReason(), "Validation successful");
            tenancyRequestPage.ClickSaveCloseButton();

            tenancyRequestsSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestsSearchPage.SetTenancyRequestSearchText(tenancyrequest);
            Assert.IsTrue(tenancyRequestsSearchPage.GetPaymentRefernceRefreshTable(tenancyrequest));

            Table table = new Table(tenancyRequestsSearchPage.GetSearchResultTable());
            table.ClickCellValue("Name", tenancyrequest, "Name");

            tenancyRequestPage = new TenancyRequestPage(driver);
            table = new Table(tenancyRequestPage.GetPaymentSummaryResultTable());

            string referencenumber = table.GetCellValue("Tenancy Request", tenancyrequest, "Reference Number");
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAY_REF_NUMBER")].Value = referencenumber;
            StringAssert.Contains(tenancyRequestPage.GetStatusReason(), "Pending Financials");
            StringAssert.Contains(tenancyRequestPage.GetFundedStatus(), "Payment pending");

            //Create BPay file with new reference number
            string dateValue = DateTime.Today.ToString("yyyyMMdd");
            string fileLocation = Utils.BPayFileCreator.bPayFileCreator(referencenumber, dateValue, tenancyrequest, MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString() + "00");
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("OUTFILE")].Value = fileLocation;

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }
               
        [TestMethod]
        [TestProperty("TestcaseID", "6837c")]
        public void ATC6837c_CRMCancelBPayRequestBatchWithTopup()
        {
            #region Start Up Excel
            MyBook = MyApp.Workbooks.Open(DatasourceDir + @"\TenancyRequests.xlsx", 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            MySheet = (Excel.Worksheet)MyBook.Sheets[Properties.Settings.Default.ENVIRONMENT.ToString()];
            MyRange = MySheet.UsedRange;

            //Get specific row for the data
            int testDataRows = MyRange.Rows.Count;
            int tRow1 = 0;
            for (int i = 2; i <= testDataRows; i++)
            {
                if (MyRange.Cells[i, 1].Value.ToString() == "Topup_6837")
                {
                    tRow1 = i;
                    break;
                }
            }
            #endregion

            string managepingParty = MyRange.Cells[tRow1, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();
            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);
            //Creating topup tenancy req for 6837
            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();

            TenancyRequestsSearchPage tenancyRequestsSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestsSearchPage.ClickNewTenancyRequestButton();

            TenancyRequestPage tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();
            tenancyRequestPage.PopulateTenancyRequestValidationSuccessful(
                MyRange.Cells[tRow1, TenancyRequestSchema.GetColumnIndex("REQUEST_TYPE")].Value.ToString(),
                MyRange.Cells[tRow1, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString(),
                managepingParty,
                MyRange.Cells[tRow1, TenancyRequestSchema.GetColumnIndex("TENANCY_TYPE")].Value.ToString(),
                MyRange.Cells[tRow1, TenancyRequestSchema.GetColumnIndex("MANAGEMENT_TYPE")].Value.ToString(),
                MyRange.Cells[tRow1, TenancyRequestSchema.GetColumnIndex("NO_ROOMS")].Value.ToString(),
                MyRange.Cells[tRow1, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString(),
                (MyRange.Cells[tRow1, TenancyRequestSchema.GetColumnIndex("WEEKLY_RENT")].Value + 50).ToString(),
                "25",
                "50",
                "Top up",
                MyRange.Cells[tRow1, TenancyRequestSchema.GetColumnIndex("TENANCY_START")].Value.ToString(),
                MyRange.Cells[tRow1, TenancyRequestSchema.GetColumnIndex("ANTICIPATED_END")].Value.ToString(),
                MyRange.Cells[tRow1, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString());
            tenancyRequestPage.SetDwellingTypeListValue(MyRange.Cells[tRow1, TenancyRequestSchema.GetColumnIndex("DWELLING_TYPE")].Value.ToString());

            tenancyRequestPage.SetTenancyValue(MyRange.Cells[tRow1 , TenancyRequestSchema.GetColumnIndex("BOND_REF")].Value.ToString());
            tenancyRequestPage.ClickSaveButton();
            string topuptenancyrequest = tenancyRequestPage.GetRequestNumber();
            MyRange.Cells[tRow1, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value = topuptenancyrequest;

            tenancyRequestPage.ClickSaveCloseButton();

            //Creating one initial tenacy request for 6837
            int tRow2 = 0;
            for (int i = 2; i <= testDataRows; i++)
            {
                if (MyRange.Cells[i, 1].Value.ToString() == "Initial_6837")
                {
                    tRow2 = i;
                    break;
                }
            }

            tenancyRequestsSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestsSearchPage.ClickNewTenancyRequestButton();

            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            tenancyRequestPage.PopulateTenancyRequestValidationSuccessful(
                MyRange.Cells[tRow2, TenancyRequestSchema.GetColumnIndex("REQUEST_TYPE")].Value.ToString(),
                MyRange.Cells[tRow2, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString(),
                managepingParty,
                MyRange.Cells[tRow2, TenancyRequestSchema.GetColumnIndex("TENANCY_TYPE")].Value.ToString(),
                MyRange.Cells[tRow2, TenancyRequestSchema.GetColumnIndex("MANAGEMENT_TYPE")].Value.ToString(),
                MyRange.Cells[tRow2, TenancyRequestSchema.GetColumnIndex("NO_ROOMS")].Value.ToString(),
                MyRange.Cells[tRow2, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString(),
                MyRange.Cells[tRow2, TenancyRequestSchema.GetColumnIndex("WEEKLY_RENT")].Value.ToString(),
                MyRange.Cells[tRow2, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString(),
                MyRange.Cells[tRow2, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString(),
                MyRange.Cells[tRow2, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString(),
                MyRange.Cells[tRow2, TenancyRequestSchema.GetColumnIndex("TENANCY_START")].Value.ToString(),
                MyRange.Cells[tRow2, TenancyRequestSchema.GetColumnIndex("ANTICIPATED_END")].Value.ToString(),
                MyRange.Cells[tRow2, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString());
            tenancyRequestPage.SetDwellingTypeListValue(MyRange.Cells[tRow2, TenancyRequestSchema.GetColumnIndex("DWELLING_TYPE")].Value.ToString());

            tenancyRequestPage.ClickSaveButton();
            string initialtenancyrequest = tenancyRequestPage.GetRequestNumber();
            tenancyRequestPage.ClickSaveCloseButton();
            MyRange.Cells[tRow2, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value = initialtenancyrequest;
        }

        [TestMethod]
        [TestProperty("TestcaseID", "6837d")]
        public void ATC6837d_CRMCancelBPayRequestBatchWithTopup()
        {
            #region Start Up Excel
            MyBook = MyApp.Workbooks.Open(DatasourceDir + @"\TenancyRequests.xlsx", 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            MySheet = (Excel.Worksheet)MyBook.Sheets[Properties.Settings.Default.ENVIRONMENT.ToString()];
            MyRange = MySheet.UsedRange;
            
            /*Data preparation:
            Request batch is required with the following:
 
            - Several TRs (Type = Bond Lodgment) under it, including initial lodgement and at least one top-up (i.e. linked to existing Bond Number via {Tenancy} field, must have same managing party / contributors / tenancy address)
 
            - RB and TR all have {Payment Type} = BPay
            - RB has been successfully validated i.e. RB and TRs have {Status Reason} = Pending financials; are read only (bar bug 6343)
            - RB BPay reference is generated for total amount of all TRs*/

            //Get specific row for the data
            int testDataRows = MyRange.Rows.Count;
            int tRow1 = 0;
            for (int i = 2; i <= testDataRows; i++)
            {
                if (MyRange.Cells[i, 1].Value.ToString() == "Initial_6837")
                {
                    tRow1 = i;
                    break;
                }
            }
            #endregion

            string initialtenancyrequest = MyRange.Cells[tRow1, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value.ToString();

            int tRow2 = 0;
            for (int i = 2; i <= testDataRows; i++)
            {
                if (MyRange.Cells[i, 1].Value.ToString() == "Topup_6837")
                {
                    tRow2 = i;
                    break;
                }
            }
            string topuptenancyrequest = MyRange.Cells[tRow2, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value.ToString();
            
            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);
            
            HomePage homePage = new HomePage(driver);

            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();

            homePage.ClickRequestBatchRibbonButton();

            RequestBatchesSearchPage requestBatchSearchPage = new RequestBatchesSearchPage(driver);
            requestBatchSearchPage.ClickNewRequestBatchButton();

            RequestBatchPage requestBatchPage = new RequestBatchPage(driver);
            requestBatchPage.ClickPageTitle();
            requestBatchPage.SetManagingPartyText(MyRange.Cells[tRow1, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString());
            requestBatchPage.SetPaymentType(MyRange.Cells[tRow1, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString());
            requestBatchPage.ClickSaveButton();
            string requestBatch = requestBatchPage.GetRequestNumber();

            StringAssert.Contains(requestBatch, "TRB-BL-");
            
            requestBatchPage.ClickAddAssociatedRequestsButton();
            requestBatchPage.SetAssociatedRequest(initialtenancyrequest);

            requestBatchPage.ClickSaveCloseButton();

            requestBatchSearchPage = new RequestBatchesSearchPage(driver);
           
            requestBatchSearchPage.SetRequestBatchSearchText(requestBatch);
            Table reqBatchTable = new Table(requestBatchSearchPage.GetSearchResultTable());
            reqBatchTable.ClickCellValue("Name", requestBatch, "Name");

            requestBatchPage = new RequestBatchPage(driver);
            requestBatchPage.ClickAddAssociatedRequestsButton();
            requestBatchPage.SetAssociatedRequest(topuptenancyrequest);

            requestBatchPage.ClickSaveButton();
            requestBatchPage.SetStatusReason("Ready for validation");

            requestBatchPage.ClickSaveCloseButton();

            requestBatchSearchPage = new RequestBatchesSearchPage(driver);
            Assert.IsTrue(requestBatchSearchPage.GetPaymentRefernceRefreshTable(requestBatch));
            requestBatchSearchPage.SetRequestBatchSearchText(requestBatch);
            Table requestBatchTable = new Table(requestBatchSearchPage.GetSearchResultTable());
            requestBatchTable.ClickCellValue("Name", requestBatch, "Name");

            requestBatchPage = new RequestBatchPage(driver);
            requestBatchPage.ClickPageTitle();
            Assert.AreEqual(requestBatchPage.GetSumBondamountPaid(), "$2,050.00");

            requestBatchTable = new Table(requestBatchPage.GetPaymentSummaryResultTable());

            StringAssert.Contains(requestBatchTable.GetCellValue("Request Batch", requestBatch, "Request Batch"), requestBatch,"Validate the payment ref record in req batch");
            StringAssert.Contains(requestBatchTable.GetCellValue("Request Batch", requestBatch, "Payment Type"), "BPay","Validate the payment ref record in req batch");
            StringAssert.Contains(requestBatchTable.GetCellValue("Request Batch", requestBatch, "Amount"), "$2,050.00","Validate the payment ref record in req batch");
            StringAssert.Contains(requestBatchTable.GetCellValue("Request Batch", requestBatch, "Client"), MyRange.Cells[tRow1, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString(),"Validate the payment ref record in req batch");

            string reqbatreferencenumber = requestBatchTable.GetCellValue("Request Batch", requestBatch, "Reference Number");
           
            Thread.Sleep(5000);

            Assert.AreEqual(requestBatchPage.GetStatusReason(), "Pending Financials","Validate the status reason of req batch:"+requestBatch);
            Assert.AreEqual(requestBatchPage.GetFundedStatus(), "Payment pending","Validate the processing status of req batch:"+requestBatch);
            Assert.AreEqual(requestBatchPage.GetRecordStatus(), "Read only","Validate the record status of req bact:"+requestBatch);

            Table tenancyReqTable = new Table(requestBatchPage.GetTenancyRequestTable());
            tenancyReqTable.ClickCellValue("Name", initialtenancyrequest, "Name");

            TenancyRequestPage tenancyReqPage = new TenancyRequestPage(driver);
            tenancyReqPage.ClickPageTitle();
            Assert.AreEqual(tenancyReqPage.GetStatusReason(), "Pending Financials", "Validating the status reason of TR" + initialtenancyrequest);
            Assert.AreEqual(tenancyReqPage.GetRecordStatus(), "Read only","Validating the record status of TR:"+initialtenancyrequest);

            homePage.HoverRBSRibbonTab();
            homePage.ClickRequestBatchRibbonButton();
            requestBatchSearchPage = new RequestBatchesSearchPage(driver);
            requestBatchSearchPage.SetRequestBatchSearchText(requestBatch);
            reqBatchTable = new Table(requestBatchSearchPage.GetSearchResultTable());
            reqBatchTable.ClickCellValue("Name", requestBatch, "Name");

            requestBatchPage = new RequestBatchPage(driver);

            tenancyReqTable = new Table(requestBatchPage.GetTenancyRequestTable());
            tenancyReqTable.ClickCellValue("Name", topuptenancyrequest, "Name");

            tenancyReqPage = new TenancyRequestPage(driver);
            tenancyReqPage.ClickPageTitle();
            Assert.AreEqual(requestBatchPage.GetStatusReason(), "Pending Financials", "Validating the status reason of TR" + topuptenancyrequest);
            Assert.AreEqual(tenancyReqPage.GetRecordStatus(), "Read only","Validating the record status of TR"+topuptenancyrequest);

            //Double-click on the BPay Payment Reference record under the Administration section.
            homePage.HoverRBSRibbonTab();
            homePage.ClickRequestBatchRibbonButton();
            requestBatchSearchPage = new RequestBatchesSearchPage(driver);
            requestBatchSearchPage.SetRequestBatchSearchText(requestBatch);
            reqBatchTable = new Table(requestBatchSearchPage.GetSearchResultTable());
            reqBatchTable.ClickCellValue("Name", requestBatch, "Name");

            requestBatchPage = new RequestBatchPage(driver);
            requestBatchTable = new Table(requestBatchPage.GetPaymentSummaryResultTable());
            IWebElement elem = requestBatchTable.GetCellElementContainsValue("Request Batch", requestBatch, "Reference Number");
            UICommon.DoubleClickElement(elem, driver);

            //Click the [Deactivate] button.Record is deactivated.
            PaymentReferncePage payRefPage = new PaymentReferncePage(driver);
            payRefPage.ClickDeactivateButton();

            WarningDialogueFramePage warnPage = new WarningDialogueFramePage(driver);            
            warnPage.ClickProcessBeginButton();
            Thread.Sleep(3000);

            payRefPage = new PaymentReferncePage(driver);
            payRefPage.ClickPageTitle();
            Assert.AreEqual("Inactive", payRefPage.GetInactiveStatusFooter(),"Validate whether payment record of req batch:"+requestBatch+",is inactive after deactivating the record");
            
            //Inspect Batch Request record.
            homePage.HoverRBSRibbonTab();
            homePage.ClickRequestBatchRibbonButton();
            requestBatchSearchPage = new RequestBatchesSearchPage(driver);
            requestBatchSearchPage.SetRequestBatchSearchText(requestBatch);
            reqBatchTable = new Table(requestBatchSearchPage.GetSearchResultTable());
            reqBatchTable.ClickCellValue("Name", requestBatch, "Name");

            /*Record is active again, editable.
            {Status Reason} = New.
            Processing Status is blank.*/

            requestBatchPage = new RequestBatchPage(driver);
            requestBatchPage.ClickPageTitle();

            Assert.AreEqual(requestBatchPage.GetStatusReason(), "New","Validating the status reason is New for req batch:"+requestBatch);
            Assert.AreEqual(requestBatchPage.GetFundedStatus(), "--","Validating the processing status is blank for req batch:"+requestBatch);

            try
            {
                requestBatchPage.SetPaymentType("BPay");
            }
            catch (Exception)
            {
                new AssertFailedException("Payment Type in Req batch not editable after deactivating the payment ref record:" + requestBatch);
            }

            //Inspect all child Tenancy Request records.
            tenancyReqTable = new Table(requestBatchPage.GetTenancyRequestTable());
            tenancyReqTable.ClickCellValue("Name", initialtenancyrequest, "Name");

            /*Records are active again, editable.
            {Status Reason} = New.
            Processing Status is blank.*/
            tenancyReqPage = new TenancyRequestPage(driver);

            StringAssert.Contains(tenancyReqPage.GetStatusReason(), "New", "Validating the status reason is New for TR" + initialtenancyrequest);
            StringAssert.Contains(tenancyReqPage.GetFundedStatus(), "","Validating the processing status is blank for TR"+initialtenancyrequest);
            try
            {
                tenancyReqPage.SetRequestTypeListValue("Bond Lodgement");
            }
            catch (Exception)
            {
                new AssertFailedException("Request Type in Child Tenacy Req" + initialtenancyrequest + " not editable after deactivating the payment ref record in req batch:" + requestBatch);
            }

            homePage.HoverRBSRibbonTab();
            homePage.ClickRequestBatchRibbonButton();
            requestBatchSearchPage = new RequestBatchesSearchPage(driver);
            requestBatchSearchPage.SetRequestBatchSearchText(requestBatch);

            reqBatchTable = new Table(requestBatchSearchPage.GetSearchResultTable());
            reqBatchTable.ClickCellValue("Name", requestBatch, "Name");

            requestBatchPage = new RequestBatchPage(driver);

            tenancyReqTable = new Table(requestBatchPage.GetTenancyRequestTable());
            tenancyReqTable.ClickCellValue("Name", topuptenancyrequest, "Name");

            tenancyReqPage = new TenancyRequestPage(driver);

            StringAssert.Contains(tenancyReqPage.GetStatusReason(), "New", "Validating the status reason is New for TR:" + topuptenancyrequest);
            StringAssert.Contains(tenancyReqPage.GetFundedStatus(), "","Validating the processing status is blank for TR:" + topuptenancyrequest);
            try
            {
                tenancyReqPage.SetRequestTypeListValue("Bond Lodgement");
            }
            catch (Exception)
            {
                new AssertFailedException("Request Type in Child Tenacy Req" + topuptenancyrequest + " not editable after deactivating the payment ref record in req batch:" + requestBatch);
            }
        }

        [TestMethod]
        [TestProperty("TestcaseID", "6904a")]
        public void ATC6904a_CRMTopupExcessBondValidation()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "Topup_6904")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            //Data preparation step-Tenancy with bond balance of 100

            //Creating a tenancy request for the data prep
            string managepingParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();

            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();

            TenancyRequestsSearchPage tenancyRequestsSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestsSearchPage.ClickNewTenancyRequestButton();

            TenancyRequestPage tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            tenancyRequestPage.SetRequestTypeListValue(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("REQUEST_TYPE")].Value.ToString());
            string rentalPremises = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString();
            
            string[] address = rentalPremises.Split(',');
            
            string roadno = address[0].Split(' ')[0];
            string roadname = address[0].Split(' ')[1];
            string locality = address[1].Split(' ')[1] + "," + address[2] + "," + address[3];
            
            tenancyRequestPage.CreateNewAddress(roadno,roadname,locality,"Room");
            tenancyRequestPage = new TenancyRequestPage(driver);

            tenancyRequestPage.SetManagingPartyListValue(managepingParty);
            tenancyRequestPage.SetTenancyTypeListValue(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_TYPE")].Value.ToString());
            tenancyRequestPage.SetTenancyManagementTypeListValue(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGEMENT_TYPE")].Value.ToString());
            tenancyRequestPage.SetInitialRequestPartyWithSearch(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString());
            tenancyRequestPage.SetInitialConrtibution(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString());
            tenancyRequestPage.SetAmountPaidWithLodgement(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString());
            tenancyRequestPage.SetLodgementTypeListValue(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString());

            tenancyRequestPage.SetWeeklyRent(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("WEEKLY_RENT")].Value.ToString());
            tenancyRequestPage.SetTenancyStartDate(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_START")].Value.ToString());
            tenancyRequestPage.SetAnticipatedEndDate(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("ANTICIPATED_END")].Value.ToString());
            tenancyRequestPage.SetPaymentType(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString());
                
            tenancyRequestPage.SetDwellingTypeListValue(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("DWELLING_TYPE")].Value.ToString());

            tenancyRequestPage.ClickSaveButton();

            string tenancyrequest = tenancyRequestPage.GetRequestNumber();
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value = tenancyrequest;
            string amount = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString();
            tenancyRequestPage.SetStatusReason("Ready for validation");
            tenancyRequestPage.ClickSaveButton();
            StringAssert.Contains(tenancyRequestPage.GetValidationStatusReason(), "Validation successful");
            tenancyRequestPage.ClickSaveCloseButton();

            tenancyRequestsSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestsSearchPage.SetTenancyRequestSearchText(tenancyrequest);
            Assert.IsTrue(tenancyRequestsSearchPage.GetPaymentRefernceRefreshTable(tenancyrequest));

            Table table = new Table(tenancyRequestsSearchPage.GetSearchResultTable());
            table.ClickCellValue("Name", tenancyrequest, "Name");

            tenancyRequestPage = new TenancyRequestPage(driver);
            table = new Table(tenancyRequestPage.GetPaymentSummaryResultTable());

            string referencenumber = table.GetCellValue("Tenancy Request", tenancyrequest, "Reference Number");
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAY_REF_NUMBER")].Value = referencenumber;
            StringAssert.Contains(tenancyRequestPage.GetStatusReason(), "Pending Financials");
            StringAssert.Contains(tenancyRequestPage.GetFundedStatus(), "Payment pending");

            //Create BPay file with new reference number
            string dateValue = DateTime.Today.ToString("yyyyMMdd");
            string fileLocation = Utils.BPayFileCreator.bPayFileCreator(referencenumber, dateValue, tenancyrequest, MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString() + "00");
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("OUTFILE")].Value = fileLocation;

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

        [TestMethod]
        [TestProperty("TestcaseID", "6904c")]
        public void ATC6904c_CRMTopupExcessBondValidation()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "Topup_6904")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string managepingParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();
            
            //Login as RBS user
            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);
           
            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();

            //(Template) - Create new Tenancy Request (Bond Lodgement)
            TenancyRequestsSearchPage tenancyRequestsSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestsSearchPage.ClickNewTenancyRequestButton();

            TenancyRequestPage tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();
            
            /*Enter details.Details entered 
            - Tenancy (prepared Tenancy)
            - Rooming Accommodation
            - Rent subsidy
            - rent $200
            - amount paId $701*/
            tenancyRequestPage.SetRequestTypeListValue(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("REQUEST_TYPE")].Value.ToString());
            tenancyRequestPage.SetRentalPremisesValue("*"+MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString());
            tenancyRequestPage.SetManagingPartyListValue(managepingParty);
            tenancyRequestPage.SetTenancyTypeListValue(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_TYPE")].Value.ToString());
            tenancyRequestPage.SetTenancyManagementTypeListValue(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGEMENT_TYPE")].Value.ToString());
            tenancyRequestPage.SetInitialRequestPartyWithSearch(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString());
            tenancyRequestPage.SetInitialConrtibution("701");
            tenancyRequestPage.SetAmountPaidWithLodgement("701");
            tenancyRequestPage.SetLodgementTypeListValue("Top up");

            tenancyRequestPage.SetWeeklyRent(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("WEEKLY_RENT")].Value.ToString());
            tenancyRequestPage.SetTenancyStartDate(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_START")].Value.ToString());
            tenancyRequestPage.SetAnticipatedEndDate(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("ANTICIPATED_END")].Value.ToString());
            tenancyRequestPage.SetPaymentType(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString());
            tenancyRequestPage.SetDwellingTypeListValue(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("DWELLING_TYPE")].Value.ToString());

            tenancyRequestPage.SetSubsidy("Yes");

            tenancyRequestPage.SetTenancyValue(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("BOND_REF")].Value.ToString());
            
            //Save Tenancy Request
            tenancyRequestPage.ClickSaveButton();

            string tenancyrequest = tenancyRequestPage.GetRequestNumber();
            //Update Status Reason to 'Ready for validation'
            tenancyRequestPage.SetStatusReason("Ready for validation");
            tenancyRequestPage.ClickSaveCloseButton();

            tenancyRequestsSearchPage = new TenancyRequestsSearchPage(driver);

            tenancyRequestsSearchPage.SetTenancyRequestSearchText(tenancyrequest);
            Table table = new Table(tenancyRequestsSearchPage.GetSearchResultTable());
            table.ClickCellValue("Name", tenancyrequest, "Name");
            
            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            //Confirm Status Reason set to 'Validation failed'
            StringAssert.Contains(tenancyRequestPage.GetStatusReason(), "Validation failed");

            //Confirm Queue reason for excess bond added.Status set to 'To be resolved'
            //tenancyRequestPage.ClickQueueReasons();
            Table queueReasonTable = new Table(tenancyRequestPage.GetQueueReasonTable());
            StringAssert.Contains(queueReasonTable.GetCellContainsValue("Reason", "Invalid bond amount : Exceeds maximum bond", "Status Reason"), "To be resolved");

            //Confirm Maximum Allowed Bond set correctly
            //Assert: Maximum Allowed Bond set to $2800
            Assert.AreEqual(tenancyRequestPage.GetMaximumAllowedAmount(), "$800.00");
            
        }
    }
}
