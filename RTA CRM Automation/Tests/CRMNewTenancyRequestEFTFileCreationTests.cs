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
using System.Threading;

namespace RTA.Automation.CRM.Tests
{
    [TestClass]
    public class CRMNewTenancyRequestEFTFileCreationTests : BaseTest
    {



        [TestInitialize]
        public void DataSourceSetup()
        {
            //open excel
            MyApp = new Excel.Application();
            MyApp.Visible = false;

        }

        [TestMethod]
        [TestProperty("TestcaseID", "7081")]
        public void ATC7081a_CRMEFTreceiptforSingleBond()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "7081")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string managepingParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();
            string tenancyrequest;

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
            string amountPaidLodgement = ((double)Int32.Parse(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString())).ToString("C");
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

            tenancyrequest = tenancyRequestPage.GetRequestNumber();
            StringAssert.Contains(tenancyrequest, "TR-BL-");
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value = tenancyrequest;

            //Add a new request party
            //tenancyRequestPage.ClickRequestPartyAssociated();

            string BaseWindow = driver.CurrentWindowHandle; //Records the current window handle
            tenancyRequestPage.ClickAddNewRequestPartyImage();

            //Enter Request Party details
            driver = tenancyRequestPage.SwitchNewBrowser(driver, BaseWindow);
            TenancyRequestPartyPage tenancyRequestPartyPage = new TenancyRequestPartyPage(driver);
            tenancyRequestPartyPage.SetClientNameValue(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MISC1")].Value.ToString());
            tenancyRequestPartyPage.SetAmountValue(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MISC2")].Value.ToString());
            tenancyRequestPartyPage.ClickSaveCloseButton();
            string amountOtherParty = (MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value +
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MISC2")].Value).ToString();
            driver = driver.SwitchTo().Window(BaseWindow);
            
            tenancyRequestPage = new TenancyRequestPage(driver);
            Thread.Sleep(2000);
            tenancyRequestPage.CheckForErrors();
            tenancyRequestPage.ClickPageTitle();
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
            StringAssert.Contains(table.GetCellValue("Tenancy Request", tenancyrequest, "Payment Type"), "EFT");
            StringAssert.Contains(table.GetCellValue("Tenancy Request", tenancyrequest, "Amount"), amountPaidLodgement);
            StringAssert.Contains(table.GetCellValue("Tenancy Request", tenancyrequest, "Client"), managepingParty);
            string referencenumber = table.GetCellValue("Tenancy Request", tenancyrequest, "Reference Number");
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAY_REF_NUMBER")].Value = referencenumber;
            StringAssert.Contains(tenancyRequestPage.GetStatusReason(), "Pending Financials");
            StringAssert.Contains(tenancyRequestPage.GetFundedStatus(), "Payment pending");

            //Create EFT file with new reference number
            //string fileLocation = Utils.EFTFileCreator.eFTFileCreator(tenancyrequest, referencenumber);
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("OUTFILE")].Value = Utils.BAI2FileCreator.bAI2FileCreator(referencenumber, amountOtherParty);


            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

        [TestMethod]
        [TestProperty("TestcaseID", "8589")]
        public void ATC85891a_CRM7098EFTReceiptRecordEFTreceiptfrombankstatementitem()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "85891")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

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
            string initialRequestParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString();
            string initialContribution = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString();
            string managingParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();
            tenancyRequestPage.PopulateTenancyRequestValidationSuccessful(
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("REQUEST_TYPE")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString(),
              managingParty, 
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_TYPE")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGEMENT_TYPE")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("NO_ROOMS")].Value.ToString(),
              initialRequestParty,
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("WEEKLY_RENT")].Value.ToString(),
              initialContribution,
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_START")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("ANTICIPATED_END")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString());


            tenancyRequestPage.ClickSaveButton();
            string tenancyrequest = tenancyRequestPage.GetRequestNumber();
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value = tenancyrequest;
            StringAssert.Contains(tenancyrequest, "TR-BL-");

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
            StringAssert.Contains(table.GetCellValue("Tenancy Request", tenancyrequest, "Payment Type"), "EFT");
            StringAssert.Contains(table.GetCellValue("Tenancy Request", tenancyrequest, "Amount"), "$2,000.00");
            StringAssert.Contains(table.GetCellValue("Tenancy Request", tenancyrequest, "Client"), managingParty);
            string referencenumber = table.GetCellValue("Tenancy Request", tenancyrequest, "Reference Number");
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAY_REF_NUMBER")].Value = referencenumber;
            StringAssert.Contains(tenancyRequestPage.GetStatusReason(), "Pending Financials");
            StringAssert.Contains(tenancyRequestPage.GetFundedStatus(), "Payment pending");

            //Create BPay file with new reference number

            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("OUTFILE")].Value = Utils.BAI2FileCreator.bAI2FileCreator(referencenumber, initialContribution);


            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }

        [TestMethod]
        [TestProperty("TestcaseID", "8589")]
        public void ATC85892a_CRM7098EFTReceiptRecordEFTreceiptfrombankstatementitem()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "858921")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

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
            string initialRequestParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString();
            string initialContribution = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString();
            tenancyRequestPage.PopulateTenancyRequestValidationSuccessful(
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("REQUEST_TYPE")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString(), 
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_TYPE")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGEMENT_TYPE")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("NO_ROOMS")].Value.ToString(),
              initialRequestParty,
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("WEEKLY_RENT")].Value.ToString(),
              initialContribution,
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_START")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("ANTICIPATED_END")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString());


            tenancyRequestPage.ClickSaveButton();
            string tenancyrequest = tenancyRequestPage.GetRequestNumber();
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value = tenancyrequest;
            StringAssert.Contains(tenancyrequest, "TR-BL-");
            tenancyRequestPage.ClickSaveCloseButton();

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }

        [TestMethod]
        [TestProperty("TestcaseID", "8589")]
        public void ATC85892b_CRM7098EFTReceiptRecordEFTreceiptfrombankstatementitem()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "858922")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

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
            string initialRequestParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString();
            string initialContribution = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString();
            tenancyRequestPage.PopulateTenancyRequestValidationSuccessful(
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("REQUEST_TYPE")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString(), 
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_TYPE")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGEMENT_TYPE")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("NO_ROOMS")].Value.ToString(),
              initialRequestParty,
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("WEEKLY_RENT")].Value.ToString(),
              initialContribution,
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_START")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("ANTICIPATED_END")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString());


            tenancyRequestPage.ClickSaveButton();
            string tenancyrequest = tenancyRequestPage.GetRequestNumber();
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value = tenancyrequest;
            StringAssert.Contains(tenancyrequest, "TR-BL-");
            tenancyRequestPage.ClickSaveCloseButton();

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }

        [TestMethod]
        [TestProperty("TestcaseID", "8589")]
        public void ATC85892c_CRM7098EFTReceiptRecordEFTreceiptfrombankstatementitem()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "858923")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            User user = this.environment.GetUser(SecurityRole.Default);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();

            homePage.ClickRequestBatchRibbonButton();

            RequestBatchesSearchPage requestBatchSearchPage = new RequestBatchesSearchPage(driver);
            requestBatchSearchPage.ClickNewRequestBatchButton();

            RequestBatchPage requestBatchPage = new RequestBatchPage(driver);
            requestBatchPage.SetManagingPartyText(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString());
            requestBatchPage.SetPaymentType(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString());
            requestBatchPage.ClickSaveButton();
            string requestBatch = requestBatchPage.GetRequestNumber();
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value = requestBatch;
            StringAssert.Contains(requestBatch, "TRB-BL-");

            //Add tenancy requests
            //Get specific row for the data
            int TR1Row = 0;
            for (int i = 2; i <= testDataRows; i++)
            {
                if (MyRange.Cells[i, 1].Value.ToString() == "858921")
                {
                    TR1Row = i;
                    break;
                }
            }
            string TR1reference = MyRange.Cells[TR1Row, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value.ToString();
            //Get specific row for the data
            int TR2Row = 0;
            for (int i = 2; i <= testDataRows; i++)
            {
                if (MyRange.Cells[i, 1].Value.ToString() == "858922")
                {
                    TR2Row = i;
                    break;
                }
            }
            string TR2reference = MyRange.Cells[TR2Row, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value.ToString();

            requestBatchPage.ClickAddAssociatedRequestsButton();
            requestBatchPage.SetAssociatedRequest(TR1reference);
            requestBatchPage.ClickAddAssociatedRequestsButton();
            requestBatchPage.SetAssociatedRequest(TR2reference);


            requestBatchPage.SetStatusReason("Ready for validation");
            //requestBatchPage.ClickSaveButton();
            //StringAssert.Contains(requestBatchPage.GetValidationStatusReason(), "Validation successful");
            requestBatchPage.ClickSaveCloseButton();

            requestBatchSearchPage = new RequestBatchesSearchPage(driver);
            requestBatchSearchPage.SetRequestBatchSearchText(requestBatch);

            Assert.IsTrue(requestBatchSearchPage.GetPaymentRefernceRefreshTable(requestBatch));

            Table table = new Table(requestBatchSearchPage.GetSearchResultTable());
            table.ClickCellValue("Name", requestBatch, "Name");

            requestBatchPage = new RequestBatchPage(driver);
            table = new Table(requestBatchPage.GetPaymentSummaryResultTable());

            StringAssert.Contains(table.GetCellValue("Request Batch", requestBatch, "Request Batch"), requestBatch);
            StringAssert.Contains(table.GetCellValue("Request Batch", requestBatch, "Payment Type"), "EFT");
            StringAssert.Contains(table.GetCellValue("Request Batch", requestBatch, "Amount"), "$4,000.00");
            StringAssert.Contains(table.GetCellValue("Request Batch", requestBatch, "Client"), MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString());
            string referencenumber = table.GetCellValue("Request Batch", requestBatch, "Reference Number");
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAY_REF_NUMBER")].Value = referencenumber;
            StringAssert.Contains(requestBatchPage.GetStatusReason(), "Pending Financials");
            StringAssert.Contains(requestBatchPage.GetFundedStatus(), "Payment pending");

            //Create EFT file with new reference number            
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("OUTFILE")].Value =
                Utils.BAI2FileCreator.bAI2FileCreator(referencenumber, MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString());


            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }

        [TestMethod]
        [TestProperty("TestcaseID", "TopUp_TestData")]
        public void ATC_TopUpTestDataa()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "TopUp_TestData")
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

            tenancyRequestPage.ClickSaveButton();
            string tenancyrequest = tenancyRequestPage.GetRequestNumber();
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value = tenancyrequest;

            //Add a new request party
            string BaseWindow = driver.CurrentWindowHandle; //Records the current window handle
            tenancyRequestPage.ClickAddNewRequestPartyImage();

            //Enter Request Party details
            driver = tenancyRequestPage.SwitchNewBrowser(driver, BaseWindow);
            TenancyRequestPartyPage tenancyRequestPartyPage = new TenancyRequestPartyPage(driver);
            tenancyRequestPartyPage.SetClientNameValue(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MISC1")].Value.ToString());
            tenancyRequestPartyPage.SetAmountValue(
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString());
            tenancyRequestPartyPage.ClickSaveCloseButton();
            string amountOtherParty = (MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value + 500).ToString();
            driver = driver.SwitchTo().Window(BaseWindow);

            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();
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

            //Create EFT file with new reference number
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("OUTFILE")].Value = Utils.BAI2FileCreator.bAI2FileCreator(referencenumber, amountOtherParty);


            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

        [TestMethod]
        [TestProperty("TestcaseID", "6709")]
        public void ATC6709c_CRMTenancyRequestBatchReadOnlywhenBPaygenerated()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "6709c")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            User user = this.environment.GetUser(SecurityRole.Default);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();

            homePage.ClickRequestBatchRibbonButton();

            RequestBatchesSearchPage requestBatchSearchPage = new RequestBatchesSearchPage(driver);
            requestBatchSearchPage.ClickNewRequestBatchButton();

            RequestBatchPage requestBatchPage = new RequestBatchPage(driver);
            requestBatchPage.SetManagingPartyText(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString());
            requestBatchPage.SetPaymentType(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString());
            requestBatchPage.ClickSaveButton();
            string requestBatch = requestBatchPage.GetRequestNumber();
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value = requestBatch;
            StringAssert.Contains(requestBatch, "TRB-BL-");

            //Add tenancy requests
            //Get specific row for the data
            int TR1Row = 0;
            for (int i = 2; i <= testDataRows; i++)
            {
                if (MyRange.Cells[i, 1].Value.ToString() == "6709a")
                {
                    TR1Row = i;
                    break;
                }
            }
            string TR1reference = MyRange.Cells[TR1Row, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value.ToString();
            //Get specific row for the data
            int TR2Row = 0;
            for (int i = 2; i <= testDataRows; i++)
            {
                if (MyRange.Cells[i, 1].Value.ToString() == "6709b")
                {
                    TR2Row = i;
                    break;
                }
            }
            string TR2reference = MyRange.Cells[TR2Row, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value.ToString();

            requestBatchPage.ClickAddAssociatedRequestsButton();
            requestBatchPage.SetAssociatedRequest(TR1reference);
            requestBatchPage.ClickAddAssociatedRequestsButton();
            requestBatchPage.SetAssociatedRequest(TR2reference);


            requestBatchPage.SetStatusReason("Ready for validation");
            //requestBatchPage.ClickSaveButton();
            //StringAssert.Contains(requestBatchPage.GetValidationStatusReason(), "Validation successful");
            requestBatchPage.ClickSaveCloseButton();

            requestBatchSearchPage = new RequestBatchesSearchPage(driver);
            requestBatchSearchPage.SetRequestBatchSearchText(requestBatch);

            Assert.IsTrue(requestBatchSearchPage.GetPaymentRefernceRefreshTable(requestBatch));

            Table table = new Table(requestBatchSearchPage.GetSearchResultTable());
            table.ClickCellValue("Name", requestBatch, "Name");

            requestBatchPage = new RequestBatchPage(driver);
            table = new Table(requestBatchPage.GetPaymentSummaryResultTable());

            StringAssert.Contains(table.GetCellValue("Request Batch", requestBatch, "Request Batch"), requestBatch);
            StringAssert.Contains(table.GetCellValue("Request Batch", requestBatch, "Payment Type"), "BPay");
            StringAssert.Contains(table.GetCellValue("Request Batch", requestBatch, "Amount"), "$1,600.00");
            StringAssert.Contains(table.GetCellValue("Request Batch", requestBatch, "Client"), MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString());
            string referencenumber = table.GetCellValue("Request Batch", requestBatch, "Reference Number");
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAY_REF_NUMBER")].Value = referencenumber;
            StringAssert.Contains(requestBatchPage.GetStatusReason(), "Pending Financials");
            StringAssert.Contains(requestBatchPage.GetFundedStatus(), "Payment pending");

            //Create EFT file with new reference number            
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("OUTFILE")].Value =
                Utils.BAI2FileCreator.bAI2FileCreator(referencenumber, MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString());


            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }

        [TestMethod]
        [TestProperty("TestcaseID", "6709")]
        public void ATC6709a_CRMTenancyRequestBatchReadOnlywhenBPaygenerated()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "6709a")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string initialRequestParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString();
            string initialContribution = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString();
            string tenancyrequest;
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
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString(), 
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_TYPE")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGEMENT_TYPE")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("NO_ROOMS")].Value.ToString(),
              initialRequestParty,
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("WEEKLY_RENT")].Value.ToString(),
              initialContribution,
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_START")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("ANTICIPATED_END")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString());


            tenancyRequestPage.ClickSaveButton();
            tenancyrequest = tenancyRequestPage.GetRequestNumber();
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value = tenancyrequest;
            StringAssert.Contains(tenancyrequest, "TR-BL-");
            tenancyRequestPage.ClickSaveCloseButton();

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }

        [TestMethod]
        [TestProperty("TestcaseID", "6709")]
        public void ATC6709b_CRMTenancyRequestBatchReadOnlywhenBPaygenerated()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "6709b")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string initialRequestParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString();
            string initialContribution = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString();
            string tenancyrequest;
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
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString(), 
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_TYPE")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGEMENT_TYPE")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("NO_ROOMS")].Value.ToString(),
              initialRequestParty,
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("WEEKLY_RENT")].Value.ToString(),
              initialContribution,
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_START")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("ANTICIPATED_END")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString());


            tenancyRequestPage.ClickSaveButton();
            tenancyrequest = tenancyRequestPage.GetRequestNumber();
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value = tenancyrequest;
            StringAssert.Contains(tenancyrequest, "TR-BL-");
            tenancyRequestPage.ClickSaveCloseButton();

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }

        [TestMethod]
        [TestProperty("TestcaseID", "6911a")]
        public void ATC6911a_TenancytopupvalidationSameaddresssamemanagingparty()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "6911")
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

            tenancyRequestPage.ClickSaveButton();
            string tenancyrequest = tenancyRequestPage.GetRequestNumber();
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value = tenancyrequest;

            //Add a new request party
            string BaseWindow = driver.CurrentWindowHandle; //Records the current window handle
            tenancyRequestPage.ClickAddNewRequestPartyImage();

            //Enter Request Party details
            driver = tenancyRequestPage.SwitchNewBrowser(driver, BaseWindow);
            TenancyRequestPartyPage tenancyRequestPartyPage = new TenancyRequestPartyPage(driver);
            Thread.Sleep(3000);
            tenancyRequestPartyPage.SetClientNameValue(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MISC1")].Value.ToString());
            tenancyRequestPartyPage.SetAmountValue(
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString());
            tenancyRequestPartyPage.ClickSaveCloseButton();
            string amountOtherParty = (MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value + 500).ToString();

            driver = driver.SwitchTo().Window(BaseWindow);

            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.CheckForErrors();
            tenancyRequestPage.ClickPageTitle();
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

            //Create EFT file with new reference number
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("OUTFILE")].Value = Utils.BAI2FileCreator.bAI2FileCreator(referencenumber, amountOtherParty);


            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

        [TestMethod]
        [TestProperty("TestcaseID", "6911c")]
        public void ATC6911c_TenancytopupvalidationSameaddresssamemanagingparty()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "6911")
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
                (MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("WEEKLY_RENT")].Value + 50).ToString(),
                "25",
                "50",
                "Top up",
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_START")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("ANTICIPATED_END")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString());

            tenancyRequestPage.SetTenancyValue(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("BOND_REF")].Value.ToString());
            tenancyRequestPage.ClickSaveButton();
            string tenancyrequest = tenancyRequestPage.GetRequestNumber();
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value = tenancyrequest;

            tenancyRequestPage.SetStatusReason("Ready for validation");
            tenancyRequestPage.ClickSaveButton();
            Thread.Sleep(5000);
            tenancyRequestPage.ClickSaveButton();
            Table table = new Table(tenancyRequestPage.GetQueueReasonTable());
            StringAssert.Contains(table.GetCellContainsValue("Reason", "Invalid list of Parties", "Tenancy Request"), tenancyrequest);
            table.ClickCellContainsValueEnterRow("Reason", "Invalid list of Parties", "Status Reason");

            TenancyRequestQueueReasonPage tenancyRequestQueueReasonPage = new TenancyRequestQueueReasonPage(driver);
            tenancyRequestQueueReasonPage.SetStatusReasonValue("Resolved", driver);
            tenancyRequestQueueReasonPage.ClickSaveCloseButton();

            tenancyRequestPage = new TenancyRequestPage(driver);
            table = new Table(tenancyRequestPage.GetQueueReasonTable());
            StringAssert.Contains(table.GetCellValue("Tenancy Request", tenancyrequest, "Status Reason"), "Resolved");

            //Add a new request party
            string BaseWindow = driver.CurrentWindowHandle; //Records the current window handle
            tenancyRequestPage.ClickAddNewRequestPartyImage();

            //Enter Request Party details
            driver = tenancyRequestPage.SwitchNewBrowser(driver, BaseWindow);
            TenancyRequestPartyPage tenancyRequestPartyPage = new TenancyRequestPartyPage(driver);
            tenancyRequestPartyPage.SetClientNameValue(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MISC1")].Value.ToString());
            tenancyRequestPartyPage.SetAmountValue("25");
            tenancyRequestPartyPage.ClickSaveCloseButton();
            driver = driver.SwitchTo().Window(BaseWindow);

            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();
            tenancyRequestPage.SetStatusReason("Ready for validation");
            tenancyRequestPage.ClickSaveButton();
            Assert.AreEqual(tenancyRequestPage.GetValidationStatusReason(), "Validation successful");
            tenancyRequestPage.ClickSaveCloseButton();

            tenancyRequestsSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestsSearchPage.SetTenancyRequestSearchText(tenancyrequest);
            Assert.IsTrue(tenancyRequestsSearchPage.GetPaymentRefernceRefreshTable(tenancyrequest));

            table = new Table(tenancyRequestsSearchPage.GetSearchResultTable());
            table.ClickCellValue("Name", tenancyrequest, "Name");

            tenancyRequestPage = new TenancyRequestPage(driver);
            table = new Table(tenancyRequestPage.GetPaymentSummaryResultTable());

            string referencenumber = table.GetCellValue("Tenancy Request", tenancyrequest, "Reference Number");
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAY_REF_NUMBER")].Value = referencenumber;
            StringAssert.Contains(tenancyRequestPage.GetStatusReason(), "Pending Financials");
            StringAssert.Contains(tenancyRequestPage.GetFundedStatus(), "Payment pending");

            ////Create EFT file with new reference number
            //MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("OUTFILE")].Value = Utils.BAI2FileCreator.bAI2FileCreator(referencenumber, amountOtherParty);


            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

        [TestMethod]
        [TestProperty("TestcaseID", "6745")]
        public void ATC6745a_CRMEFTFeedbacktoCRMMultipleBonds()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "6745")
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

            tenancyRequestPage.ClickSaveButton();
            string tenancyrequest = tenancyRequestPage.GetRequestNumber();
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value = tenancyrequest;

            //Add a new request party
            string BaseWindow = driver.CurrentWindowHandle; //Records the current window handle
            tenancyRequestPage.ClickAddNewRequestPartyImage();

            //Enter Request Party details
            driver = tenancyRequestPage.SwitchNewBrowser(driver, BaseWindow);
            TenancyRequestPartyPage tenancyRequestPartyPage = new TenancyRequestPartyPage(driver);
            tenancyRequestPartyPage.SetClientNameValue(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MISC1")].Value.ToString());
            tenancyRequestPartyPage.SetAmountValue(
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString());
            tenancyRequestPartyPage.ClickSaveCloseButton();
            string amountOtherParty = (MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value +
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MISC2")].Value).ToString();
            driver = driver.SwitchTo().Window(BaseWindow);

            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();
            tenancyRequestPage.SetStatusReason("Ready for validation");
            tenancyRequestPage.ClickSaveButton();
            StringAssert.Contains(tenancyRequestPage.GetValidationStatusReason(), "Validation successful");
            Thread.Sleep(3000);
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

            //Create EFT file with new reference number
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("OUTFILE")].Value = Utils.BAI2FileCreator.bAI2FileCreator(referencenumber, amountOtherParty);


            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

        [TestMethod]
        [TestProperty("TestcaseID", "6746")]
        public void ATC6746a_CRMEFTreceiptforSingleBond()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "6746")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string managepingParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();
            string tenancyrequest;

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
            string amountPaidLodgement = ((double)Int32.Parse(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString())).ToString("C");
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

            tenancyrequest = tenancyRequestPage.GetRequestNumber();
            StringAssert.Contains(tenancyrequest, "TR-BL-");
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value = tenancyrequest;

            //Add a new request party
        
            string BaseWindow = driver.CurrentWindowHandle; //Records the current window handle
            tenancyRequestPage.ClickAddNewRequestPartyImage();

            //Enter Request Party details
            driver = tenancyRequestPage.SwitchNewBrowser(driver, BaseWindow);
            TenancyRequestPartyPage tenancyRequestPartyPage = new TenancyRequestPartyPage(driver);
            tenancyRequestPartyPage.SetClientNameValue(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MISC1")].Value.ToString());
            tenancyRequestPartyPage.SetAmountValue(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MISC2")].Value.ToString());
            tenancyRequestPartyPage.ClickSaveCloseButton();
            string amountOtherParty = (MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value +
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MISC2")].Value).ToString();
            driver = driver.SwitchTo().Window(BaseWindow);

            tenancyRequestPage = new TenancyRequestPage(driver);
           
            tenancyRequestPage.ClickPageTitle();
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
            StringAssert.Contains(table.GetCellValue("Tenancy Request", tenancyrequest, "Payment Type"), "EFT");
            StringAssert.Contains(table.GetCellValue("Tenancy Request", tenancyrequest, "Amount"), amountPaidLodgement);
            StringAssert.Contains(table.GetCellValue("Tenancy Request", tenancyrequest, "Client"), managepingParty);
            string referencenumber = table.GetCellValue("Tenancy Request", tenancyrequest, "Reference Number");
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAY_REF_NUMBER")].Value = referencenumber;
            StringAssert.Contains(tenancyRequestPage.GetStatusReason(), "Pending Financials");
            StringAssert.Contains(tenancyRequestPage.GetFundedStatus(), "Payment pending");

            //Create EFT file with new reference number
            
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("OUTFILE")].Value = Utils.BAI2FileCreator.bAI2FileCreator(referencenumber, amountOtherParty);


            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }



        [TestMethod]
        [TestProperty("TestcaseID", "6746")]
        public void ATC6746c_CRMEFTreceiptforSingleBond()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "6746")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string tenancyrequest = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value.ToString();
            string bond = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("BOND_REF")].Value.ToString();
            string initialContribution = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString();
            string amountTotalContribution = (MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value + 
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MISC2")].Value).ToString();
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
            StringAssert.Contains(tenancyRequestPage.GetAmountMatched(), ((double)Int32.Parse(amountTotalContribution)).ToString("C"));
            StringAssert.Contains(tenancyRequestPage.GetStatusReason(), "Completed");


            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }
    }

}
