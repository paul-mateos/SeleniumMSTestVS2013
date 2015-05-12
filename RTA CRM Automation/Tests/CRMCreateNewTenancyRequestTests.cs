using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.TestTools.UITesting;
using RTA.Automation.CRM.UI;
using RTA.Automation.CRM.Environment;
using OpenQA.Selenium;
using System.Threading;
using RTA.Automation.CRM.Pages;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reflection;
using RTA.Automation.CRM.DataSource;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;

namespace RTA.Automation.CRM.Tests
{
    [TestClass]
    public class CRMCreateNewTenancyRequestTests : BaseTest
    {
        private string warningMessage;
        private string controlMode;
       


        [TestInitialize]
        public void DataSourceSetup()
        {
            //open excel
            MyApp = new Excel.Application();
            MyApp.Visible = false;

        }
        [TestMethod]
        [TestProperty("TestcaseID", "4412")]
        public void ATC4412_CRMCheckTenancyRequestID()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "4412")
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
           

            TenancyRequestsSearchPage tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);

            tenancyRequestSearchPage.ClickNewTenancyRequestButton();

            TenancyRequestPage tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();
  
            tenancyRequestPage.PopulateTenancyRequestFormResidentialTenancy(
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("REQUEST_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString(), 
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGEMENT_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("NO_ROOMS")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString());


            tenancyRequestPage.ClickSaveButton();

            string caseId1 = tenancyRequestPage.GetRequestNumber();
            int caseNum1 = int.Parse(caseId1.Substring(6));

            tenancyRequestPage.ClickSaveCloseButton();

            tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestSearchPage.ClickNewTenancyRequestButton();

            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();
            tenancyRequestPage.PopulateTenancyRequestFormResidentialTenancy(
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("REQUEST_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGEMENT_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("NO_ROOMS")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString());
 
            tenancyRequestPage.ClickSaveButton();

            string caseId2 = tenancyRequestPage.GetRequestNumber();
            int caseNum2 = int.Parse(caseId2.Substring(6));

             //checks if Investigation case starts with "TR-BL-"
            StringAssert.Contains(caseId1,"TR-BL-");
            StringAssert.Contains(caseId2,"TR-BL-"); 
            Assert.AreEqual(caseNum1, caseNum2 - 1);      //if caseNum2 is a single increment from caseNum1 then TRUE

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

        [TestMethod]
        [TestProperty("TestcaseID", "6780")]
        public void ATC6780_CRMNewTenancyRequestTestResidentialTenancyTest()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "6780")
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

            //Attemp to save with NO Mandatory data
            tenancyRequestPage.SetRequestTypeListValue("Bond Lodgement");
            tenancyRequestPage.ClickSaveButton();
            //StringAssert.StartsWith(tenancyRequestPage.GetRentalPremiseAddressErrorText(), "You must provide a value for Rental Premises.");


            //Attempt to Enter Invalid Bedroom values
            tenancyRequestPage.SetNumberOfBedrooms("0");
            StringAssert.Contains(tenancyRequestPage.GetAlertMessage(), "You must enter a whole number between 1 and 12.");
            StringAssert.Contains(tenancyRequestPage.GetAlertMessage(), "You must enter a whole number between 1 and 12.");

            //Attempt to Enter Invalid Bedoom values
            tenancyRequestPage.SetNumberOfBedrooms("13");
            StringAssert.Contains(tenancyRequestPage.GetAlertMessage(), "You must enter a whole number between 1 and 12.");
            StringAssert.Contains(tenancyRequestPage.GetAlertMessage(), "You must enter a whole number between 1 and 12.");

            //Attempt to Enter Invalid Bedroom values
            tenancyRequestPage.SetNumberOfBedrooms("1");
            //StringAssert.Contains(tenancyRequestPage.GetRTAValidationMessage(), "You must");  This needs a ! contrains

            //Attemp to save with Mandatory data
            // tenancyRequestPage.PopulateTenancyRequestFormResidentialTenancy("1 THOMAS ST, BIRKDALE, QLD, 4159", "Residential Tenancy", "AMANDA TEST", "3", "AARON BALL", "700", "700", "Initial");
            tenancyRequestPage.PopulateTenancyRequestFormResidentialTenancy(
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("REQUEST_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString(), 
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGEMENT_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("NO_ROOMS")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString());

            tenancyRequestPage.ClickSaveButton();
            StringAssert.Equals(tenancyRequestPage.GetRequestNumber(), "TR-BL-");


            //Assert the warning message regarding Tenancy Start Date and Total Contribution amount
            warningMessage = tenancyRequestPage.GetWarningMessage();
            StringAssert.Contains(warningMessage, "Tenancy Start is blank, please select a date.");

            //Fill in start date and check that warning message is removed
            //tenancyRequestPage.ClickTenancyStartDate();
            tenancyRequestPage.SetTenancyStartDate("01/03/2015");
            tenancyRequestPage.ClickSaveButton();
            //StringAssert.Contains(tenancyRequestPage.GetWarningMessage(),"Tenancy Start is blank, please select a date.");
            
            //Assert that Date Bond Received at RTA is read only
            controlMode = tenancyRequestPage.GetPropertyDataControlModeRTADateReceivedAtRTA();
            StringAssert.Equals(controlMode, "locked");

            //Assert that Funded Status is read only
            controlMode = tenancyRequestPage.GetPropertyDataControlModeRTAFundedStatus();
            StringAssert.Equals(controlMode, "locked");

            String tenancyNumber = tenancyRequestPage.GetRequestNumber();
            StringAssert.StartsWith(tenancyNumber, "TR-BL-");

            //Assert that saved record can be found in the record grid
            tenancyRequestPage.ClickSaveCloseButton();
            
            tenancyRequestsSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestsSearchPage.SetTenancyRequestSearchText(tenancyNumber);
          
            Table table = new Table(tenancyRequestsSearchPage.GetSearchResultTable());
         
            //Assert that the status reason can be confirmed on the table
            StringAssert.Equals(table.GetCellValue("Name", tenancyNumber, "Status Reason"), "New");

            //Assert that a value in a particular column in a table exists
            StringAssert.Equals(table.GetCellContainsValue("Name", tenancyNumber, "Amount Bond Paid with Lodgement"),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString());
            
            ////Assert that a search result record can be opened
            table.ClickCellValue("Name", tenancyNumber, "Name");
            tenancyRequestPage = new TenancyRequestPage(driver);
            StringAssert.Equals(tenancyRequestPage.GetRequestNumber(), tenancyNumber);

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }

        [TestMethod]
        [TestProperty("TestcaseID", "4394")]
        public void ATC4394_CRMEnterdatafromForm()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "4394")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion
            
            string tenancyRequest;
            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();


            TenancyRequestsSearchPage tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);

            tenancyRequestSearchPage.ClickNewTenancyRequestButton();

            TenancyRequestPage tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();
          
            //tenancyRequestPage.PopulateTenancyRequestFormResidentialTenancy("1 THOMAS ST, BIRKDALE, QLD, 4159", "Residential Tenancy", "AMANDA TEST", "3", "AARON BALL", "700", "700", "Initial");
            tenancyRequestPage.PopulateTRNoTenancyType(
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("REQUEST_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGEMENT_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("NO_ROOMS")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString());
  
            
            tenancyRequestPage.ClickSaveButton();
            tenancyRequestPage.ClickSaveButton();
            //Assert Save unsuccessful without Tenancy Type
            StringAssert.StartsWith(tenancyRequestPage.GetTenancyTypeErrorText(), "You must provide a value for Tenancy Type.");
            tenancyRequestPage.ClickPageTitle();
            tenancyRequestPage.SetResidentialTenancyTypeList("Residential Tenancy");
            tenancyRequestPage.SetResidentialManagementTypeList("Lessor/Owner");
            //Assert Save Successful
            tenancyRequestPage.ClickSaveButton();
            StringAssert.Contains(tenancyRequestPage.GetRequestNumber(), "TR-BL-");
            tenancyRequest = tenancyRequestPage.GetRequestNumber();

            //Assert List Values
            //tenancyRequestPage.ClickTenancyTypeList();
            Assert.IsTrue(tenancyRequestPage.GetTenancyTypeListValue("Residential Tenancy"));
            Assert.IsTrue(tenancyRequestPage.GetTenancyTypeListValue("Rooming Accommodation"));
            tenancyRequestPage.SetResidentialManagementTypeList("Lessor/Owner");

            
            tenancyRequestPage.ClickSaveCloseButton();
            
            //Assert record in search table
            tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestSearchPage.SetTenancyRequestSearchText(tenancyRequest);
            Table table = new Table(tenancyRequestSearchPage.GetSearchResultTable());
            StringAssert.Contains(table.GetCellValue("Name", tenancyRequest, "Status Reason"), "New");

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

        [TestMethod]
        [TestProperty("TestcaseID", "4398")]
        public void ATC4398_CRMEnterdata2dwellingtype()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "4398")
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

            TenancyRequestsSearchPage tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestSearchPage.ClickNewTenancyRequestButton();

            TenancyRequestPage tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();
            //tenancyRequestPage.PopulateTenancyRequestFormResidentialTenancy("1 THOMAS ST, BIRKDALE, QLD, 4159", "Residential Tenancy", "AMANDA TEST", "3", "AARON BALL", "700", "700", "Initial");
            tenancyRequestPage.PopulateTenancyRequestFormResidentialTenancy(
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("REQUEST_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGEMENT_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("NO_ROOMS")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString());
                 
 
            ////Assert Dwelling Type field

            Assert.IsTrue(tenancyRequestPage.GetDwellingTypeText("Flat/Unit"));
            tenancyRequestPage.ClickPageTitle();
            Assert.IsTrue(tenancyRequestPage.GetDwellingTypeText("House"));
            tenancyRequestPage.ClickPageTitle();
            Assert.IsTrue(tenancyRequestPage.GetDwellingTypeText("Townhouse/Semi-Detached House"));
            tenancyRequestPage.ClickPageTitle();
            Assert.IsTrue(tenancyRequestPage.GetDwellingTypeText("Student Accommodation on Campus"));
            tenancyRequestPage.ClickPageTitle(); 
            Assert.IsTrue(tenancyRequestPage.GetDwellingTypeText("Moveable Dwelling/Site"));
            tenancyRequestPage.ClickPageTitle(); 
            Assert.IsTrue(tenancyRequestPage.GetDwellingTypeText("Studio"));
            tenancyRequestPage.ClickPageTitle(); 
            Assert.IsTrue(tenancyRequestPage.GetDwellingTypeText("Moveable Dwelling/Site with electricity"));

            ////Change Tenancy type
            tenancyRequestPage.SetResidentialTenancyTypeList("Rooming Accommodation");
            tenancyRequestPage.SetResidentialManagementTypeList("Lessor/Owner");
            ////Assert Dwelling Type field
            Assert.IsTrue(tenancyRequestPage.GetDwellingTypeText("Boarding House"));
            tenancyRequestPage.ClickPageTitle();
            Assert.IsTrue(tenancyRequestPage.GetDwellingTypeText("Supported Accommodation"));
            tenancyRequestPage.ClickPageTitle();
            Assert.IsTrue(tenancyRequestPage.GetDwellingTypeText("Student Accommodation off Campus"));
            tenancyRequestPage.ClickPageTitle();
            Assert.IsTrue(tenancyRequestPage.GetDwellingTypeText("Studio"));

            //Assert Save Successful
            tenancyRequestPage.ClickSaveButton();
            StringAssert.Contains(tenancyRequestPage.GetRequestNumber(), "TR-BL-");

            string tenancyRequest = tenancyRequestPage.GetRequestNumber();
            tenancyRequestPage.ClickSaveCloseButton();

            //Assert record in search table
            tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestSearchPage.SetTenancyRequestSearchText(tenancyRequest);
            Table table = new Table(tenancyRequestSearchPage.GetSearchResultTable());
            StringAssert.Contains(table.GetCellValue("Name", tenancyRequest, "Tenancy Type"), "Rooming Accommodation");

            //Change Dwelling and check search table again
            table.ClickCellValue("Name", tenancyRequest, "Name");
            tenancyRequestPage = new TenancyRequestPage(driver);

            //tenancyRequestPage.ClickTenancyTypeList();
            tenancyRequestPage.SetResidentialTenancyTypeList("Residential Tenancy");
            tenancyRequestPage.SetResidentialManagementTypeList("Lessor/Owner");
            tenancyRequestPage.ClickSaveCloseButton();

            //Assert record in search table
            tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestSearchPage.SetTenancyRequestSearchText(tenancyRequest);
            table = new Table(tenancyRequestSearchPage.GetSearchResultTable());
            StringAssert.Contains(table.GetCellValue("Name", tenancyRequest, "Tenancy Type"), "Residential Tenancy");

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }

        [TestMethod]
        [TestProperty("TestcaseID", "4399")]
        public void ATC4399_CRMDisplaydwellingtype()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "4399")
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

            TenancyRequestsSearchPage tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestSearchPage.ClickNewTenancyRequestButton();

            TenancyRequestPage tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();
            //tenancyRequestPage.PopulateTenancyRequestFormResidentialTenancy("1 THOMAS ST, BIRKDALE, QLD, 4159", "Residential Tenancy", "AMANDA TEST", "3", "AARON BALL", "700", "700", "Initial");
            tenancyRequestPage.PopulateTenancyRequestFormResidentialTenancy(
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("REQUEST_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGEMENT_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("NO_ROOMS")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString());
                
            //tenancyRequestPage.ClickSaveButton();

            ////Assert Dwelling Type field
            
            Assert.IsTrue(tenancyRequestPage.GetDwellingTypeText("Flat/Unit"));
            tenancyRequestPage.ClickPageTitle();
            Assert.IsTrue(tenancyRequestPage.GetDwellingTypeText("House"));
            tenancyRequestPage.ClickPageTitle(); 
            Assert.IsTrue(tenancyRequestPage.GetDwellingTypeText("Townhouse/Semi-Detached House"));
            tenancyRequestPage.ClickPageTitle(); 
            Assert.IsTrue(tenancyRequestPage.GetDwellingTypeText("Student Accommodation on Campus"));
            tenancyRequestPage.ClickPageTitle(); 
            Assert.IsTrue(tenancyRequestPage.GetDwellingTypeText("Moveable Dwelling/Site"));
            tenancyRequestPage.ClickPageTitle();
            Assert.IsTrue(tenancyRequestPage.GetDwellingTypeText("Moveable Dwelling/Site with electricity supplied and individually metered"));
            tenancyRequestPage.ClickPageTitle(); 
            Assert.IsTrue(tenancyRequestPage.GetDwellingTypeText("Studio"));
         
            ////Change Tenancy type
            //tenancyRequestPage.ClickTenancyTypeList();
            tenancyRequestPage.SetResidentialTenancyTypeList("Rooming Accommodation");
            tenancyRequestPage.SetResidentialManagementTypeList("Lessor/Owner");
            ////Assert Dwelling Type field
            Assert.IsTrue(tenancyRequestPage.GetDwellingTypeText("Boarding House"));
            tenancyRequestPage.ClickPageTitle(); 
            Assert.IsTrue(tenancyRequestPage.GetDwellingTypeText("Supported Accommodation"));
            tenancyRequestPage.ClickPageTitle(); 
            Assert.IsTrue(tenancyRequestPage.GetDwellingTypeText("Student Accommodation off Campus"));
            tenancyRequestPage.ClickPageTitle(); 
            Assert.IsTrue(tenancyRequestPage.GetDwellingTypeText("Studio"));

            //Assert Save Successful
            tenancyRequestPage.ClickSaveButton();
            StringAssert.Contains(tenancyRequestPage.GetRequestNumber(), "TR-BL-");

            string tenancyRequest = tenancyRequestPage.GetRequestNumber();
            tenancyRequestPage.ClickSaveCloseButton();

            //Assert record in search table
            tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestSearchPage.SetTenancyRequestSearchText(tenancyRequest);
            Table table = new Table(tenancyRequestSearchPage.GetSearchResultTable());
            StringAssert.Contains(table.GetCellValue("Name", tenancyRequest, "Tenancy Type"), "Rooming Accommodation");

            //Change Dwelling and check search table again
            table.ClickCellValue("Name", tenancyRequest, "Name");
            tenancyRequestPage = new TenancyRequestPage(driver);

            //tenancyRequestPage.ClickTenancyTypeList();
            tenancyRequestPage.SetResidentialTenancyTypeList("Residential Tenancy");
            tenancyRequestPage.SetResidentialManagementTypeList("Lessor/Owner");
            tenancyRequestPage.ClickSaveCloseButton();

            //Assert record in search table
            tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestSearchPage.SetTenancyRequestSearchText(tenancyRequest);
            table = new Table(tenancyRequestSearchPage.GetSearchResultTable());
            StringAssert.Contains(table.GetCellValue("Name", tenancyRequest, "Tenancy Type"), "Residential Tenancy");
           
            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }

        [TestMethod]
        [TestProperty("TestcaseID", "4400")]
        public void ATC4400_CRM1514EnterdataManagementTypes()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "4400")
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

            TenancyRequestsSearchPage tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestSearchPage.ClickNewTenancyRequestButton();

            TenancyRequestPage tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();
            tenancyRequestPage.PopulateTenancyRequestFormResidentialTenancy(
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("REQUEST_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString(), 
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGEMENT_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("NO_ROOMS")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString());
               
            //tenancyRequestPage.ClickSaveButton();

            ////Assert Management Type field
            Assert.IsTrue(tenancyRequestPage.GetTenancyManagementTypeListValue("Lessor/Owner"));
            Assert.IsTrue(tenancyRequestPage.GetTenancyManagementTypeListValue("Real Estate Agent"));
            Assert.IsTrue(tenancyRequestPage.GetTenancyManagementTypeListValue("Moveable dwelling owner/manager"));
            Assert.IsTrue(tenancyRequestPage.GetTenancyManagementTypeListValue("Community Housing Organisation"));
            //Assert.IsTrue(tenancyRequestPage.GetRTATenancyManagementTypeListValue("Other));
            
            ////Change Tenancy type
            tenancyRequestPage.SetResidentialTenancyTypeList("Rooming Accommodation");
            tenancyRequestPage.SetResidentialManagementTypeList("Lessor/Owner");
            ////Assert Dwelling Type field
            Assert.IsTrue(tenancyRequestPage.GetTenancyManagementTypeListValue("Lessor/Owner"));
            Assert.IsTrue(tenancyRequestPage.GetTenancyManagementTypeListValue("Real Estate Agent"));
            Assert.IsTrue(tenancyRequestPage.GetTenancyManagementTypeListValue("Manager/provider"));
            //Assert.IsTrue(tenancyRequestPage.GetRTATenancyManagementTypeListValue("Other));

            //Assert Save Successful
            tenancyRequestPage.ClickSaveButton();
            StringAssert.Contains(tenancyRequestPage.GetRequestNumber(), "TR-BL-");

            string tenancyRequest = tenancyRequestPage.GetRequestNumber();
            tenancyRequestPage.ClickSaveCloseButton();

            //Assert record in search table
            tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestSearchPage.SetTenancyRequestSearchText(tenancyRequest);
            Table table = new Table(tenancyRequestSearchPage.GetSearchResultTable());
            StringAssert.Contains(table.GetCellValue("Name", tenancyRequest, "Name"), tenancyRequest);
           
            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }

        [TestMethod]
        [TestProperty("TestcaseID", "4410")]
        public void ATC4410_CRMDisplaythefundedstatus()
        {
            User user = this.environment.GetUser(SecurityRole.SystemAdministrator);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();

            TenancyRequestsSearchPage tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);
            //Assert record in search table
            tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestSearchPage.SetTenancyRequestSearchText("BLAIR TEST");
            Table table = new Table(tenancyRequestSearchPage.GetSearchResultTable());
            table.ClickCellValue("Managing Party", "BLAIR TEST", "Name");

            TenancyRequestPage tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();
            StringAssert.Contains(tenancyRequestPage.GetPropertyDataControlModeRTAFundedStatus(), "locked");


        }


        [TestMethod]
        [TestProperty("TestcaseID", "4414")]
        public void ATC4414_CRMRoomingAccommodationneedstodefaultto1bedroom()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "4414")
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

            TenancyRequestsSearchPage tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestSearchPage.ClickNewTenancyRequestButton();

            TenancyRequestPage tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();
            //tenancyRequestPage.PopulateTenancyRequestFormResidentialTenancy("1 THOMAS ST, BIRKDALE, QLD, 4159", "Residential Tenancy", "AMANDA TEST", "3", "AARON BALL", "700", "700", "Initial");
            tenancyRequestPage.PopulateTenancyRequestFormRoomingAccomodation(
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("REQUEST_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGEMENT_TYPE")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString(),
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString());
            
            tenancyRequestPage.ClickSaveButton();

            
            StringAssert.Contains(tenancyRequestPage.GetNumberOfBedrooms(),"1");

            ////Change Tenancy type
           tenancyRequestPage.SetResidentialTenancyTypeList("Residential Tenancy");
            tenancyRequestPage.ClickSaveButton();
            StringAssert.Contains(tenancyRequestPage.GetNumberOfBedrooms(), "--");
            
            string tenancyRequest = tenancyRequestPage.GetRequestNumber();
            tenancyRequestPage.ClickSaveCloseButton();

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

        [TestMethod]
        [TestProperty("TestcaseID", "6905")]
        public void ATC6905_CRMTenancyrequestStatusReasonRemoveValidationSuccessfuldropdownvalue()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "6905")
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
            //Attemp to save with Mandatory data
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

            tenancyRequestPage.SetStatusReason("Ready for validation");
            tenancyRequestPage.ClickSaveButton();
            StringAssert.Contains(tenancyRequestPage.GetValidationStatusReason(), "Validation successful");           
            tenancyRequestPage.ClickSaveCloseButton();
            

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion


        }

            
        [TestMethod]
        [TestProperty("TestcaseID", "7629")]
        public void ATC7629_CRMGenerateEFTReferenceNumberForIndividualTR()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "7629")
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

            string managingParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();
            tenancyRequestPage.PopulateTenancyRequestValidationSuccessful(
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("REQUEST_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString(),
                managingParty, 
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
            tenancyRequestPage.ClickPageTitle();
            table = new Table(tenancyRequestPage.GetPaymentSummaryResultTable());

            //asserts if Payment Reference Number is >= 30000000
            string referencenumber = table.GetCellValue("Tenancy Request", tenancyrequest, "Reference Number");
            StringAssert.StartsWith(referencenumber, "3");
            Assert.IsTrue(referencenumber.Length == 8);
            Assert.IsTrue(int.Parse(referencenumber.Substring(1).TrimStart('0')) > 0);

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }


        [TestMethod]
        [TestProperty("TestcaseID", "7630")]
        public void ATC7630_TopUpBPayReferenceTRSuccessful()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "7630")
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

            //Attemp to save with Mandatory data
            string managingParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();
            tenancyRequestPage.PopulateTenancyRequestTopUpValidationSuccessful(
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("REQUEST_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString(),
                managingParty,
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
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY")].Value.ToString());

            tenancyRequestPage.ClickSaveButton();
            string tenancyrequest = tenancyRequestPage.GetRequestNumber();
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
            tenancyRequestPage.ClickPageTitle();
            table = new Table(tenancyRequestPage.GetPaymentSummaryResultTable());

            //asserts if Payment Reference Number is >= 200000000
            string referencenumber = table.GetCellValue("Tenancy Request", tenancyrequest, "Reference Number");
            StringAssert.StartsWith(referencenumber, "2");
            Assert.IsTrue(referencenumber.Length == 9);
            Assert.IsTrue(int.Parse(referencenumber.Substring(1).TrimStart('0')) > 0);

            StringAssert.Contains(tenancyRequestPage.GetStatusReason(), "Pending Financials");

            //asserts if record is read only
            Console.Write(tenancyRequestPage.GetRecordStatus());
            StringAssert.Equals(tenancyRequestPage.GetRecordStatus(), "Read only");

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }


        [TestMethod]
        [TestProperty("TestcaseID", "7632")]
        public void ATC7632_TopUpEFTReferenceTRSuccessful()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "7632")
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

            //Attemp to save with Mandatory data
            string managingParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();
            tenancyRequestPage.PopulateTenancyRequestTopUpValidationSuccessful(
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("REQUEST_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString(),
                managingParty,
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
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY")].Value.ToString());

            tenancyRequestPage.ClickSaveButton();
            string tenancyrequest = tenancyRequestPage.GetRequestNumber();
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value = tenancyrequest;
            tenancyRequestPage.SetStatusReason("Ready for validation");
            tenancyRequestPage.ClickSaveButton();
            StringAssert.Contains(tenancyRequestPage.GetValidationStatusReason(), "Validation successful");
            tenancyRequestPage.ClickSaveCloseButton();

            tenancyRequestsSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestsSearchPage.SetTenancyRequestSearchText(tenancyrequest);
            
            Assert.IsTrue(tenancyRequestsSearchPage.GetPaymentRefernceRefreshTable(tenancyrequest));
            Table table = new Table(tenancyRequestPage.GetPaymentSummaryResultTable());
            table.ClickCellValue("Name", tenancyrequest, "Name");

            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();
            table = new Table(tenancyRequestPage.GetPaymentSummaryResultTable());

            //asserts if Payment Reference Number is >= 30000000
            string referencenumber = table.GetCellValue("Tenancy Request", tenancyrequest, "Reference Number");
            StringAssert.StartsWith(referencenumber, "3");
            Assert.IsTrue(referencenumber.Length == 8);
            Assert.IsTrue(int.Parse(referencenumber.Substring(1).TrimStart('0')) > 0);

            StringAssert.Contains(tenancyRequestPage.GetStatusReason(), "Pending Financials");

            //asserts if record is read only
            Console.Write(tenancyRequestPage.GetRecordStatus());
            StringAssert.Equals(tenancyRequestPage.GetRecordStatus(), "Read only");

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }

        [TestMethod]
        [TestProperty("TestcaseID", "6736")]
        public void ATC6736_CRMPrepopulateTenancyfieldsfromknownpremisedetailsDwellingType()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "6736")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string managingParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();

            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();


            TenancyRequestsSearchPage tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);

            tenancyRequestSearchPage.ClickNewTenancyRequestButton();

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
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_START")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("ANTICIPATED_END")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString());

            tenancyRequestPage.ClickSaveButton();
            
            //Assert Dwelling Value
            StringAssert.Contains(tenancyRequestPage.GetDwellingType(), 
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("DWELLING_TYPE")].Value.ToString());
            tenancyRequestPage.ClickSaveCloseButton();

            
            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

        [TestMethod]
        [TestProperty("TestcaseID", "4457")]
        public void ATC4457_CRMPaymentTypefieldforTenancyRequestTestaudithistory()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "TR_TestData")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string managingParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();
            string tenancyRequestReference;
            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();


            TenancyRequestsSearchPage tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);

            tenancyRequestSearchPage.ClickNewTenancyRequestButton();

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
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_START")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("ANTICIPATED_END")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString());

            tenancyRequestPage.ClickSaveButton();
            tenancyRequestReference = tenancyRequestPage.GetRequestNumber();

            //Review Audit
            homePage.HoverTRRibbonTab(tenancyRequestReference);
            homePage.ClickAuditRibbonButton();

            tenancyRequestPage = new TenancyRequestPage(driver);
            Table table = new Table(tenancyRequestPage.GetAuditHistoryTable());
            string todaysDate = DateTime.Now.ToString("d/MM/yyyy");
            StringAssert.Contains(table.GetCellContainsValue("Event", "Create", "Changed Date"), todaysDate);

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

        [TestMethod]
        [TestProperty("TestcaseID", "4459")]
        public void ATC4459_CRMPaymentTypefieldforTenancyRequestTestaudithistory()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "TR_TestData")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string managingParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();
            string tenancyRequestReference;
            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();


            TenancyRequestsSearchPage tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);

            tenancyRequestSearchPage.ClickNewTenancyRequestButton();

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
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_START")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("ANTICIPATED_END")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString());

            tenancyRequestPage.ClickSaveButton();
            tenancyRequestReference = tenancyRequestPage.GetRequestNumber();

            //Change Tenancy Type
            tenancyRequestPage.SetResidentialTenancyTypeList("Rooming Accommodation");
            tenancyRequestPage.ClickSaveButton();
            StringAssert.Contains(tenancyRequestPage.GetNumberOfBedrooms(), "1");
            StringAssert.Contains(tenancyRequestPage.GetNumberOfBedroomsProperty("data-controlmode"), "locked");

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

       

        [TestMethod]
        [TestProperty("TestcaseID", "4462")]
        public void ATC4462_CRMPaymentTypefieldforTenancyRequestTestaudithistory()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "TR_TestData")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string managingParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();
            string tenancyRequestReference;
            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();


            TenancyRequestsSearchPage tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);

            tenancyRequestSearchPage.ClickNewTenancyRequestButton();

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
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_START")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("ANTICIPATED_END")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString());

            tenancyRequestPage.SetResidentialTenancyTypeList("Rooming Accommodation");
            tenancyRequestPage.SetResidentialManagementTypeList("Lessor/Owner");
            tenancyRequestPage.ClickSaveButton();
            tenancyRequestReference = tenancyRequestPage.GetRequestNumber();

            //Change Tenancy Type
            
            StringAssert.Contains(tenancyRequestPage.GetNumberOfBedrooms(), "1");
            StringAssert.Contains(tenancyRequestPage.GetNumberOfBedroomsProperty("data-controlmode"), "locked");

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

        [TestMethod]
        [TestProperty("TestcaseID", "4483")]
        public void ATC4483_CRMAutomaticallyrecordapaymentreferencenumberforthelodgement()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "TR_TestData")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string managingParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();
            //string tenancyRequestReference;
            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();


            TenancyRequestsSearchPage tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);

            tenancyRequestSearchPage.ClickNewTenancyRequestButton();

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
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_START")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("ANTICIPATED_END")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString());

            tenancyRequestPage.ClickSaveButton();
            string tenancyrequest = tenancyRequestPage.GetRequestNumber();
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value = tenancyrequest;

            tenancyRequestPage.SetStatusReason("Ready for validation");
            tenancyRequestPage.ClickSaveButton();
            StringAssert.Contains(tenancyRequestPage.GetValidationStatusReason(), "Validation successful");
            tenancyRequestPage.ClickSaveCloseButton();

            tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestSearchPage.SetTenancyRequestSearchText(tenancyrequest);

            Assert.IsTrue(tenancyRequestSearchPage.GetPaymentRefernceRefreshTable(tenancyrequest));

            Table table = new Table(tenancyRequestSearchPage.GetSearchResultTable());
            table.ClickCellValue("Name", tenancyrequest, "Name");

            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();
            StringAssert.Contains(tenancyRequestPage.GetFundedStatus(), "Payment pending");
            

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

        [TestMethod]
        [TestProperty("TestcaseID", "4489")]
        public void ATC4489_CRMForm2LodgementRequestBatch()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "4489")
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

            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();

            homePage.ClickRequestBatchRibbonButton();

            RequestBatchesSearchPage requestBatchSearchPage = new RequestBatchesSearchPage(driver);
            requestBatchSearchPage.ClickNewRequestBatchButton();

            RequestBatchPage requestBatchPage = new RequestBatchPage(driver);
            requestBatchPage.ClickPageTitle();
            requestBatchPage.SetManagingPartyText(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString());
            requestBatchPage.SetPaymentType(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString());
            requestBatchPage.ClickSaveButton();
            string requestBatch = requestBatchPage.GetRequestNumber();
            StringAssert.Contains(requestBatch, "TRB-BL-");

            //Add tenancy requests
            
            string TR1reference = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value.ToString();
            string TR2reference = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MISC1")].Value.ToString();

            requestBatchPage.ClickAddAssociatedRequestsButton();
            requestBatchPage.SetAssociatedRequest(TR1reference);
            requestBatchPage.ClickAddAssociatedRequestsButton();
            requestBatchPage.SetAssociatedRequest(TR2reference);
            StringAssert.Contains(requestBatchPage.GetDialogErrorMessageText(), "The record is already associated with another record.");
            requestBatchPage.ClickErrorMessageOkButton();

            
            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }


        [TestMethod]
        [TestProperty("TestcaseID", "4492")]
        public void ATC4492_CRMBPayreferenceinRequestBatch ()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "TR_TestData")
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

            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();

            homePage.ClickRequestBatchRibbonButton();

            RequestBatchesSearchPage requestBatchSearchPage = new RequestBatchesSearchPage(driver);
            requestBatchSearchPage.ClickNewRequestBatchButton();

            RequestBatchPage requestBatchPage = new RequestBatchPage(driver);
            requestBatchPage.ClickPageTitle();
            requestBatchPage.ClickPageTitle();
            requestBatchPage.SetManagingPartyText(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString());
            requestBatchPage.SetPaymentType(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString());
            requestBatchPage.ClickSaveButton();
            string requestBatch = requestBatchPage.GetRequestNumber();
            StringAssert.Contains(requestBatch, "TRB-BL-");

            //Add tenancy requests
            string TR1reference = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value.ToString();
            requestBatchPage.ClickAddAssociatedRequestsButton();
            requestBatchPage.SetAssociatedRequest(TR1reference);

            requestBatchPage.SetStatusReason("Ready for validation");
            tenancyRequestPage.ClickSaveButton();
            StringAssert.Contains(tenancyRequestPage.GetValidationStatusReason(), "Validation successful");
            tenancyRequestPage.ClickSaveCloseButton();

            requestBatchSearchPage = new RequestBatchesSearchPage(driver);
            requestBatchSearchPage.SetRequestBatchSearchText(requestBatch);
            Table table = new Table(requestBatchSearchPage.GetSearchResultTable());
            Assert.IsTrue(requestBatchSearchPage.GetPaymentRefernceRefreshTable(requestBatch));
            table = new Table(requestBatchSearchPage.GetSearchResultTable()); 
            table.ClickCellValue("Name", requestBatch, "Name");

            requestBatchPage = new RequestBatchPage(driver);
            table = new Table(requestBatchPage.GetPaymentSummaryResultTable());
            table.GetCellValue("Request Batch", requestBatch, "Reference Number");
            


            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }

        [TestMethod]
        [TestProperty("TestcaseID", "5506")]
        public void ATC5506_CRMAutomaticallyrecordapaymentreferencenumberforthelodgement()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "TR_TestData")
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

            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();

            homePage.ClickRequestBatchRibbonButton();

            RequestBatchesSearchPage requestBatchSearchPage = new RequestBatchesSearchPage(driver);
            requestBatchSearchPage.ClickNewRequestBatchButton();

            RequestBatchPage requestBatchPage = new RequestBatchPage(driver);
            requestBatchPage.ClickPageTitle();
            requestBatchPage.ClickPageTitle();
            requestBatchPage.SetManagingPartyText(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString());
            requestBatchPage.SetPaymentType(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString());
            requestBatchPage.ClickSaveButton();
            string requestBatch = requestBatchPage.GetRequestNumber();
            StringAssert.Contains(requestBatch, "TRB-BL-");

            //Add tenancy requests
            string TR1reference = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value.ToString();
            requestBatchPage.ClickAddAssociatedRequestsButton();
            requestBatchPage.SetAssociatedRequest(TR1reference);

            requestBatchPage.SetStatusReason("Ready for validation");
            tenancyRequestPage.ClickSaveButton();
            StringAssert.Contains(tenancyRequestPage.GetValidationStatusReason(), "Validation successful");
            tenancyRequestPage.ClickSaveCloseButton();

            requestBatchSearchPage = new RequestBatchesSearchPage(driver);
            requestBatchSearchPage.SetRequestBatchSearchText(requestBatch);
            Table table = new Table(requestBatchSearchPage.GetSearchResultTable());
            Assert.IsTrue(requestBatchSearchPage.GetPaymentRefernceRefreshTable(requestBatch));
            table = new Table(requestBatchSearchPage.GetSearchResultTable());
            table.ClickCellValue("Name", requestBatch, "Name");

            requestBatchPage = new RequestBatchPage(driver);
            table = new Table(requestBatchPage.GetPaymentSummaryResultTable());
            table.GetCellValue("Request Batch", requestBatch, "Reference Number");

            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();


            tenancyRequestsSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestsSearchPage.SetTenancyRequestSearchText(tenancyrequest);
            table = new Table(tenancyRequestsSearchPage.GetSearchResultTable());
            table.ClickCellValue("Name", tenancyrequest, "Name");
            tenancyRequestPage = new TenancyRequestPage(driver);

            StringAssert.Contains(tenancyRequestPage.GetFundedStatus(), "Payment pending");

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

        [TestMethod]
        [TestProperty("TestcaseID", "5508")]
        public void ATC5508_CRMAutomaticallyrecordapaymentreferencenumberforthelodgementValidationUnsuccessful()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "5508")
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

            tenancyRequestPage.SetStatusReason("Ready for validation");
            tenancyRequestPage.ClickSaveButton();
            StringAssert.Contains(tenancyRequestPage.GetValidationStatusReason(),"Validation failed");
            StringAssert.Contains(tenancyRequestPage.GetFundedStatus(), "--");
            tenancyRequestPage.ClickSaveCloseButton();

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }

        [TestMethod]
        [TestProperty("TestcaseID", "5511")]
        public void ATC5511_CRMBPayReferenceNewLodgement()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "5511")
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

            string initialRequestParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString();
            string initialContribution = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString();

            //Create New Tenancy Request(Bond Lodgement)
            TenancyRequestsSearchPage tenancyRequestsSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestsSearchPage.ClickNewTenancyRequestButton();

            tenancyRequestsSearchPage.ClickNewTenancyRequestButton();

            TenancyRequestPage tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();
                        
            tenancyRequestPage.PopulateMandatoryFieldValues(
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
                       
            //Create New Batch Request from Tenancy Reuest Page and confirm batch number format correct
            string requestBatch = tenancyRequestPage.CreateNewBatchRequest(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString(), MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString());
            StringAssert.Contains(requestBatch, "TRB-BL-");

            tenancyRequestPage.ClickSaveButton();
            string tenancyrequest = tenancyRequestPage.GetRequestNumber();
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value = tenancyrequest;
            StringAssert.Contains(tenancyrequest, "TR-BL-");

            tenancyRequestPage.ClickSaveCloseButton();

            tenancyRequestPage = new TenancyRequestPage(driver);
             
            //Open RBS>New BatchRequest 
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();

            homePage.ClickRequestBatchRibbonButton();

            //Confirm Tenancy request listed in the New Batch Request
            RequestBatchesSearchPage requestBatchSearchPage = new RequestBatchesSearchPage(driver);
            requestBatchSearchPage.SetRequestBatchSearchText(requestBatch);

            Table searchTable = new Table(requestBatchSearchPage.GetSearchResultTable());
            searchTable.ClickCellValue("Name", requestBatch, "Name");

            RequestBatchPage requestBatchPage = new RequestBatchPage(driver);
            requestBatchPage.ClickPageTitle();

            Table trTable = new Table(requestBatchPage.GetTenancyRequestTable());

            
            string actualTenancyRequest = trTable.GetCellValue("Managing Party", MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString(),"Name");
            Assert.AreEqual(tenancyrequest, actualTenancyRequest,"Validating the Tenancy request listed in NewBatch Request");
           
            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }


        [TestMethod]
        [TestProperty("TestCaseID", "9857")]
        public void ATC9857_CRMRBSClaimsUserCannotTouchBLTR()
        {
            string requestType = "Bond Lodgement";

            //Login as RBS Claims Officer user role.
            User user = this.environment.GetUser(SecurityRole.RBSClaimsOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            //Navigate to: Rental Bond Services > Tenancy Requests
            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();

            TenancyRequestsSearchPage tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);

            //Click [+NEW] button. 
            tenancyRequestSearchPage.ClickNewTenancyRequestButton();

            TenancyRequestPage tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            //Inspect {Type} field dropdown values.
            //Assert:Dropdown item "Bond Lodgement" should not be available.
            
            Assert.IsFalse(tenancyRequestPage.isRequestType(requestType), "We expected Request Type: " + requestType + ", to not be in the list but we found it.");

        }

        [TestMethod]
        [TestProperty("TestCaseID","9851")]
        public void ATC9851_CRMChangeOfMPMandatoryFieldsValidationsAC3NoContributorsExceptionOverride()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "9851")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion
            
            //Login as RBS Claims Officer user role.
            User user = this.environment.GetUser(SecurityRole.RBSClaimsOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            //Navigate to: Rental Bond Services > Tenancy Requests
            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();

            TenancyRequestsSearchPage tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);

            //Click [+NEW] button.
            tenancyRequestSearchPage.ClickNewTenancyRequestButton();

            TenancyRequestPage tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            //Set {Type} field to Change of Managing Party.
            //
            tenancyRequestPage.SetRequestTypeListValue("Change of Managing Party");

            //Populate all other mandatory fields required for successful validation (as per stories 6792/6793) without a Request Party
            tenancyRequestPage.PopulateTenancyRequestChangeOfManagingPartyWithoutRequestParty(
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("REQUEST_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MISC1")].Value.ToString(),
                DateTime.Now.ToString("dd/MM/yyyy"),
                DateTime.Now.ToString("dd/MM/yyyy"),
                DateTime.Now.ToString("dd/MM/yyyy"),
                DateTime.Now.ToString("dd/MM/yyyy"),
                "Signature verified",
                "Signature verified");

            //Click [Save] button to unlock {Status Reason}
            tenancyRequestPage.ClickSaveButton();

            //Set {Status Reason} to "Ready for validation".
            tenancyRequestPage.SetStatusReason("Ready for validation");

            //Click [Save] button.
            tenancyRequestPage.ClickSaveButton();

            //Assert: {Status Reason} becomes "Validation failed".
            StringAssert.Equals(tenancyRequestPage.GetStatusReason(), "Validation failed");

            //Inspect Queue Reasons section.
            tenancyRequestPage.ClickQueueReasons();  //This may not be necessary.

            //Assert: An exception queue entry is created with reason type: "At least one contributor required"
            Table table = new Table(tenancyRequestPage.GetQueueReasonTable());
            StringAssert.Contains(table.GetCellContainsValue("Reason", "Invalid number of contributors", "Reason"), "Invalid number of contributors");
            
            //Double-click on exception queue record.
            //UICommon.DoubleClickElement(table.GetCellElementContainsValue("Reason", "Invalid Contributor : At least one contributor required", "Status Reason"), driver);
            table.ClickCellContainsValueEnterRow("Reason", "Invalid number of contributors", "Status Reason");
            TenancyRequestQueueReasonPage tenancyRequestQueueReasonPage = new TenancyRequestQueueReasonPage(driver);

            //Click {Override} checkbox to select/tick it.
            tenancyRequestQueueReasonPage.SetOverrideCheckBox(true);

            //Click [SAVE & CLOSE] button.
            tenancyRequestQueueReasonPage.ClickSaveCloseButton();
            tenancyRequestPage = new TenancyRequestPage(driver);

            //Set {Status Reason} to "Ready for validation".
            tenancyRequestPage.SetStatusReason("Ready for validation");

            //Click [Save] button.
            tenancyRequestPage.ClickSaveButton();

            string tenancyRequest = tenancyRequestPage.GetRequestNumber();
            tenancyRequestPage.ClickSaveCloseButton();

            tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestSearchPage.SetTenancyRequestSearchText(tenancyRequest);

            table = new Table(tenancyRequestSearchPage.GetSearchResultTable());
            table.GetCellValue("", "", "");
            //Assert: {Status Reason} becomes "Validation successful".
            //StringAssert.Contains(tenancyRequestPage.GetStatusReason(), "Validation successful");

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

        [TestMethod]
        [TestProperty("TestID", "5512")]
        public void ATC5512_CRMRaiseExceptionForLodgementExceedingBond()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "5512")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion
            
            //Login as RBS user
            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            //Navigate to: Rental Bond Services > Tenancy Requests
            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();

            TenancyRequestsSearchPage tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);

            //Click [+NEW] button.
            tenancyRequestSearchPage.ClickNewTenancyRequestButton();

            TenancyRequestPage tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            //Set {Type} field to Bond Lodgement.
            tenancyRequestPage.SetRequestTypeListValue("Bond Lodgement");

            //Enter details
            tenancyRequestPage.PopulateTenancyRequestFormBondLodgement(
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("REQUEST_TYPE")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_TYPE")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGEMENT_TYPE")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("WEEKLY_RENT")].Value.ToString(), 
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString(), 
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_START")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("ANTICIPATED_END")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MISC1")].Value.ToString()); //Rental subsidy

            //Save Tenancy Request
            tenancyRequestPage.ClickSaveButton();

            //Update Status Reason to 'Ready for validation' and Save
            tenancyRequestPage.SetStatusReason("Ready for validation");
            tenancyRequestPage.ClickSaveButton();
            StringAssert.Contains(tenancyRequestPage.GetValidationStatusReason(), "Validation successful");
            
            //Confirm Queue reason for excess bond not added
            tenancyRequestPage.ClickQueueReasons();
            Table table = new Table(tenancyRequestPage.GetQueueReasonTable());
            Assert.IsTrue(table.GetNoRecordsInTable());

            //Confirm Maximum Allowed Bond set correctly
            //Assert: Maximum Allowed Bond set to $400
            StringAssert.Contains(tenancyRequestPage.GetMaximumAllowedAmount(), "400");

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

        [TestMethod]
        [TestProperty("TestCaseID", "6708")]
        public void ATC6708_CRMDistributeAmountForContributors()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "6708")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string secondContributor = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MISC1")].Value.ToString();
            string thridContributor = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MISC2")].Value.ToString();

            //Login as RBS Operations Standard user role.
            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            //Navigate to: Rental Bond Services > Tenancy Requests
            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();

            //Create new Tenancy Request with mandatory fields populated
            TenancyRequestsSearchPage tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);

            tenancyRequestSearchPage.ClickNewTenancyRequestButton();

            TenancyRequestPage tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            string initialRequestParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString();
            string initialContribution = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString();
            
             tenancyRequestPage.PopulateMandatoryFieldValues(
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

            //Save Tenancy Record
            tenancyRequestPage.ClickSaveButton();
            string tenancyRequest = tenancyRequestPage.GetRequestNumber();

            //Inspect Screen for Warnings "Warning displays that {Amount Paid with Lodgement} field <> sum of contributors."
            string warningMessage = tenancyRequestPage.GetWarningMessage();
            StringAssert.Contains(warningMessage, "Sum of Request does not equal the amount entered for Amount Paid with Lodgement","Validating the warning appears");

            //Click the ribbon button Distribute and Save the record. Warnings should disappear
            string alertMessage = tenancyRequestPage.ClickDistributeButton();
            StringAssert.Contains(alertMessage, "Distribution process has been initiated");
            
            //TODO:This is an issue an extra save dialog pops up and clicking OK button on this. 
            string saveMessage = tenancyRequestPage.GetAlertMessage();
            StringAssert.Contains(saveMessage, "Your changes have not been saved");
                       

            warningMessage = tenancyRequestPage.GetWarningMessage();
            Assert.AreEqual(warningMessage, "","Validating that warning message disappeared");

            //Add another contributor and populate amount field with > $1. Save Record. Warning should appear.
            string BaseWindow = driver.CurrentWindowHandle; //Records the current window handle
            tenancyRequestPage.ClickAddNewRequestPartyImage();
            Thread.Sleep(2000);
  
            driver = tenancyRequestPage.SwitchNewBrowser(driver, BaseWindow);

            TenancyRequestPartyPage tenancyRequestPartyPage = new TenancyRequestPartyPage(driver);
            tenancyRequestPartyPage.ClickPageTitle();
            tenancyRequestPartyPage.SetClientNameValue(secondContributor);
            tenancyRequestPartyPage.SetAmountValue("100");
            tenancyRequestPartyPage.ClickSaveCloseButton();
            
            driver = driver.SwitchTo().Window(BaseWindow);
                       
            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            tenancyRequestPage.ClickSaveButton();

           //TODO:Issue the waning does not appear without refreshing the record.This is a workaround.
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();

            tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestSearchPage.SetTenancyRequestSearchText(tenancyRequest);
            Table table = new Table(tenancyRequestSearchPage.GetSearchResultTable());
            StringAssert.Contains(table.GetCellValue("Name", tenancyRequest, "Status Reason"), "New");

            table.ClickCellValue("Name", tenancyRequest, "Name");

            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            //Inspect warning appears
            warningMessage = tenancyRequestPage.GetWarningMessage();
            StringAssert.Contains(warningMessage, "Sum of Request does not equal the amount entered for Amount Paid with Lodgement","Validating that the warning appears");

            //Add another contributor and populate amount field with > $1. Save Record. Warning should appear.
            BaseWindow = driver.CurrentWindowHandle; //Records the current window handle
            tenancyRequestPage.ClickAddNewRequestPartyImage();
            Thread.Sleep(2000);
            driver = tenancyRequestPage.SwitchNewBrowser(driver, BaseWindow);

            tenancyRequestPartyPage = new TenancyRequestPartyPage(driver);
            tenancyRequestPartyPage.ClickPageTitle();
            tenancyRequestPartyPage.SetClientNameValue(thridContributor);
            tenancyRequestPartyPage.SetAmountValue("100");
            tenancyRequestPartyPage.ClickSaveCloseButton();
            

            driver = driver.SwitchTo().Window(BaseWindow);
            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            tenancyRequestPage.ClickSaveButton();
     
            //Inspect warning appears
            warningMessage = tenancyRequestPage.GetWarningMessage();
            StringAssert.Contains(warningMessage, "Sum of Request does not equal the amount entered for Amount Paid with Lodgement","Validating that the warning appears");

            //Click the ribbon button Distribute and Save the record. Warning should disappear.
            alertMessage = tenancyRequestPage.ClickDistributeButton();
            StringAssert.Contains(alertMessage, "Distribution process has been initiated");
                        
            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickSaveButton();

            //Issue the waning does disappear not appear without refreshing the record.This is a workaround.
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();

            tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestSearchPage.SetTenancyRequestSearchText(tenancyRequest);
            table = new Table(tenancyRequestSearchPage.GetSearchResultTable());
            StringAssert.Contains(table.GetCellValue("Name", tenancyRequest, "Status Reason"), "New");

            table.ClickCellValue("Name", tenancyRequest, "Name");

            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            //Warning disappears
            warningMessage = tenancyRequestPage.GetWarningMessage();
            Assert.AreEqual(warningMessage, "", "Validating that warning message disappeared");

            //Modify amount paid with lodgemnt to $1500 and inspect warning appears
            tenancyRequestPage.SetAmountPaidWithLodgement("1500");
            tenancyRequestPage.ClickSaveButton();
                        
            warningMessage = tenancyRequestPage.GetWarningMessage();
            StringAssert.Contains(warningMessage, "Sum of Request does not equal the amount entered for Amount Paid with Lodgement","Validating that the warning appears");

            //TODO:Modify Contibutor amount to be different but still adds up to amount. Inspect warning disappears
            Table requestPartyTable = new Table(tenancyRequestPage.GetRequestPartyTable());
            StringAssert.Contains(requestPartyTable.GetCellContainsValue("Client", secondContributor, "Request Amount"), "400");
            
            BaseWindow = driver.CurrentWindowHandle;
            UICommon.DoubleClickElement(requestPartyTable.GetCellElementContainsValue("Client", secondContributor, "Client"), driver);
            Thread.Sleep(1000);

            driver = tenancyRequestPage.SwitchNewBrowser(driver, BaseWindow);
            
            tenancyRequestPartyPage = new TenancyRequestPartyPage(driver);
            tenancyRequestPartyPage.ClickPageTitle();

            tenancyRequestPartyPage.SetAmountValue("700");
            tenancyRequestPage.ClickSaveCloseButton();

            driver = driver.SwitchTo().Window(BaseWindow);

            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            tenancyRequestPage.ClickSaveButton();

            //TODO.WArnings does not disappear without refreshing the record. This is a workaround.
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();

            tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestSearchPage.SetTenancyRequestSearchText(tenancyRequest);
            table = new Table(tenancyRequestSearchPage.GetSearchResultTable());
            StringAssert.Contains(table.GetCellValue("Name", tenancyRequest, "Status Reason"), "New");

            table.ClickCellValue("Name", tenancyRequest, "Name");

            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            //Warning disappears
            warningMessage = tenancyRequestPage.GetWarningMessage();
            Assert.AreEqual(warningMessage, "", "Validating that warning message disappeared");

            //Modify amount paid with lodgement to be $900
            tenancyRequestPage.SetAmountPaidWithLodgement("900");
            tenancyRequestPage.ClickSaveButton();

            //Warning appears
            warningMessage = tenancyRequestPage.GetWarningMessage();
            StringAssert.Contains(warningMessage, "Sum of Request does not equal the amount entered for Amount Paid with Lodgement","Validating that the warning appesrs");

            //Click Distribute and save the record. Warning should disappear
            tenancyRequestPage.ClickDistributeButton();
            tenancyRequestPage.ClickSaveButton();

            warningMessage = tenancyRequestPage.GetWarningMessage();
            Assert.AreEqual(warningMessage, "", "Validating that warning message disappeared");

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

        [TestMethod]
        [TestProperty("TestCaseID", "6908")]
        public void ATC6908_CRMUpdateTenancyRequestPaymentTypeExceptionQueue()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "TR_TestData")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion
            
            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            //Navigate to: Rental Bond Services > Tenancy Requests
            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();

            //Create new Tenancy Request with fields populated except Payment type
            TenancyRequestsSearchPage tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);

            tenancyRequestSearchPage.ClickNewTenancyRequestButton();

            TenancyRequestPage tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            tenancyRequestPage.PopulateTenancyRequestWithoutPaymentType(
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
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("ANTICIPATED_END")].Value.ToString());
                     
            //Record should save successfully
            tenancyRequestPage.ClickSaveButton();

            string tenancyRequest = tenancyRequestPage.GetRequestNumber();
            StringAssert.Contains(tenancyRequest, "TR-BL-","Validating the Tenancy Request saved with the correct TR no fromat");

            tenancyRequestPage.ClickSaveCloseButton();

            //Open the record again and change the Status Reason to Ready for validation and save
            tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestSearchPage.SetTenancyRequestSearchText(tenancyRequest);
            Table table = new Table(tenancyRequestSearchPage.GetSearchResultTable());
            StringAssert.Contains(table.GetCellValue("Name", tenancyRequest, "Status Reason"), "New");

            table.ClickCellValue("Name", tenancyRequest, "Name");

            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            tenancyRequestPage.SetStatusReason("Ready for validation");
            tenancyRequestPage.ClickSaveButton();
            StringAssert.Contains(tenancyRequestPage.GetValidationStatusReason(), "Validation failed","Validation should fail when payment type is balnk");

            //There should be one exception in the queue for Pyment type missing
            tenancyRequestPage.ClickQueueReasons();
            Table queueTable = new Table(tenancyRequestPage.GetQueueReasonTable());
            StringAssert.Contains(queueTable.GetCellContainsValue("Reason", "Payment Type is blank", "Reason"), "Payment Type is blank");

            //Open the exception queue and change the reason to resolved. In the tenancy record select a payment type and change the status reason to Ready for validation - SAVE
            UICommon.DoubleClickElement(queueTable.GetCellElementContainsValue("Reason", "Payment Type is blank", "Status Reason"), driver);
            
            TenancyRequestQueueReasonStatusPage tenancyQueueReasonStatusPage = new TenancyRequestQueueReasonStatusPage(driver);

            StringAssert.Contains(tenancyQueueReasonStatusPage.GetReasonValue(), "Payment Type is blank");
            tenancyQueueReasonStatusPage.SetReasonValue("Resolved");
            tenancyQueueReasonStatusPage.ClickSaveCloseButton();

            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.SetPaymentType("BPay");

            tenancyRequestPage.ClickSaveButton();
            //Record should be saved and validation should be successful
            tenancyRequestPage.SetStatusReason("Ready for validation");
            tenancyRequestPage.ClickSaveButton();
            StringAssert.Contains(tenancyRequestPage.GetValidationStatusReason(), "Validation successful","Validation should be successful after adding payment type");

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }

        [TestMethod]
        [TestProperty("TestcaseID", "6907")]
        public void ATC6907_CRMOverrideInvalidKeywordExceptionQueue()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "6907")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            //Navigate to: Rental Bond Services > Tenancy Requests
            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();

            //Create new Tenancy Request. Fillup the Rental Premises with some address which has keyword like garage
            TenancyRequestsSearchPage tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);

            tenancyRequestSearchPage.ClickNewTenancyRequestButton();

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
            
            tenancyRequestPage.SetDwellingTypeListValue("House");
            
            tenancyRequestPage.ClickSaveButton();

            string tenancyRequest = tenancyRequestPage.GetRequestNumber();
            StringAssert.Contains(tenancyRequest, "TR-BL-", "Validating the Tenancy Request saved with the correct TR no fromat");

            tenancyRequestPage.ClickSaveCloseButton();

            //Reopen the record and change the status reason to Ready for Validation and save.The record should show an exception with Invalid Keyword
            tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestSearchPage.SetTenancyRequestSearchText(tenancyRequest);
            Table table = new Table(tenancyRequestSearchPage.GetSearchResultTable());
            StringAssert.Contains(table.GetCellValue("Name", tenancyRequest, "Status Reason"), "New");

            table.ClickCellValue("Name", tenancyRequest, "Name");

            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            tenancyRequestPage.SetStatusReason("Ready for validation");
            tenancyRequestPage.ClickSaveButton();
            StringAssert.Contains(tenancyRequestPage.GetValidationStatusReason(), "Validation failed", "Validation should fail when invalid keyword entered in address details");

            tenancyRequestPage.ClickQueueReasons();
            Table queueTable = new Table(tenancyRequestPage.GetQueueReasonTable());
            StringAssert.Contains(queueTable.GetCellContainsValue("Reason", "Invalid keyword detected", "Reason"), "Invalid keyword detected");
                        
            UICommon.DoubleClickElement(queueTable.GetCellElementContainsValue("Reason", "Invalid keyword detected", "Status Reason"),driver);

            TenancyRequestQueueReasonStatusPage tenancyQueueReasonStatusPage = new TenancyRequestQueueReasonStatusPage(driver);
            Assert.IsFalse(tenancyQueueReasonStatusPage.GetOverrideCheckBoxValue(),"Validating the override checkbox for Invalid keyword queue reason is available for RBS officer");

           //Login with any other user other than RBS operation user and see the override can be done by that user

            driver.Close();
            driver = null;
            this.TestSetup();

            user = this.environment.GetUser(SecurityRole.GeneralStaff);
            new LoginDialog().Login(user.Id, user.Password);

            //Navigate to: Rental Bond Services > Tenancy Requests
            homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();

            tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);

            tenancyRequestSearchPage.SetTenancyRequestSearchText(tenancyRequest);
            table = new Table(tenancyRequestSearchPage.GetSearchResultTable());
            
            table.ClickCellValue("Name", tenancyRequest, "Name");

            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            tenancyRequestPage.ClickQueueReasons();
            queueTable = new Table(tenancyRequestPage.GetQueueReasonTable());
            StringAssert.Contains(queueTable.GetCellContainsValue("Reason", "Invalid keyword detected", "Reason"), "Invalid keyword detected");

            //Open the exception queue and change the reason to resolved. In the tenancy record select a payment type and change the status reason to Ready for validation - SAVE
            UICommon.DoubleClickElement(queueTable.GetCellElementContainsValue("Reason", "Invalid keyword detected", "Status Reason"), driver);

            tenancyQueueReasonStatusPage = new TenancyRequestQueueReasonStatusPage(driver);
            tenancyQueueReasonStatusPage.SetOverrideCheckBox(true);
            
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();

            StringAssert.Contains(tenancyRequestPage.GetAlertMessage(), "Your changes have not been saved");

            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();

            //Create new Tenancy Request with fields populated except Payment type
            tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);

            tenancyRequestSearchPage.SetTenancyRequestSearchText(tenancyRequest);
            table = new Table(tenancyRequestSearchPage.GetSearchResultTable());
           
            table.ClickCellValue("Name", tenancyRequest, "Name");

            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            tenancyRequestPage.ClickQueueReasons();
            queueTable = new Table(tenancyRequestPage.GetQueueReasonTable());
            StringAssert.Contains(queueTable.GetCellContainsValue("Reason", "Invalid keyword detected", "Override"), "No");

            //Once override is done for that tenancy record for Invalid Keyword exception, it should appear as resolved in the queue in subsequent validations.
            //Validation should be successful
            driver.Close();
            driver = null;

            this.TestSetup();

            user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            //Navigate to: Rental Bond Services > Tenancy Requests
            homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();

            tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);

            tenancyRequestSearchPage.SetTenancyRequestSearchText(tenancyRequest);
            table = new Table(tenancyRequestSearchPage.GetSearchResultTable());
            
            table.ClickCellValue("Name", tenancyRequest, "Name");

            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            tenancyRequestPage.ClickQueueReasons();
            queueTable = new Table(tenancyRequestPage.GetQueueReasonTable());
            StringAssert.Contains(queueTable.GetCellContainsValue("Reason", "Invalid keyword detected", "Reason"), "Invalid keyword detected");

            UICommon.DoubleClickElement(queueTable.GetCellElementContainsValue("Reason", "Invalid keyword detected", "Status Reason"), driver);

            tenancyQueueReasonStatusPage = new TenancyRequestQueueReasonStatusPage(driver);
            tenancyQueueReasonStatusPage.SetOverrideCheckBox(true);
            tenancyQueueReasonStatusPage.ClickSaveCloseButton();

            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            tenancyRequestPage.ClickQueueReasons();
            queueTable = new Table(tenancyRequestPage.GetQueueReasonTable());
            StringAssert.Contains(queueTable.GetCellContainsValue("Reason", "Invalid keyword detected", "Override"), "Yes");

            tenancyRequestPage.SetStatusReason("Ready for validation");
            tenancyRequestPage.ClickSaveButton();
            StringAssert.Contains(tenancyRequestPage.GetValidationStatusReason(), "Validation successful");
                        
            queueTable = new Table(tenancyRequestPage.GetQueueReasonTable());
            StringAssert.Contains(queueTable.GetCellContainsValue("Reason", "Invalid keyword detected", "Status Reason"), "Resolved");

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }

        [TestMethod]
        [TestProperty("TestCaseID", "4422")]
        public void ATC4422_CRMTenancyRequestWithNoContributor()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "TR_TestData")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion
                        
            string initialRequestParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString();
            string initialContribution = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString();
            string rentalPremises = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString();

            //Login as RBS user
            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            //As a RBS Officer, navigate to an Address Detail
            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaAddressDetailRibbonButton();

            AddressDetailSearchPage addressDetailSearchPage = new AddressDetailSearchPage(driver);
            addressDetailSearchPage.SetAddressDetailSearchText("MITCHELTON");

            Table addressSearchTable = new Table(addressDetailSearchPage.GetSearchResultTable());
            addressSearchTable.ClickCellContainsValue("Name", rentalPremises, "Name");

            AddressDetailPage addressDetailPage = new AddressDetailPage(driver);
            
            //Click on the Tenancy Requests Associated View in the dropdown ribbon menu
            homePage.HoverAddressDetailRibbonTab(rentalPremises);
            homePage.ClickTRAddressDetailViewRibbonButton();

            addressDetailPage = new AddressDetailPage(driver);

            string BaseWindow = driver.CurrentWindowHandle;
            addressDetailPage.ClickAddNewTenancyRequestButton();
            addressDetailPage.SwitchNewBrowser(driver, BaseWindow);
           
            //Fill in all mandatory details and Click Save&close
            TenancyRequestPage tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            //Set {Type} field to Bond Lodgement.
            tenancyRequestPage.SetRequestTypeListValue("Bond Lodgement");
                        
            tenancyRequestPage.PopulateTenancyRequestWithoutRentalPremisesFromAddressDetailAssociatedView(
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("REQUEST_TYPE")].Value.ToString(),
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

            tenancyRequestPage.SetDwellingTypeListValue(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("DWELLING_TYPE")].Value.ToString());
            tenancyRequestPage.ClickSaveButton();

            string tenancyRequest = tenancyRequestPage.GetRequestNumber();
            StringAssert.Contains(tenancyRequest, "TR-BL-", "Validating the Tenancy Request saved with the correct TR no fromat");

            tenancyRequestPage.ClickSaveCloseButton();
            driver = driver.SwitchTo().Window(BaseWindow);

            //Open the Tenancy Request which was just created
            addressDetailPage = new AddressDetailPage(driver);
            addressDetailPage.SetTenancyRequestSearchText(tenancyRequest);

            Table tenancyRequestSearchTable = new Table(addressDetailPage.GetTenancyRequestSearchResultTable());
            tenancyRequestSearchTable.ClickCellContainsValue("Name", tenancyRequest, "Name");
  
            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            Table requestPartyTable = new Table(tenancyRequestPage.GetRequestPartyTable());

            //Confirm the tenancy request has only one contributor i.e initial Request Party
            Assert.AreEqual(requestPartyTable.GetCellContainsValue("Request Amount",initialContribution,"Client"), initialRequestParty);
            Assert.AreEqual(1, requestPartyTable.GetRowCount()-1,"Validating only one contributor is avaialble");
         
            UICommon.DoubleClickElement(requestPartyTable.GetCellElementContainsValue("Client", initialRequestParty, "Client"), driver);
            Thread.Sleep(1000);
            TenancyRequestPartyPage tenancyRequestPartyPage = new TenancyRequestPartyPage(driver);
            tenancyRequestPartyPage.ClickPageTitle();

            //Try to deactivate the contributor. Fails with an error dialog
            tenancyRequestPartyPage.ClickDeactivateButton();
            WarningDialogueFramePage warningPage = new WarningDialogueFramePage(driver);
            warningPage.ClickProcessBeginButton();

            tenancyRequestPartyPage.ClickSaveCloseButton();

            //Go back to tenancy request form. The Delete button should not be available for RBS user
            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            requestPartyTable = new Table(tenancyRequestPage.GetRequestPartyTable());

            Assert.AreEqual(requestPartyTable.GetCellContainsValue("Request Amount", initialContribution, "Client"), initialRequestParty);
            Assert.AreEqual(1, requestPartyTable.GetRowCount() - 1, "Validating the initial contributor is still avaialble after trying to deactivate");
                    
            IWebElement element = requestPartyTable.GetCellElementContainsValue("Client", initialRequestParty, "Request Amount");
            Assert.IsFalse(tenancyRequestPage.ClickDeleteButtonIfDisplayed(element),"Validating that delete button not displyed to click");

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }
        
        [TestMethod]
        [TestProperty("TestCaseID", "6362")]
        public void ATC6362a_E2ESingleBPAYCancelTenancyRequest()
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

            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            //Navigate to: Rental Bond Services > Tenancy Requests
            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();

            string managingParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();
            string initialRequestParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString();

            /*Data prep:
            1 x Address Detail record including the word "Garage";
            1 x Tenancy Request (Bond Lodgment) record at "New" that will fail address validation for invalid keyword (i.e. using prepared address), not already associated to a Tenancy, not associated to a batch, Amount Paid with Lodgement matches Sum of Contributions, Managing Party client that does not already exist within AX;
            n x Tenancy Request Party records associated to the Tenancy Request using Client records that do not already exist within AX*/
           
            TenancyRequestsSearchPage tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestSearchPage.ClickNewTenancyRequestButton();

            TenancyRequestPage tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            tenancyRequestPage.PopulateTenancyRequestWithNoInitialAndManagingParty(
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("REQUEST_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGEMENT_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("NO_ROOMS")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("WEEKLY_RENT")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_START")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("ANTICIPATED_END")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString());

            tenancyRequestPage.CreateNewClient(managingParty);
            tenancyRequestPage = new TenancyRequestPage(driver);

            tenancyRequestPage.SetInitialRequestPartyWithSearch(initialRequestParty);
            tenancyRequestPage.SetDwellingTypeListValue("House");

            tenancyRequestPage.ClickSaveButton();

            string tenancyRequest = tenancyRequestPage.GetRequestNumber();
            StringAssert.Contains(tenancyRequest, "TR-BL-", "Validating the Tenancy Request saved with the correct TR no fromat");

            tenancyRequestPage.ClickSaveCloseButton();

            //Reopen the record and change the status reason to Ready for Validation and save.The record should show an exception with Invalid Keyword
            tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestSearchPage.SetTenancyRequestSearchText(tenancyRequest);
            Table table = new Table(tenancyRequestSearchPage.GetSearchResultTable());
            StringAssert.Contains(table.GetCellValue("Name", tenancyRequest, "Status Reason"), "New");

            table.ClickCellValue("Name", tenancyRequest, "Name");

            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            tenancyRequestPage.SetStatusReason("Ready for validation");
            tenancyRequestPage.ClickSaveCloseButton();
                  
            //Navigate to Queue - BL failed validation
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaQueuesButton();

            QueueSearchPage queuePage = new QueueSearchPage(driver);
            queuePage.SetPageFilterList("All Items");
            queuePage.SetQueue("Bond lodgement failed validation");
            
            Table queueTable = new Table(queuePage.GetSearchResultTable());
            
            //Open the Tenancy Request from the queue.Tenancy Request record displayed, status reason "Validation failed"at least one Request Queue Reason record created for invalid address keyword with status reason "To be resolved"

            queueTable.ClickCellContainsValue("Title", tenancyRequest, "Title");
            tenancyRequestPage = new TenancyRequestPage(driver);

            StringAssert.Contains(tenancyRequestPage.GetValidationStatusReason(), "Validation failed", "Validation should fail when invalid keyword entered in address details");

            //tenancyRequestPage.ClickQueueReasons();
            Table queueReasonTable = new Table(tenancyRequestPage.GetQueueReasonTable());
            StringAssert.Contains(queueReasonTable.GetCellContainsValue("Reason", "Invalid keyword detected", "Status Reason"), "To be resolved");

            //De-activate the Tenancy Request and refresh the record.Tenancy Request record updated with Status "Inactive", Status Reason "Cancelled", Tenancy not populated
            tenancyRequestPage.ClickDeactivateButton();

            WarningDialogueFramePage warningPage = new WarningDialogueFramePage(driver);
            warningPage.ClickProcessBeginButton();
            Thread.Sleep(5000);

            tenancyRequestPage = new TenancyRequestPage(driver);
            Assert.AreEqual(tenancyRequestPage.GetStatusReason(), "Cancelled", "Validating that TR:" + tenancyRequest + " Status Reason becomes Cancelled after deactivating the record");
            
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();

            //Validate whether TR found in the list of recent Inactive Tenancy Requests
            tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestSearchPage.SetPageFilterList("Inactive Tenancy Requests");
          
            //Table tenancyRequestSearchHeaderTable = new Table(tenancyRequestSearchPage.GetHeaderSearchResultTable());
            //tenancyRequestSearchHeaderTable.ClickTableColumnHeader("Created On");
           // tenancyRequestSearchHeaderTable.ClickTableColumnHeader("Created On");

            Table tenancyRequestSearchTable = new Table(tenancyRequestSearchPage.GetSearchResultTable());
            StringAssert.Contains(tenancyRequestSearchTable.GetCellContainsValue("Name", tenancyRequest, "Name"), tenancyRequest,"Validating the inactive tenancy request table has the deactivated tenancy request");

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

        [TestMethod]
        [TestProperty("TestCaseID", "4520")]
        public void ATC4520_RaiseExceptionForExcessBond()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "4520")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            //Login as RBS user
            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            //Navigate to: Rental Bond Services > Tenancy Requests
            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();

            TenancyRequestsSearchPage tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);

            //(Template) - Create new Tenancy Request (Bond Lodgement)
            tenancyRequestSearchPage.ClickNewTenancyRequestButton();

            TenancyRequestPage tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            //Set {Type} field to Bond Lodgement.
            tenancyRequestPage.SetRequestTypeListValue("Bond Lodgement");

            /*Enter details:Details entered 
             - Residential Tenancy
             - No rent subsidy
             - rent $700
             - amount paId $2801*/

            tenancyRequestPage.PopulateTenancyRequestFormBondLodgement(
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("REQUEST_TYPE")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_TYPE")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGEMENT_TYPE")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("WEEKLY_RENT")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_START")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("ANTICIPATED_END")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString(),
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MISC1")].Value.ToString()); //Rental subsidy

            //Save Tenancy Request
            tenancyRequestPage.ClickSaveButton();

            //Update Status Reason to 'Ready for validation' and Save
            tenancyRequestPage.SetStatusReason("Ready for validation");
            tenancyRequestPage.ClickSaveButton();

            //Confirm Status Reason set to 'Validation failed'
            StringAssert.Contains(tenancyRequestPage.GetStatusReason(), "Validation failed");

            //Confirm Queue reason for excess bond added.Status set to 'To be resolved'
            tenancyRequestPage.ClickQueueReasons();
            Table queueReasonTable = new Table(tenancyRequestPage.GetQueueReasonTable());
            StringAssert.Contains(queueReasonTable.GetCellContainsValue("Reason", "Invalid bond amount : Exceeds maximum bond", "Status Reason"), "To be resolved");

            //Confirm Maximum Allowed Bond set correctly
            //Assert: Maximum Allowed Bond set to $2800
            StringAssert.Contains(tenancyRequestPage.GetMaximumAllowedAmount(), "2,800.00");

            //TODO:Step is not clear.To be implemented

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

        [TestMethod]
        [TestProperty("TestCaseID", "6738")]
        public void ATC6738_RaiseExceptionInvalidKeyword()
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

            string initialRequestParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString();
            string initialContribution = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString();
            string rentalPremises = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString();

            //Login as RBS user
            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            //Navigate to: Rental Bond Services > Tenancy Requests
            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();

            TenancyRequestsSearchPage tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);

            //(Template) - Create new Tenancy Request (Bond Lodgement)
            tenancyRequestSearchPage.ClickNewTenancyRequestButton();

            TenancyRequestPage tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            //Set {Type} field to Bond Lodgement.
            tenancyRequestPage.SetRequestTypeListValue("Bond Lodgement");

            /*Select 'Australian Physical' as the Type.Enter an address into the fields on the Address Detail record. As part of the address that you enter, Type in the word 'Garage' into the 'Complex Unit Number' field 
            Click on the Save & Close button*/
            string[] address = rentalPremises.Split(',');
            string complexunitno = address[0];
            string roadno = address[1].Split(' ')[1];
            string roadname = address[1].Split(' ')[2];
            string locality = address[2].Split(' ')[1] + "," + address[3] + "," + address[4];
            
            tenancyRequestPage.CreateNewAddress(roadno,roadname,locality,"","",complexunitno);
            tenancyRequestPage = new TenancyRequestPage(driver);

            //Select the mandatory fields for the Tenancy Request.Add a new Request Party whose contribution amount is the same as the Amount Paid with Lodgement 
            tenancyRequestPage.PopulateTenancyRequestWithoutRentalPremisesFromAddressDetailAssociatedView(
              MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("REQUEST_TYPE")].Value.ToString(),
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
        
            tenancyRequestPage.SetDwellingTypeListValue("House");
            //Save the Tenancy Request 
            tenancyRequestPage.ClickSaveButton();

            string tenancyrequest = tenancyRequestPage.GetRequestNumber();

            //Set the status of the Tenancy Request to be 'Ready for Validation'	Validation process for the Tenancy Request is started
            tenancyRequestPage.SetStatusReason("Ready for validation");
            tenancyRequestPage.ClickSaveButton();
            Thread.Sleep(3000);

            //The status of the Tenancy Request should be set to Validation failed 
            StringAssert.Contains(tenancyRequestPage.GetStatusReason(), "Validation failed");

            //Check the Queue Reasons of the Tenancy Request.There should be a queue reason that exists which is for 'Invalid address: Invalid key word detected' 
            //tenancyRequestPage.ClickQueueReasons();
            Table queueReasonTable = new Table(tenancyRequestPage.GetQueueReasonTable());
            StringAssert.Contains(queueReasonTable.GetCellContainsValue("Reason", "Invalid address : Invalid keyword detected", "Status Reason"), "To be resolved");

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

        [TestMethod]
        public void ATC6790_TenancyRequestRoomingAccomodation()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "6790")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string initialRequestParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString();
            string initialContribution = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString();

            //Login as RBS user
            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            //Navigate to: Rental Bond Services > Tenancy Requests
            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();

            TenancyRequestsSearchPage tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestSearchPage.SetPageFilterList("All Tenancy Requests");
            Table table = new Table(tenancyRequestSearchPage.GetHeaderSearchResultTable());
            table.ClickTableColumnHeader("Created On");
            table.ClickTableColumnHeader("Created On");

            Table searchTable = new Table(tenancyRequestSearchPage.GetSearchResultTable());
            string lastCreatedTenancyReqNo = searchTable.GetCellContainsValue("Name", "TR-BL", "Name");

            //(Template) - Create new Tenancy Request (Bond Lodgement)
            tenancyRequestSearchPage.ClickNewTenancyRequestButton();

            TenancyRequestPage tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            //Set {Type} field to Bond Lodgement.
            tenancyRequestPage.SetRequestTypeListValue("Bond Lodgement");
            tenancyRequestPage.SetTenancyTypeListValue("Rooming Accommodation");

            //Check Dwelling Type field	1. Lookup list filtered to only: [ Boarding House | Supported Accommodation | Student Accommodation off Campus |Studio]
            Assert.IsTrue(tenancyRequestPage.GetDwellingTypeText("Boarding House"));
            tenancyRequestPage.ClickPageTitle();
            Assert.IsTrue(tenancyRequestPage.GetDwellingTypeText("Supported Accommodation"));
            tenancyRequestPage.ClickPageTitle();
            Assert.IsTrue(tenancyRequestPage.GetDwellingTypeText("Student Accommodation off Campus"));
            tenancyRequestPage.ClickPageTitle();
            Assert.IsTrue(tenancyRequestPage.GetDwellingTypeText("Studio"));

            //Set Dwelling Type to Boarding House.Number of Bedrooms defaults to 1 and shows as read only
            tenancyRequestPage.SetDwellingTypeListValue("Boarding House");
            tenancyRequestPage.ClickPageTitle();
            StringAssert.Contains(tenancyRequestPage.GetNumberOfBedrooms(), "1");
            StringAssert.Contains(tenancyRequestPage.GetNumberOfBedroomsProperty("data-controlmode"), "locked");

            //Set Dwelling Type to Supported Accommodation.Number of Bedrooms defaults to 1 and shows as read only
            tenancyRequestPage.SetDwellingTypeListValue("Supported Accommodation");
            tenancyRequestPage.ClickPageTitle();
            StringAssert.Contains(tenancyRequestPage.GetNumberOfBedrooms(), "1");
            StringAssert.Contains(tenancyRequestPage.GetNumberOfBedroomsProperty("data-controlmode"), "locked");

            //Set Dwelling Type to Studio	Number of Bedrooms defaults to 1 and shows as read only
            tenancyRequestPage.SetDwellingTypeListValue("Studio");
            tenancyRequestPage.ClickPageTitle();
            StringAssert.Contains(tenancyRequestPage.GetNumberOfBedrooms(), "1");
            StringAssert.Contains(tenancyRequestPage.GetNumberOfBedroomsProperty("data-controlmode"), "locked");

            //Set Dwelling Type to Student Accommodation off Campus.Number of Bedrooms defaults to 1 and shows as read only
            tenancyRequestPage.SetDwellingTypeListValue("Student Accommodation off Campus");
            tenancyRequestPage.ClickPageTitle();
            StringAssert.Contains(tenancyRequestPage.GetNumberOfBedrooms(), "1");
            StringAssert.Contains(tenancyRequestPage.GetNumberOfBedroomsProperty("data-controlmode"), "locked");

            //Check Management Type field. Lookup list filtered to only: [ Owner (or can be Lessor/Owner) | Real Estate Agent | Manager/provider | Other ]
            Assert.IsTrue(tenancyRequestPage.GetTenancyManagementTypeListValue("Lessor/Owner"));
            Assert.IsTrue(tenancyRequestPage.GetTenancyManagementTypeListValue("Real Estate Agent"));
            Assert.IsTrue(tenancyRequestPage.GetTenancyManagementTypeListValue("Manager/provider"));
            Assert.IsTrue(tenancyRequestPage.GetTenancyManagementTypeListValue("Other"));

            //Set Rental Premise to an address with the following details: [Type, "Australian Physical" | Locality populated | Room/Site Number NOT populated] 
            //Example:"33 SHORE ST, EAST BRISBANE, QLD, 4169"
            string roadnumber = "33";
            string roadname = "SHORE ST";
            string locality = "EAST BRISBANE, QLD, 4169";

            tenancyRequestPage.CreateNewAddress(roadnumber,roadname,locality);
            tenancyRequestPage = new TenancyRequestPage(driver);
            
            //Populate mandatory fields. Save record.
            tenancyRequestPage.SetManagingPartyListValue(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString());
            tenancyRequestPage.SetTenancyManagementTypeListValue(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGEMENT_TYPE")].Value.ToString());
            tenancyRequestPage.SetInitialRequestPartyWithSearch(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString());
            tenancyRequestPage.SetInitialConrtibution(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString());
            tenancyRequestPage.SetAmountPaidWithLodgement(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString());
            tenancyRequestPage.SetLodgementTypeListValue(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString());

            tenancyRequestPage.ClickSaveButton();

            //Record saves. Soft warning (Missing room number): DISPLAYED
            StringAssert.Contains(tenancyRequestPage.GetWarningMessage(), "The selected address does not have a Room/Site number.");

            /*The Name field for the Tenancy Request will be populated with 'TR-BL-' and concatenated with the Reference Number field that will be populated 
            with increments of 1 from the most recent Tenancy request record. Example: Given the most recent Tenancy Request name is 'TR-BL-20000016' then this record will be 'TR-BL-20000017'*/
            string tenancyRequest = tenancyRequestPage.GetRequestNumber();
            StringAssert.Contains(tenancyRequest, "TR-BL", "Validating the newly created tenancy request has TR-BL format");
            int oldTRNo = Int32.Parse(lastCreatedTenancyReqNo.Split('-')[2]);
            int newTRNo = Int32.Parse(tenancyRequest.Split('-')[2]);
            Assert.AreEqual(oldTRNo + 1, newTRNo, "Validating the TR No is incremented by 1 from the last created TR");
         
            tenancyRequestPage.ClickSaveCloseButton();

            tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestSearchPage.SetTenancyRequestSearchText(tenancyRequest);

            searchTable = new Table(tenancyRequestSearchPage.GetSearchResultTable());
            //Check record in record grid. As updated/entered
            StringAssert.Contains(searchTable.GetCellContainsValue("Name",tenancyRequest,"Tenancy Type"), "Rooming Accommodation");
            StringAssert.Contains(searchTable.GetCellContainsValue("Name",tenancyRequest,"Managing Party"),MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString());
            StringAssert.Contains(searchTable.GetCellContainsValue("Name",tenancyRequest,"Rental Premises"),MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("RENTAL_PREMISES")].Value.ToString());
            StringAssert.Contains(searchTable.GetCellContainsValue("Name", tenancyRequest, "Dwelling Type"), "Student Accommodation off Campus");
            StringAssert.Contains(searchTable.GetCellContainsValue("Name", tenancyRequest, "Amount Bond Paid with Lodgement").Replace(",",""), MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString());
            
            searchTable.ClickCellContainsValue("Name", tenancyRequest, "Name");

            //Remove Rental Premise address. Save record.	Record not saved as field is mandatory.
            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            tenancyRequestPage.ClearRentalPremisesValue();

            tenancyRequestPage.ClickSaveButton();
            StringAssert.Contains(tenancyRequestPage.GetRentalPremiseAddressErrorText(), "You must provide a value for Rental Premises.");
            tenancyRequestPage.ClickPageTitle();

            //Set Rental Premise to an address with the following details: [Type, "Australian Physical" | Locality populated | Room/Site Number populated] 
            //Example: "RM 4, 33 SHORE ST, EAST BRISBANE, QLD, 4101"	Record saves. Soft warning (Missing room number): NOT DISPLAYED
            string roomtype = "Room";
            string roomno = "4";
            tenancyRequestPage.CreateNewAddress(roadnumber, roadname, locality, roomtype, roomno);
            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickSaveButton();

            Assert.IsFalse(tenancyRequestPage.GetWarningMessage().Contains("The selected address does not have a Room/Site number."));
            tenancyRequestPage.ClickSaveCloseButton();

            //Check record in record grid	As updated/entered
            tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestSearchPage.SetTenancyRequestSearchText(tenancyRequest);
            searchTable = new Table(tenancyRequestSearchPage.GetSearchResultTable());
            StringAssert.Contains(searchTable.GetCellContainsValue("Name", tenancyRequest, "Rental Premises"), "RM 4, 33 SHORE ST, EAST BRISBANE, QLD, 4169");

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

        [TestMethod]
        [TestProperty("TestcaseID", "6751")]
        public void ATC6751a_CRMBTriggerUpdateToCheque()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "6751")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string managingParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();
            string initialRequestParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString();
            string amountOtherParty = (MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value +
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MISC2")].Value).ToString();

            //Login as RBS user
            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            //Navigate to: Rental Bond Services > Tenancy Requests
            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();

            TenancyRequestsSearchPage tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestSearchPage.ClickNewTenancyRequestButton();

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
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("LODGEMENT_TYPE")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY_START")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("ANTICIPATED_END")].Value.ToString(),
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAYMENT_TYPE")].Value.ToString());

            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.SetDwellingTypeListValue("House");

            tenancyRequestPage.ClickSaveButton();

            string tenancyRequest = tenancyRequestPage.GetRequestNumber();
            tenancyRequestPage.ClickSaveButton();
            string tenancyrequest = tenancyRequestPage.GetRequestNumber();
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value = tenancyrequest;
            tenancyRequestPage.SetStatusReason("Ready for validation");
            tenancyRequestPage.ClickSaveButton();
            StringAssert.Contains(tenancyRequestPage.GetValidationStatusReason(), "Validation successful");
            tenancyRequestPage.ClickSaveCloseButton();

            //Reopen the record and change the status reason to Ready for Validation and save.The record should show an exception with Invalid Keyword
            tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);
            tenancyRequestSearchPage.SetTenancyRequestSearchText(tenancyRequest);
            Table table = new Table(tenancyRequestSearchPage.GetSearchResultTable());
            table.ClickCellValue("Name", tenancyRequest, "Name");

            tenancyRequestPage = new TenancyRequestPage(driver);
            tenancyRequestPage.ClickPageTitle();

            table = new Table(tenancyRequestPage.GetPaymentSummaryResultTable());

            string referencenumber = table.GetCellValue("Tenancy Request", tenancyRequest, "Reference Number");
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
   }     
}
