using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.TestTools.UITesting;
using RTA.Automation.CRM.UI;
using RTA.Automation.CRM.Environment;
using OpenQA.Selenium;
using RTA.Automation.CRM.Pages;
using System.Collections.ObjectModel;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using RTA.Automation.CRM.DataSource;

namespace RTA.Automation.CRM.Tests
{
    [TestClass]
    public class CRMCreateNewTenancyTests : BaseTest
    {
        [TestInitialize]
        public void DataSourceSetup()
        {
            //open excel
            MyApp = new Excel.Application();
            MyApp.Visible = false;

        }
      
        [TestMethod]
        [TestProperty("TestcaseID", "4406")]
        public void ATC4406_CRMContributorRemovedStartEndDatefieldvalidation()
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

            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRibbonButton();

            //Search for test data
            TenancySearchPage tenancySearchPage = new TenancySearchPage(driver);
            tenancySearchPage.SetTenancySearchText(managingParty);
            Table table = new Table(tenancySearchPage.GetSearchResultTable());
            table.SelectTableRow("Contributors", managingParty);

            TenancyPage tenancyPage = new TenancyPage(driver);
            tenancyPage.ClickAllContributionsElement();

            table = new Table(tenancyPage.GetAllContributorsTable());
            table.SelectTableRow("Client", managingParty);

            //StringAssert.Contains(tenancyPage.GetRTAValidationMessage(), "Your changes have not been saved.");

            TenancyContributorPage tenancyContributorPage = new TenancyContributorPage(driver);
            //tenancyContributorPage.ClickStartDate();
            //tenancyContributorPage.SetStartDate("01/01/2016");
            //connectionPage.ClickStartDate();
            tenancyContributorPage.SetStartDateValue("01/01/2016");
            tenancyContributorPage.ClickSaveCloseButton();

            //Assert date was saved successfully
            tenancyPage = new TenancyPage(driver);
            tenancyPage.ClickAllContributionsElement();
            table = new Table(tenancyPage.GetAllContributorsTable());
            StringAssert.Contains(table.GetCellValue("Client", managingParty, "Start Date"), "1/01/2016");

            table.SelectTableRow("Client", managingParty);
           // StringAssert.Contains(tenancyPage.GetRTAValidationMessage(), "Your changes have not been saved.");

            tenancyContributorPage = new TenancyContributorPage(driver);
            tenancyContributorPage.SetStartDateValue("");
            tenancyContributorPage.SetEndDateValue("01/01/2016");
            tenancyContributorPage.ClickSaveCloseButton();

            //Assert date was saved successfully
            tenancyPage = new TenancyPage(driver);
            tenancyPage.ClickAllContributionsElement();
            table = new Table(tenancyPage.GetAllContributorsTable());
            StringAssert.Contains(table.GetCellValue("Client", managingParty, "End Date"), "1/01/2016");

            //Assert start date can not be after end date
            table.SelectTableRow("Client", managingParty);
           // StringAssert.Contains(tenancyPage.GetRTAValidationMessage(), "Your changes have not been saved.");

            tenancyContributorPage = new TenancyContributorPage(driver);
            tenancyContributorPage.ClickPageTitle();
            tenancyContributorPage.SetStartDateValue("02/01/2016");
            tenancyContributorPage.ClickSaveCloseButton();
            tenancyContributorPage = new TenancyContributorPage(driver);
            StringAssert.StartsWith(tenancyContributorPage.GetStartDateErrorText(), "Start date must be earlier than or the same as end date");

            //Remove and assert save successfull
            tenancyContributorPage.ClickPageTitle();
            tenancyContributorPage.SetStartDateValue("01/01/2016");
            tenancyContributorPage.ClickSaveCloseButton();
            tenancyPage = new TenancyPage(driver);
            tenancyPage.ClickAllContributionsElement();
            table = new Table(tenancyPage.GetAllContributorsTable());
            StringAssert.Contains(table.GetCellValue("Client", managingParty, "Start Date"), "1/01/2016");
            StringAssert.Contains(table.GetCellValue("Client", managingParty, "End Date"), "1/01/2016");

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }


        [TestMethod]
        [TestProperty("TestcaseID", "4409")]
        public void ATC4409_CRMEntermanagingpartyfromForm()
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
            User user = this.environment.GetUser(SecurityRole.SystemAdministrator);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRibbonButton();

            //Search for test data
            TenancySearchPage tenancySearchPage = new TenancySearchPage(driver);
            tenancySearchPage.SetTenancySearchText(managingParty);
            Table table = new Table(tenancySearchPage.GetSearchResultTable());
            table.SelectTableRow("Contributors", managingParty);

            TenancyPage tenancyPage = new TenancyPage(driver);
            tenancyPage.ClickConnectionsElement();
            tenancyPage.ClickConnectionsAssociationsIMG();

            string BaseWindow = driver.CurrentWindowHandle;
           
            tenancyPage.SetConnectList("To Another");

            driver = tenancyPage.SwitchNewBrowser(driver, BaseWindow, "Connection");
           
            //Issues with Connectoions page (name field) preventing test case to be completed
            ConnectionPage connectionsPage = new ConnectionPage(driver);
            connectionsPage.ClickPageTitle();
            connectionsPage.SetNameText(managingParty);
            connectionsPage.ClickPageTitle();
            connectionsPage.SetAsThisRoleText("Managing Party");
            connectionsPage.SetStartDate("01/01/2015");
            connectionsPage.SetEndDate("01/02/2015");
            connectionsPage.SetDesctiptionText("Test description text");
            
            connectionsPage.ClickSaveIMG();

            connectionsPage.ClickSaveCloseIMG();
            driver = driver.SwitchTo().Window(BaseWindow);

            tenancyPage.SetTenancyConnectionSearchText(managingParty);

            table = new Table(tenancyPage.GetSearchResultTable());
            StringAssert.Contains(table.GetCellValue("Connected To", managingParty, "Role (To)"), "Managing Party");

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

        [TestMethod]
        [TestProperty("TestcaseID", "6624")]
        public void ATC6624_ContributorsQuickSearch()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "6624")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion
           
            string contributor = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_REQUEST_PARTY")].Value.ToString();
            string bondnumber = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TENANCY")].Value.ToString();
            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRibbonButton();

            //Search for test data
            TenancySearchPage tenancySearchPage = new TenancySearchPage(driver);
            tenancySearchPage.SetTenancySearchText(contributor);
            Table table = new Table(tenancySearchPage.GetSearchResultTable());

            if(table.GetRowCount() > 0)
            {
                string actualContributor = table.GetCellValue("Bond Number", bondnumber,"Contributors");
                Assert.AreEqual(contributor, actualContributor, "Verify searching with contributor value in tenancies displays tenancy record with the contibutor");

            }
            else
            {
                throw new AssertFailedException("There are no records displayed when searched with contibutor:" + contributor);
            }
        }

        [TestMethod]
        [TestProperty("TestcaseID", "4428")]
        public void ATC4428_RoomnoCheckForRentalPremisesField()
        {
            //Login as a "CRM User" /* Notes to tester, use a login with role "CRM User", e.g. IMSTestU12 */
            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            /*Navigate to:
            RBS group > Tenancies tile*/
            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRibbonButton();

            //Click +NEW
            TenancySearchPage tenancySearchPage = new TenancySearchPage(driver);
            tenancySearchPage.ClickNewTenancyButton();

            TenancyPage tenancyPage = new TenancyPage(driver);
            tenancyPage.ClickPageTitle();

            //Set Tenancy Type to Rooming Accommodation
            tenancyPage.SetResidentialTenancyTypeList("Rooming Accommodation");
            
            /*Set Rental Premises to an address with the following details: Type "Australian Physical", Locality populated, Room/Site Number populated, i.e. "RM 4, 100 VULTURE ST, SOUTH BRISBANE, QLD, 4101"*/
            string roomtype = "Room";
            string roomno = "4";
            string roadnumber = "33";
            string roadname = "SHORE ST";
            string locality = "EAST BRISBANE, QLD, 4169";

            tenancyPage.CreateNewAddress(roadnumber, roadname, locality, roomtype, roomno);
            tenancyPage = new TenancyPage(driver);
            tenancyPage.ClickPageTitle();

            //Populate mandatory fields and save record. Record saves
            tenancyPage.SetManagingPartyListValue("ANDY TEST");
            tenancyPage.SetTenancyManagementTypeListValue("Lessor/Owner");
            tenancyPage.SetDwellingTypeListValue("Student Accommodation off Campus");

            tenancyPage.ClickSaveButton();

            //Open record, change Rental Premises to an address with the following details: Type "Australian Physical", Locality populated, Room/Site Number not populated, i.e. "33 SHORE ST, EAST BRISBANE, QLD, 4169"
            Assert.AreEqual("RM 4, 33 SHORE ST, EAST BRISBANE, QLD, 4169", tenancyPage.GetRentalPremises(),"Rental premises set as per entered");
            
            //Soft validation warning display due to missing room number.
            Assert.IsFalse(tenancyPage.VerifyWarningMessagePresent("rta_rental_premise_addressid"),"The selected address does not have a room number warning not present");
            
            tenancyPage.ClearRentalPremisesValue();
            tenancyPage.ClickSaveButton();
            tenancyPage.ClickPageTitle();

            tenancyPage.CreateNewAddress(roadnumber, roadname, locality);
            //Save record.Record saves
            tenancyPage.ClickSaveButton();
            Assert.IsTrue(tenancyPage.VerifyWarningMessagePresent("rta_rental_premise_addressid"), "The selected address does not have a room number warning is present");
            Assert.AreEqual("33 SHORE ST, EAST BRISBANE, QLD, 4169", tenancyPage.GetRentalPremises(),"Rental premises set as per entered");
        }

   }
}
