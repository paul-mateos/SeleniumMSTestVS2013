using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.TestTools.UITesting;
using RTA.Automation.CRM.UI;
using RTA.Automation.CRM.Environment;
using OpenQA.Selenium;
using RTA.Automation.CRM.Pages;
using Excel = Microsoft.Office.Interop.Excel;
using RTA.Automation.CRM.Pages.Investigations;
using System.IO;
using System.Reflection;
using System.Threading;
using RTA.Automation.CRM.DataSource;
using System.Windows.Forms;
//using Kobets.Automation.Infrastructure.OfficeTools.OpenXML;
using System.Data;
using System.Data.Linq;
using System.Linq;

namespace RTA.Automation.CRM.Tests
{
    [TestClass]
    public class CRMNewActivityTests : BaseTest
    {
      
        
        [TestInitialize]
        public void DataSourceSetup()
        {
            //open excel
            MyApp = new Excel.Application();
            MyApp.Visible = false;

        }

        

        [TestMethod]
        [TestProperty("TestcaseID", "3367")]
        public void ATC3367b_CRMSendemailfromCRMNewActivity()
        {
           
            #region Start Up Excel
            MyBook = MyApp.Workbooks.Open(DatasourceDir + @"\Investigations.xlsx", 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            MySheet = (Excel.Worksheet)MyBook.Sheets[Properties.Settings.Default.ENVIRONMENT.ToString()];
            MyRange = MySheet.UsedRange;

            //Get specific row for the data
            int testDataRows = MyRange.Rows.Count;
            int MyRow = 0;
            for (int i = 2; i <= testDataRows; i++)
            {
                if (MyRange.Cells[i, 1].Value.ToString()== "3367")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion
                      
            User user = this.environment.GetUser(SecurityRole.SystemAdministrator);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.ClickNewActivityIMG();
            homePage.ClickEmailIMG();

            //Set values for email
            EmailPage emailPage = new EmailPage(driver);
            emailPage.SetRegardingValueText(MyRange.Cells[MyRow, InvestigationSchema.GetColumnIndex(ColumnName.IN_ID)].Value.ToString());
            emailPage.SetToValueText("BLAIR TEST");
            emailPage.ClickPageTitle();
            emailPage.SetSubjectValueText("Test 3367 CRM Email Creation");
            emailPage.SetEmailBODYText("Main body of email test");
            emailPage.ClickSaveIMG();
            //Need to add attachment
 
            emailPage.ClickSendEmailIMG();
            //Check that the page has closed

            //emailPage.AcceptRTAValidationMessage("The sender's mailbox is disabled.");

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }



        [TestMethod]
        [TestProperty("TestcaseID", "3307")]
        public void ATC3307_CRMRecordphonecallactivity()
        {
            User user = this.environment.GetUser(SecurityRole.SystemAdministrator);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.ClickCreateIMG();
            homePage.ClickCreatePhoneActivityRibbonButton();

            PhoneCallPage phoneCallPage = new PhoneCallPage(driver);
            phoneCallPage.ClickPageTitle();
            phoneCallPage.SetSubject(DateTime.Now.ToString());
            phoneCallPage.SetRecipient("IMSTestU12");

            phoneCallPage.ClickSaveButton();
            string activityName = phoneCallPage.GetPageTitle();

            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientActivitiesRibbonButton();

            ActivitiesSearchPage activitiesSearchPage = new ActivitiesSearchPage(driver);
            activitiesSearchPage.SetTenancyRequestSearchText(activityName);
            Table table = new Table(activitiesSearchPage.GetSearchResultTable());
            StringAssert.Contains(table.GetCellValue("Subject", activityName, "Activity Type"), "Phone Call");


        }

        [TestMethod]
        [TestProperty("TestcaseID", "3344")]
        public void ATC3344_CRMFrontCounterNewActivity()
        {
            #region Start Up Excel
            MyBook = MyApp.Workbooks.Open(DatasourceDir + @"\Investigations.xlsx", 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            MySheet = (Excel.Worksheet)MyBook.Sheets[Properties.Settings.Default.ENVIRONMENT.ToString()];
            MyRange = MySheet.UsedRange;

            //Get specific row for the data
            int testDataRows = MyRange.Rows.Count;
            int MyRow = 0;
            for (int i = 2; i <= testDataRows; i++)
            {
                if (MyRange.Cells[i, 1].Value.ToString() == "ClientTestData")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string ClientName = MyRange.Cells[MyRow, InvestigationSchema.GetColumnIndex(ColumnName.CLIENT_NAME)].Value.ToString();
            string FrontCounterActivityName = "TC 3344 Front Counter Activity";
            User user = this.environment.GetUser(SecurityRole.Default);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            String HomeWindow = driver.CurrentWindowHandle;
            homePage.ClickCreateIMG();
            homePage.ClickFrontCounterContactActivityRibbonButton();

            FrontCounterContactPage frontCounterContactPage = new FrontCounterContactPage(driver);
            string FrontCounterActivityWindow = driver.CurrentWindowHandle;
            frontCounterContactPage.ClickPageTitle();

            // Assign a Client and fill in all possisble fields.
            frontCounterContactPage.SetSelectSubjectValue("Bond existence");
            frontCounterContactPage.SetSubjectValue(FrontCounterActivityName);
            frontCounterContactPage.SetClientName(ClientName);
            frontCounterContactPage.SetActualEndDate(DateTime.Today.AddDays(2).ToString("dd/MM/yyyy"));
            frontCounterContactPage.SetActualDuration("30 minutes");
            frontCounterContactPage.SetAssistiveService("English");
            frontCounterContactPage.SetRegardingClientValue(ClientName);
            frontCounterContactPage.ClickSaveButton();

            // Add the Activity to a queue through the entity menu overflow.
            frontCounterContactPage.ClickAddToQueueButton();
            frontCounterContactPage.SetQueue("Investigations Support Officers");
            frontCounterContactPage.ClickDialogAddButton();

            // add a connection to the Activity.
            driver = driver.SwitchTo().Window(HomeWindow);
            homePage.HoverFrontCounterContactXRibbonTab(FrontCounterActivityName);
            Thread.Sleep(500);
            homePage.ClickFrontCounterContactXConnectionsButton();

            driver = driver.SwitchTo().Window(FrontCounterActivityWindow);
            frontCounterContactPage.SetConnectList("To Me");

            Thread.Sleep(2000);

            driver = frontCounterContactPage.SwitchNewBrowser(driver, FrontCounterActivityWindow, "Connection");

            ConnectionPage connectionsPage = new ConnectionPage(driver);
            connectionsPage.ClickPageTitle();
            connectionsPage.ClickSaveCloseIMG();

            // Open the targeted entity of the connection. Confirm the connection is present.
            driver = driver.SwitchTo().Window(FrontCounterActivityWindow);
            frontCounterContactPage = new FrontCounterContactPage(driver);
            Table table = new Table(frontCounterContactPage.GetConnectionsTable());
            Thread.Sleep(1000);

            frontCounterContactPage.SwitchFrame();
            table.ClickCellValue("Connected To", "IMSTestU12 User", "Connected To");

            driver = driver.SwitchTo().Window(HomeWindow);
            homePage.HoverFrontCounterContactXRibbonTab("IMSTestU12 User");
            Thread.Sleep(500);
            homePage.ClickFrontCounterContactXConnectionsButton();

            UserPage userPage = new UserPage(driver);
            userPage.SetConnectionsSearchRecord(FrontCounterActivityName);
            table = new Table(userPage.GetConnectionsTable());
            Assert.IsTrue(table.MatchingCellFound("Connected To", FrontCounterActivityName), "New Connection NOT found !!!");

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }
    }
}
