using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.TestTools.UITesting;
using RTA.Automation.CRM.UI;
using RTA.Automation.CRM.Environment;
using OpenQA.Selenium;
using RTA.Automation.CRM.Pages;
using RTA.Automation.CRM.Pages.Investigations;
using RTA.Automation.CRM.Pages.Clients;
using System.Threading;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Input;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;
using System.Collections.ObjectModel;
using Excel = Microsoft.Office.Interop.Excel;
using RTA.Automation.CRM.DataSource;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;


namespace RTA.Automation.CRM.Tests
{
    [TestClass]
    public class CRMInvestigationsTests : BaseTest
    {
        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);


        public CRMInvestigationsTests()
        {
        }
        [TestInitialize]
        public void DataSourceSetup()
        {
            //open excel
            MyApp = new Excel.Application();
            MyApp.Visible = false;

        }
       
        [TestMethod]
        [TestProperty("TestcaseID", "3284")]
        public void ATC3284_CRMEntityAllegedOffenceBeliefDateblankedStatusReasonreverts()
        {

            string allegedoffenceId;
            string todayDate = DateTime.Now.ToString("d/MM/yyyy");
            string investigationID; 
            
            //Login in as role
            User user = this.environment.GetUser(SecurityRole.Investigations);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePageInvestigation = new HomePage(driver);
            homePageInvestigation.HoverCRMRibbonTab();
            homePageInvestigation.ClickInvestigationsRibbonButton();
            homePageInvestigation.HoverInvestigationsRibbonTab();
            homePageInvestigation.ClickInvestigationsCasesRibbonButton();

            InvestigationCaseSearchPage investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationCaseSearchPage.ClickNewInvestigationCaseButton();

            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickSaveButton();
            investigationID = investigationCasePage.GetInvestigationCaseNumber();
            investigationCasePage.ClickSaveCloseButton();


            homePageInvestigation.HoverCRMRibbonTab();
            homePageInvestigation.ClickInvestigationsRibbonButton();
            homePageInvestigation.HoverInvestigationsRibbonTab();
            homePageInvestigation.ClickRightScrollRibbonButton();
            homePageInvestigation.ClickAllegedOffencesButton();

            AllegendOffensesSearchPage allegedOffencesSearchPage = new AllegendOffensesSearchPage(driver);

            allegedOffencesSearchPage.ClickNewAllegedOffenceButton();

            AllegedOffencePage allegedOffencesPage = new AllegedOffencePage(driver);
            allegedOffencesPage.SetInvestigationCaseValue(investigationID);
            allegedOffencesPage.SetProvisionValue("RTRA 116(1)");
            allegedOffencesPage.SetOffenceDateValue("1/01/2015");
            allegedOffencesPage.SetBeliefFormedDateValue(todayDate);
            allegedOffencesPage.ClickSaveButton();
            allegedoffenceId = allegedOffencesPage.GetReferenceNumber();
            allegedOffencesPage.ClickSaveCloseButton();

            allegedOffencesSearchPage = new AllegendOffensesSearchPage(driver);
            allegedOffencesSearchPage.SetInvestigationSearchText(investigationID);
            Table table = new Table(allegedOffencesSearchPage.GetSearchResultTable()); 
            table.SelectTableRow("Status Reason", "Belief");

            allegedOffencesPage = new AllegedOffencePage(driver);
            allegedOffencesPage.SetOffenceDateValue("");
            StringAssert.Contains(allegedOffencesPage.GetOffenceDateValue(),"--");
            allegedOffencesPage.SetBeliefFormedDateValue("");
            StringAssert.Contains(allegedOffencesPage.GetStatusReason(), "Suspicion");

            allegedoffenceId = allegedOffencesPage.GetReferenceNumber();
            allegedOffencesPage.ClickSaveCloseButton();

            allegedOffencesSearchPage = new AllegendOffensesSearchPage(driver);
            allegedOffencesSearchPage.SetInvestigationSearchText(allegedoffenceId);
            table = new Table(allegedOffencesSearchPage.GetSearchResultTable());
            table.SelectContainsTableRow("Investigation Case", investigationID);

            allegedOffencesPage = new AllegedOffencePage(driver);
            StringAssert.Contains(allegedOffencesPage.GetStatusReason(), "Suspicion");
            StringAssert.Contains(allegedOffencesPage.GetOffenceDateValue(), "--");
            StringAssert.Contains(allegedOffencesPage.GetStatutoryLimitationValue(), "--");

            allegedOffencesPage.SetOffenceDateValue("01/01/2015");
            allegedOffencesPage.ClickSaveButton();

            StringAssert.Contains(allegedOffencesPage.GetStatutoryLimitationValue(), "1/01/2016");
            
            StringAssert.Contains(allegedOffencesPage.GetBefliefFormedDateValue(),"");

        }

        [TestMethod]
        [TestProperty("TestcaseID", "3288")]
        public void ATC3288_CRMEntityConnectionCreateactiveclientconnections()
        {

            //Login in as role
            User user = this.environment.GetUser(SecurityRole.Investigations);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsClientRibbonButton();

            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.SetClientSearchText("BLAIR TEST");

            Table table = new Table(clientsSearchPage.GetSearchResultTable());
            table.ClickCellValue("Full Name", "BLAIR TEST","Full Name");

            homePage.HoverClientXRibbonTab("BLAIR TEST");
            homePage.ClickClientXConnectionsRibbonButton();


            ClientPage clientPage = new ClientPage(driver);
            string BaseWindow = driver.CurrentWindowHandle;
            clientPage.SetConnectList("To Another");

            driver = clientPage.SwitchNewBrowser(driver, BaseWindow, "Connection");

            
            ConnectionPage connectionsPage = new ConnectionPage(driver);
            connectionsPage.ClickPageTitle();
            connectionsPage.SetNameText("BLAIR TEST");
            connectionsPage.ClickPageTitle();
            connectionsPage.SetAsThisRoleText("Bankrupt");
            connectionsPage.ClickPageTitle();
            connectionsPage.SetAsThisRoleText("Child");
            connectionsPage.ClickPageTitle();
            connectionsPage.SetAsThisRoleText("Co Owner");
            connectionsPage.ClickPageTitle(); 
            connectionsPage.SetAsThisRoleText("Colleague");
            connectionsPage.ClickPageTitle(); 
            connectionsPage.SetAsThisRoleText("Co-Tenant");
            connectionsPage.ClickPageTitle(); 
            connectionsPage.SetAsThisRoleText("Employee");
            connectionsPage.ClickPageTitle(); 
            connectionsPage.SetAsThisRoleText("Employer");
            connectionsPage.ClickPageTitle(); 
            connectionsPage.SetAsThisRoleText("External Agency");
            connectionsPage.ClickPageTitle(); 
            connectionsPage.SetAsThisRoleText("Former Employer");
            connectionsPage.ClickPageTitle(); 
            connectionsPage.SetAsThisRoleText("Friend");
            connectionsPage.ClickPageTitle(); 
            connectionsPage.SetAsThisRoleText("Licensee/Business Owner");
            connectionsPage.ClickPageTitle(); 
            connectionsPage.SetAsThisRoleText("Neighbour");
            connectionsPage.ClickPageTitle(); 
            connectionsPage.SetAsThisRoleText("On Site Manager");
            connectionsPage.ClickPageTitle(); 
            connectionsPage.SetAsThisRoleText("Owner/Lessor");
            connectionsPage.ClickPageTitle(); 
            connectionsPage.SetAsThisRoleText("Parent");
            connectionsPage.ClickPageTitle(); 
            connectionsPage.SetAsThisRoleText("Partner");
            connectionsPage.ClickPageTitle(); 
            connectionsPage.SetAsThisRoleText("Potential Duplicate");
            connectionsPage.ClickPageTitle(); 
            connectionsPage.SetAsThisRoleText("Primary Case");
            connectionsPage.ClickPageTitle(); 
            connectionsPage.SetAsThisRoleText("Process Failure");
            connectionsPage.ClickPageTitle(); 
            connectionsPage.SetAsThisRoleText("Property Manager");
            connectionsPage.ClickPageTitle(); 
            connectionsPage.SetAsThisRoleText("Receiver");
            connectionsPage.ClickPageTitle(); 
            connectionsPage.SetAsThisRoleText("Referral");
            connectionsPage.ClickPageTitle(); 
            connectionsPage.SetAsThisRoleText("Referred by");
            connectionsPage.ClickPageTitle(); 
            connectionsPage.SetAsThisRoleText("Related case");
            connectionsPage.ClickPageTitle(); 
            connectionsPage.SetAsThisRoleText("Resident");
            connectionsPage.ClickPageTitle(); 
            connectionsPage.SetAsThisRoleText("Spouse");
            connectionsPage.ClickPageTitle(); 
            connectionsPage.SetAsThisRoleText("Stakeholder");
            connectionsPage.ClickPageTitle(); 
            connectionsPage.SetAsThisRoleText("Tenant");
            connectionsPage.ClickPageTitle(); 
            connectionsPage.SetAsThisRoleText("Witness");
            connectionsPage.ClickPageTitle(); 
            connectionsPage.SetAsThisRoleText("Managing Party");
            
            connectionsPage.SetStartDate("01/01/2015");
            connectionsPage.SetEndDate("01/02/2015");
            connectionsPage.SetDesctiptionText("Test description text");

            connectionsPage.ClickSaveIMG();

            connectionsPage.ClickSaveCloseIMG();
            driver = driver.SwitchTo().Window(BaseWindow);

            clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();
            clientPage.SetClientSearchText("BLAIR TEST");

            table = new Table(clientPage.GetSearchResultTable());
            StringAssert.Contains(table.GetCellValue("Connected To", "BLAIR TEST", "Role (To)"), "Managing Party");

        }

        [TestMethod]
        [TestProperty("TestcaseID", "3306")]
        public void ATC3306_CRMInvestigationCaseReopenclosedcase()
        {

            //Login in as role
            User user = this.environment.GetUser(SecurityRole.Investigations);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsCasesRibbonButton();

            InvestigationCaseSearchPage investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
            //investigationCaseSearchPage.ClickNewInvestigationCaseButton();
            investigationCaseSearchPage.SetPageFilterList("All Investigation Cases");
            Table table = new Table(investigationCaseSearchPage.GetSearchResultTable());
            table.SelectTableRow("Investigation Status", "Closed");

             InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickStartDialogButton();
           
            table = new Table(investigationCasePage.GetProcessSearchResultTable());
            table.ClickCell("Process Name", "INV: Re-open investigation case", "Created On");
            string BaseWindow = driver.CurrentWindowHandle; //Records the current window handle
            investigationCasePage.ClickDialogAddButton();

            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "INV:");
           
            INVPage iNVPage = new INVPage(driver);
            iNVPage.ClickNextButton();
            iNVPage.ClickNextButton();
            iNVPage.ClickFinishButton();

            driver = driver.SwitchTo().Window(BaseWindow);

        }


        [TestMethod]
        [TestProperty("TestcaseID", "3342")]
        public void ATC3342_CRMInvestigationTestActivitiesCreatedBySelectingSubStatus()
        {
            InvestigationCaseSearchPage investigationCaseSearchPage;
            InvestigationCasePage investigationCasePage;
            string investigationID;

            //Login in as role
            User user = this.environment.GetUser(SecurityRole.Investigations);
            new LoginDialog().Login(user.Id, user.Password);

            string[] Status = new string[] { "New case", "Investigation current", "Investigation current" , "Investigation current", "Investigation current", "Investigation current", "Investigation current", "Case finalisation", "Case finalisation", "Case finalisation"};
            string[] SubStatus = new string[] { "Initial assessment", "Conducting background searches", "Requesting information from complainant(s)" ,"Requesting information from third parties", "Search warrant(s) required", "Allegations to be put to respondent(s)", "Seeking legal advice", "Case closed discussion", "Educating respondent(s)", "Prosecution required"};
            string[] Activities0 = { "Carry out initial assessment", "Allocate"};
            string[] Activities1 = { "Conduct respondent(s) background searches", "Conduct complainant(s) background searches", "Conduct relevant dispute searches", "Conduct relevant premises background searches" };
            string[] Activities2 = { "Send request for information to complainant(s)", "Receive response from complainant" };  
            string[] Activities3 = { "Send request for information to third parties", "Receive response from third parties" };
            string[] Activities4 = {"Discuss with Manager", "Prepare warrant application(s)", "Receive approval of warrant(s)"};
            string[] Activities5 = { "Send letter of allegations to respondent", "Receive response from respondent" };  
            string[] Activities6 = { "Discuss with SIO or Manager", "Prepare request for legal advice", "Receive legal advice" };
            string[] Activities7 = { "Conduct closure discussion", "Record closure discussion outcome" };
            string[] Activities8 = { "Send educative letter to respondent(s)" };
            string[] Activities9 = { "Prepare brief of evidence", "Seek legal endorsement", "Receive legal advice" };

            HomePage homePage = new HomePage(driver);
            String HomeWindow = driver.CurrentWindowHandle; 

            for (int i = 0; i < Status.Length; i++)
            {
  
                homePage.HoverCRMRibbonTab();
                homePage.ClickInvestigationsRibbonButton();
                homePage.HoverInvestigationsRibbonTab();
                homePage.ClickInvestigationsCasesRibbonButton();

                investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
                String BaseWindow = driver.CurrentWindowHandle; 
                investigationCaseSearchPage.ClickNewInvestigationCaseButton();

                investigationCasePage = new InvestigationCasePage(driver);
                investigationCasePage.ClickSaveButton();
                investigationID = investigationCasePage.GetInvestigationCaseNumber();
                investigationCasePage.ClickSaveCloseButton();

                // Search for the Investigation Case ID
                investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
                investigationCaseSearchPage.SetSearchRecord(investigationID);

                Table table = new Table(investigationCaseSearchPage.GetSearchResultTable());
                table.ClickCellValue("Case Number", investigationID, "Case Number");

                investigationCasePage = new InvestigationCasePage(driver);
                             
                investigationCasePage.SetStatus(Status[i]);
                investigationCasePage.ClickPageTitle();
                investigationCasePage.SetSubStatus(SubStatus[i]);

                investigationCasePage.ClickSaveCloseButton();

                // Search for the Investigation Case ID
                investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
                investigationCaseSearchPage.SetSearchRecord(investigationID);

                table = new Table(investigationCaseSearchPage.GetSearchResultTable());
                table.ClickCellValue("Case Number", investigationID, "Case Number");

                // Verify the Activities created
                driver.Navigate().Refresh();
                Thread.Sleep(1000);

                investigationCasePage = new InvestigationCasePage(driver);
                investigationCasePage.SetPageFilterList("All Activities");
                Thread.Sleep(1000);
                table = new Table(investigationCasePage.GetActivitiesSearchResultTable());

                string[] ActivityArray = {""};
                switch (i)
                { 
                    case 0:
                        ActivityArray = Activities0;
                        break;
                    case 1:
                        ActivityArray = Activities1;
                        break;
                    case 2:
                        ActivityArray = Activities2;
                        break;
                    case 3:
                        ActivityArray = Activities3;
                        break;
                    case 4:
                        ActivityArray = Activities4;
                        break;
                    case 5:
                        ActivityArray = Activities5;
                        break;
                    case 6:
                        ActivityArray = Activities6;
                        break;
                    case 7:
                        ActivityArray = Activities7;
                        break;
                    case 8:
                        ActivityArray = Activities8;
                        break;
                    case 9:
                        ActivityArray = Activities9;
                        break;
                }
                for (int j = 0; j < ActivityArray.Length; j++)
                {
                    StringAssert.Contains(table.GetCellContainsValue("Subject", investigationID + ": " + ActivityArray[j], "Activity Status"), "Open");
                }
                driver.SwitchTo().Window(HomeWindow);
            }
        }
    

        [TestMethod]
        [TestProperty("TestcaseID", "3327")]
        public void ATC3327_CRMInvestigationCloseCaseWithMultipleOpenCloseActions()
        {
           //Login in as role
            User user = this.environment.GetUser(SecurityRole.Investigations);
            
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsCasesRibbonButton();

            // Create new investigation case
            InvestigationCaseSearchPage investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationCaseSearchPage.ClickNewInvestigationCaseButton();

            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickSaveButton();
            string investigationID = investigationCasePage.GetInvestigationCaseNumber();
            investigationCasePage.ClickSaveCloseButton();

            // Search for the Investigation Case ID
            investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationCaseSearchPage.SetSearchRecord(investigationID);

            Table table = new Table(investigationCaseSearchPage.GetSearchResultTable());
            table.ClickCellValue("Case Number", investigationID, "Case Number");


            driver.Navigate().Refresh();
            Thread.Sleep(1000);

            investigationCasePage = new InvestigationCasePage(driver);
            string BaseWindow = driver.CurrentWindowHandle; //Records the current window handle
            investigationCasePage.SetPageFilterList("All Activities");
            Thread.Sleep(1000);
            table = new Table(investigationCasePage.GetActivitiesSearchResultTable());

            // Add 1st completed activity
            investigationCasePage.ClickActivitiesAddButton();
            investigationCasePage.ClickAddTaskButton("Task");

            //Enter Request Party details
            driver = investigationCasePage.SwitchNewBrowser(driver, BaseWindow);
            TaskPage taskPage = new TaskPage(driver);
            Thread.Sleep(100);
            taskPage.ClickPageTitle();
            taskPage.SetSelectSubjectValue("Relevant address added");
            taskPage.ClickSaveButton();
            taskPage.ClickMarkCompleteButton();

            driver = driver.SwitchTo().Window(BaseWindow);
            
            // Add 2nd completed activity
            investigationCasePage.ClickPageTitle();
            investigationCasePage.ClickActivitiesAddButton();
            BaseWindow = driver.CurrentWindowHandle; //Records the current window handle
            investigationCasePage.ClickAddTaskButton("Task");

            //Enter Request Party details
            driver = investigationCasePage.SwitchNewBrowser(driver, BaseWindow);
            Thread.Sleep(100);
            taskPage = new TaskPage(driver);

            taskPage.ClickPageTitle();
            taskPage.SetSelectSubjectValue("Allocate to investigator");
            taskPage.ClickSaveButton();
            taskPage.ClickMarkCompleteButton();

            driver = driver.SwitchTo().Window(BaseWindow);

            investigationCasePage.ClickPageTitle();
            investigationCasePage.SetPageFilterList("All Activities");
            Thread.Sleep(1000);
            table = new Table(investigationCasePage.GetActivitiesSearchResultTable());

            // Verify that we are having 2 completed and 2 open activities
            StringAssert.Contains(table.GetCellContainsValue("Subject", "Relevant address added", "Activity Status"), "Completed");
            StringAssert.Contains(table.GetCellContainsValue("Subject", "Allocate to investigator", "Activity Status"), "Completed");
            StringAssert.Contains(table.GetCellContainsValue("Subject", investigationID + ": Scan documents", "Activity Status"), "Open");
            StringAssert.Contains(table.GetCellContainsValue("Subject", investigationID + ": Add parties", "Activity Status"), "Open");
            StringAssert.Contains(table.GetCellContainsValue("Subject", investigationID + ": Submit for initial assessment", "Activity Status"), "Open");

            Assert.AreEqual(5, table.GetRowCount()-1, "Additional activities seen");   

            // Now try to close the case
            investigationCasePage.ClickStartDialogButton();
            table = new Table(investigationCasePage.GetProcessSearchResultTable());
            table.ClickCell("Process Name", "INV: Close investigation case", "Created On");
            BaseWindow = driver.CurrentWindowHandle; //Records the current window handle
            investigationCasePage.ClickDialogAddButton();

            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "INV:");

            INVPage iNVPage = new INVPage(driver);
            iNVPage.ClickNextButton();
            iNVPage.ClickNextButton();
            iNVPage.ClickNextButton();
            iNVPage.ClickFinishButton();

            driver = driver.SwitchTo().Window(BaseWindow);

            // Now Investigation case is closed.. verify that this case is not seen in active Investigation case
            investigationCasePage.ClickSaveCloseButton();

            investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationCaseSearchPage.SetPageFilterList("Active Investigation Cases");

            investigationCaseSearchPage.SetSearchRecord(investigationID);

            StringAssert.Contains(driver.FindElement(By.ClassName("ms-crm-List-MessageText")).Text, "No Investigation Case records are available in this view.");
        }


        [TestMethod]
        [TestProperty("TestcaseID", "3330")]
        public void ATC3330_CRMInvestigationValidateSubstatusAfterSelectingStatus()
        {
            //Login in as role
            User user = this.environment.GetUser(SecurityRole.Investigations);

            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsCasesRibbonButton();

            // Create new investigation case
            InvestigationCaseSearchPage investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationCaseSearchPage.ClickNewInvestigationCaseButton();

            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            string BaseWindow = driver.CurrentWindowHandle; //Records the current window handles
            investigationCasePage.ClickPageTitle();
            investigationCasePage.SetStatus("Investigation current");
            investigationCasePage.ClickPageTitle();
            investigationCasePage.SetSubStatus("Conducting background searches");
            investigationCasePage.ClickPageTitle();
            investigationCasePage.ClickSaveButton();
            string investigationID = investigationCasePage.GetInvestigationCaseNumber();


            // Now close the investigation case
            investigationCasePage.ClickStartDialogButton();
            Table table = new Table(investigationCasePage.GetProcessSearchResultTable());
            table.ClickCell("Process Name", "INV: Close investigation case", "Created On");
            BaseWindow = driver.CurrentWindowHandle; //Records the current window handle
            investigationCasePage.ClickDialogAddButton();

            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "INV:");

            INVPage iNVPage = new INVPage(driver);
            iNVPage.ClickNextButton();
            iNVPage.ClickNextButton();
            iNVPage.ClickNextButton();
            iNVPage.ClickFinishButton();

            driver = driver.SwitchTo().Window(BaseWindow);

            // Now Investigation case is closed.. verify that this case is not seen in active Investigation case
            investigationCasePage.ClickSaveButton();

            // RE-OPEN the Investigation Case by running the dialog
            investigationCasePage.ClickStartDialogButton();
            table = new Table(investigationCasePage.GetProcessSearchResultTable());
            table.ClickCell("Process Name", "INV: Re-open investigation case", "Created On");
            BaseWindow = driver.CurrentWindowHandle; //Records the current window handle
            investigationCasePage.ClickDialogAddButton();

            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "INV:");

            iNVPage = new INVPage(driver);
            iNVPage.ClickNextButton();
            iNVPage.ClickNextButton();
            iNVPage.ClickFinishButton();

            driver = driver.SwitchTo().Window(BaseWindow);
            investigationCasePage.ClickSaveButton();
            // Select any Status and verify that list of "Sub-status" is correctly displayed

            investigationCasePage.SetStatus("Case finalisation");
            investigationCasePage.ClickPageTitle();
            // Verify that Sub-status is cleared off and error message is thrown
            investigationCasePage.ClickSubStatusText();
            StringAssert.Contains(investigationCasePage.GetSubStatusErrorText(), "You must provide a value for Sub Status.");

            investigationCasePage.ClickPageTitle();

            investigationCasePage.ClickSubStatusSearchButton();

            Assert.IsTrue(investigationCasePage.FindSubStatusFromDropdown("Case closed discussion"));
            Assert.IsTrue(investigationCasePage.FindSubStatusFromDropdown("Close without further action"));
            Assert.IsTrue(investigationCasePage.FindSubStatusFromDropdown("Educating respondent(s)"));
            Assert.IsTrue(investigationCasePage.FindSubStatusFromDropdown("Prosecution required"));
            Assert.AreEqual(4, investigationCasePage.GetSubStatusCount());

        }

        [TestMethod]
        [TestProperty("TestcaseID", "4445")]
        public void ATC4445_CRMInvestigationFutureDateNotAllowedForFinalisedDate()
        {
            //Login in as role
            User user = this.environment.GetUser(SecurityRole.Investigations);

            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            
            homePage.ClickInvestigationsCasesRibbonButton();

            // Create new investigation case
            InvestigationCaseSearchPage investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationCaseSearchPage.ClickNewInvestigationCaseButton();

            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickSaveButton();
            string investigationID = investigationCasePage.GetInvestigationCaseNumber();
            investigationCasePage.ClickStartDialogButton();

            Table table = new Table(investigationCasePage.GetProcessSearchResultTable());
            table.ClickCell("Process Name", "INV: Close investigation case", "Created On");
            string BaseWindow = driver.CurrentWindowHandle; //Records the current window handle
            investigationCasePage.ClickDialogAddButton();

            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "INV:");
            
            INVPage iNVPage = new INVPage(driver);
            iNVPage.ClickNextButton();
            iNVPage.ClickNextButton();

            string tomorrowDate = DateTime.Today.AddDays(1).ToString("dd-MM-yyyy");
            string todayDate = DateTime.Today.ToString("dd-MM-yyyy");

            iNVPage.SetDate(tomorrowDate);

            iNVPage.ClickNextButton();

            Assert.AreEqual("! WARNING !\r\n\r\nYou have entered a finalised date in the future which is not allowed. Either go back and update the finalised date or continue for the current date to be used by default.", iNVPage.GetErrorMessage());

            iNVPage.ClickPreviousButton();

            iNVPage.SetDate(todayDate);

            iNVPage.ClickNextButton();

            Assert.AreEqual("This is the end of the dialog. Click Finish to close it.", iNVPage.GetFinishMessage());

            iNVPage.ClickFinishButton();                        
            

        }


        [TestMethod]
        [TestProperty("TestcaseID", "3315")]
        public void ATC3315_CRMInvestigationValidateOwnerChanges()
        {
            //Login in as role
            User user = this.environment.GetUser(SecurityRole.InvestigationsOfficer);

            new LoginDialog().Login(user.Id, user.Password);

            string[] allowedUserName = { "IMSTestU03", "IMSTestU04", "IMSTestU05", "IMSTestU06", "IMSTestU07", "IMSTestU09", "IMSTestU10", "IMSTestU12" };
            string[] notAllowedUserName = { "IMSTestU08", "IMSTestU01", "IMSTestU11", "IMSTestU13" }; 

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsCasesRibbonButton();

            // Create new investigation case
            InvestigationCaseSearchPage investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationCaseSearchPage.ClickNewInvestigationCaseButton();

            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            string BaseWindow = driver.CurrentWindowHandle; //Records the current window handles
            investigationCasePage.ClickPageTitle();

            investigationCasePage.ClickSaveButton();

            // Verify the NOT ALLOWED users
            for (int userNo = 0; userNo < notAllowedUserName.Length; userNo++)
            {
                investigationCasePage.SetOwnerValue(notAllowedUserName[userNo]);
                investigationCasePage.ClickPageTitle();
                investigationCasePage.ClickSaveButton();

                // Get the Alert message
                StringAssert.Contains(investigationCasePage.GetOwnerValidationMessageString("ErrorTitle"), "Assignee has insufficient privileges");
                StringAssert.Contains(investigationCasePage.GetOwnerValidationMessageString("ErrorMessage"), "The selected user does not have sufficient privileges to be assigned records of this type. For more information, contact your system administrator.");

                investigationCasePage.ClickDialogAddButton();
            }

            for (int userNo = 0; userNo < allowedUserName.Length; userNo++)
            {
                investigationCasePage.SetOwnerValue(allowedUserName[userNo]);
                investigationCasePage.ClickPageTitle();
                investigationCasePage.ClickSaveButton();

                StringAssert.Contains(investigationCasePage.GetOwnerNameValue(), allowedUserName[userNo]);
            }
        }

        [TestMethod]
        [TestProperty("TestcaseID", "3328")]
        public void ATC3328_CRMInvestigationReopenCaseProvidingAReason()
        {
            //Login in as role
            //Login in as role
            User user = this.environment.GetUser(SecurityRole.Investigations);

            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();

            homePage.ClickInvestigationsCasesRibbonButton();

            // Create new investigation case
            InvestigationCaseSearchPage investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationCaseSearchPage.ClickNewInvestigationCaseButton();

            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickSaveButton();
            string investigationID = investigationCasePage.GetInvestigationCaseNumber();
            investigationCasePage.ClickStartDialogButton();

            Table table = new Table(investigationCasePage.GetProcessSearchResultTable());
            table.ClickCell("Process Name", "INV: Close investigation case", "Created On");
            string BaseWindow = driver.CurrentWindowHandle; //Records the current window handle
            investigationCasePage.ClickDialogAddButton();

            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "INV:");
            INVPage iNVPage = new INVPage(driver);
            iNVPage.ClickNextButton();
            iNVPage.ClickNextButton();
            string todayDate = DateTime.Today.ToString("dd-MM-yyyy");
            iNVPage.SetDate(todayDate);
            iNVPage.ClickNextButton();
            iNVPage.ClickFinishButton();

            driver = driver.SwitchTo().Window(BaseWindow);

            investigationCasePage.ClickSaveCloseButton();
            investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationCaseSearchPage.SetPageFilterList("Inactive Investigation Cases");

            table = new Table(investigationCaseSearchPage.GetHeaderSearchResultTable());
            table.ClickTableColumnHeader("Case Number");
            table = new Table(investigationCaseSearchPage.GetSearchResultTable());
            table.ClickCellValue("Case Number", investigationID, "Case Number");

            investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickStartDialogButton();

            table = new Table(investigationCasePage.GetProcessSearchResultTable());
            table.ClickCell("Process Name", "INV: Re-open investigation case", "Created On");
            string BaseWindow1 = driver.CurrentWindowHandle; 

            investigationCasePage.ClickDialogAddButton();

            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "INV:");
            
            iNVPage = new INVPage(driver);
            iNVPage.ClickNextButton();
            String reason="Automation test";
            iNVPage.SetSubStatusReason(reason);
            iNVPage.ClickNextButton();
            iNVPage.ClickFinishButton();

            driver = driver.SwitchTo().Window(BaseWindow1);

            driver.Navigate().Refresh();
            investigationCasePage = new InvestigationCasePage(driver);
            table = new Table(investigationCasePage.GetActivitiesSearchResultTable());
            table.ClickCell("Subject", investigationID + " case re-opened: Case finalisation - Case closed discussion", "Subject");
            
            Thread.Sleep(1000);

            InvestigationCaseManagementActivity caseManagementActivity = new InvestigationCaseManagementActivity(driver);
            Assert.AreEqual(reason, caseManagementActivity.GetDescription());                     
                               
         
        }


        [TestMethod]
        [TestProperty("TestcaseID", "6622")]
        public void ATC6622_CRMInvestigationValidateMandatoryTasksForPendingSubStatus()
        {
            //Login in as role
            User user = this.environment.GetUser(SecurityRole.InvestigationsOfficer);

            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsCasesRibbonButton();

            // Create new investigation case
            InvestigationCaseSearchPage investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationCaseSearchPage.ClickNewInvestigationCaseButton();

            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            string BaseWindow = driver.CurrentWindowHandle;
            investigationCasePage.ClickPageTitle();
            investigationCasePage.SetSubStatus("Case pending");
            investigationCasePage.ClickPageTitle();
            investigationCasePage.ClickSaveButton();
            string investigationID = investigationCasePage.GetInvestigationCaseNumber();

            investigationCasePage.ClickSaveCloseButton();

            // Search for the Investigation Case ID
            investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationCaseSearchPage.SetSearchRecord(investigationID);

            Table table = new Table(investigationCaseSearchPage.GetSearchResultTable());
            table.ClickCellValue("Case Number", investigationID, "Case Number");

            driver = driver.SwitchTo().Window(BaseWindow);
            investigationCasePage.ClickPageTitle();
            investigationCasePage.SetPageFilterList("All Activities");
            Thread.Sleep(1000);
            table = new Table(investigationCasePage.GetActivitiesSearchResultTable());

            // Verify that we are having 2 completed and 2 open activities
            StringAssert.Contains(table.GetCellContainsValue("Subject", investigationID + ": Set future date", "Activity Status"), "Open");
            Assert.AreEqual( 1, table.GetRowCount() - 1 , "Additional tasks are created for Case pending status");

        }

        [TestMethod]
        [TestProperty("TestcaseID", "9200")]
        public void ATC9200_CRMInvestigationClosingACaseWhichDoesNotHaveCasePartyShouldNotGiveError()
        {
            //Login in as role
            User user = this.environment.GetUser(SecurityRole.Investigations);

            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();

            homePage.ClickInvestigationsCasesRibbonButton();

            // Create new investigation case
            InvestigationCaseSearchPage investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationCaseSearchPage.ClickNewInvestigationCaseButton();

            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickSaveButton();
            string investigationID = investigationCasePage.GetInvestigationCaseNumber();
            investigationCasePage.ClickStartDialogButton();

            Table table = new Table(investigationCasePage.GetProcessSearchResultTable());
            table.ClickCell("Process Name", "INV: Close investigation case", "Created On");
            string BaseWindow = driver.CurrentWindowHandle; //Records the current window handle
            investigationCasePage.ClickDialogAddButton();

            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "INV:");
            INVPage iNVPage = new INVPage(driver);
            iNVPage.ClickNextButton();
            iNVPage.ClickNextButton();
            iNVPage.ClickNextButton();

            iNVPage.ClickFinishButton();

            driver = driver.SwitchTo().Window(BaseWindow);

            investigationCasePage.ClickSaveCloseButton();

            investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);

            investigationCaseSearchPage.SetPageFilterList("Inactive Investigation Cases");


            table = new Table(investigationCaseSearchPage.GetHeaderSearchResultTable());
            table.ClickTableColumnHeader("Case Number");
            table.ClickTableColumnHeader("Case Number");

            table = new Table(investigationCaseSearchPage.GetSearchResultTable());
            table.ClickCellValue("Case Number", investigationID, "Case Number");

            Thread.Sleep(1000);

            investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickStartDialogButton();

            table = new Table(investigationCasePage.GetProcessSearchResultTable());
            table.ClickCell("Process Name", "INV: Re-open investigation case", "Created On");
            string HomeWindow = driver.CurrentWindowHandle;

            investigationCasePage.ClickDialogAddButton();

            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "INV:");

            iNVPage = new INVPage(driver);
            iNVPage.ClickNextButton();

            iNVPage.ClickNextButton();
            iNVPage.ClickFinishButton();

            driver = driver.SwitchTo().Window(HomeWindow);

            driver.Navigate().Refresh();
            Thread.Sleep(1000);
        }

        [TestMethod]
        [TestProperty("TestcaseID", "6833")]
        public void ATC6833_CRMInvestigationValidateDurationWithReceivedDateSet()
        {
            //Login in as role
            User user = this.environment.GetUser(SecurityRole.InvestigationsOfficer);

            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsCasesRibbonButton();

            // Create new investigation case
            InvestigationCaseSearchPage investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationCaseSearchPage.ClickNewInvestigationCaseButton();

            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            string BaseWindow = driver.CurrentWindowHandle; //Records the current window handles
            investigationCasePage.ClickPageTitle();
            investigationCasePage.ClickSaveButton();

            // Set the Received Date to 4 days before todays date
            string investigationID = investigationCasePage.GetInvestigationCaseNumber();
            investigationCasePage.SetReceivedDateValue(DateTime.Today.AddDays(-4).ToString("dd-MM-yyyy"));
            investigationCasePage.ClickPageTitle();
            investigationCasePage.ClickSaveButton();
            Thread.Sleep(1000);

            // Now close the investigation case
            investigationCasePage.ClickStartDialogButton();
            Table table = new Table(investigationCasePage.GetProcessSearchResultTable());
            table.ClickCell("Process Name", "INV: Close investigation case", "Created On");
            BaseWindow = driver.CurrentWindowHandle; //Records the current window handle
            investigationCasePage.ClickDialogAddButton();

            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "INV:");
            
            INVPage iNVPage = new INVPage(driver);
            iNVPage.ClickNextButton();
            iNVPage.ClickNextButton();
            iNVPage.SetDate(DateTime.Today.ToString("dd-MM-yyyy"));
            iNVPage.ClickNextButton();
            iNVPage.ClickFinishButton();

            driver = driver.SwitchTo().Window(BaseWindow);

            // Now Investigation case is closed.. Verify that this case is not seen in active Investigation case
            investigationCasePage.ClickPageTitle();
            investigationCasePage.ClickSaveCloseButton();
            
            investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);

            investigationCaseSearchPage.SetPageFilterList("Inactive Investigation Cases");
            investigationCaseSearchPage.SetSearchRecord(investigationID);
            table = new Table(investigationCaseSearchPage.GetSearchResultTable());
            table.ClickCellValue("Case Number", investigationID, "Case Number");

            Thread.Sleep(1000);

            investigationCasePage = new InvestigationCasePage(driver);    
            investigationCasePage.ClickPageTitle();

            // Verify the Duration field value... 
            Assert.AreEqual(4, investigationCasePage.GetDurationDaysValue(), "Duration field showing wrong days");
        }

        [TestMethod]
        [TestProperty("TestcaseID", "6834")]
        public void ATC6834_CRMInvestigationValidateDurationWithoutReceivedDateSet()
        {
            //Login in as role
            User user = this.environment.GetUser(SecurityRole.InvestigationsOfficer);

            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsCasesRibbonButton();

            // Create new investigation case
            InvestigationCaseSearchPage investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationCaseSearchPage.ClickNewInvestigationCaseButton();

            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            string BaseWindow = driver.CurrentWindowHandle; 
            investigationCasePage.ClickPageTitle();
            investigationCasePage.ClickSaveButton();

            string investigationID = investigationCasePage.GetInvestigationCaseNumber();

            // Now close the investigation case
            investigationCasePage.ClickStartDialogButton();
            Table table = new Table(investigationCasePage.GetProcessSearchResultTable());
            table.ClickCell("Process Name", "INV: Close investigation case", "Created On");
            BaseWindow = driver.CurrentWindowHandle; //Records the current window handle
            investigationCasePage.ClickDialogAddButton();

            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "INV:");

            INVPage iNVPage = new INVPage(driver);
            iNVPage.ClickNextButton();
            iNVPage.ClickNextButton();
            iNVPage.SetDate(DateTime.Today.ToString("dd-MM-yyyy"));
            iNVPage.ClickNextButton();
            iNVPage.ClickFinishButton();

            driver = driver.SwitchTo().Window(BaseWindow);

            // Now Investigation case is closed.. Verify that this case is not seen in active Investigation case
            investigationCasePage.ClickPageTitle();
            investigationCasePage.ClickSaveCloseButton();

            investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);

            investigationCaseSearchPage.SetPageFilterList("Inactive Investigation Cases");
            investigationCaseSearchPage.SetSearchRecord(investigationID);
            table = new Table(investigationCaseSearchPage.GetSearchResultTable());
            table.ClickCellValue("Case Number", investigationID, "Case Number");

            Thread.Sleep(100);

            investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickPageTitle();

            // Verify the Duration field value... 
            Assert.AreEqual(0, investigationCasePage.GetDurationDaysValue(), "Duration field showing wrong days");
        }

        [TestMethod]
        [TestProperty("TestcaseID", "3352")]
        public void ATC3352_CRMInvestigationValidateDurationClearedOnReactivation()
        {
            //Login in as role
            User user = this.environment.GetUser(SecurityRole.InvestigationsOfficer);

            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsCasesRibbonButton();

            // Create new investigation case
            InvestigationCaseSearchPage investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationCaseSearchPage.ClickNewInvestigationCaseButton();

            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            string BaseWindow = driver.CurrentWindowHandle; 
            investigationCasePage.ClickPageTitle();
            investigationCasePage.SetReceivedDateValue(DateTime.Today.AddDays(-1).ToString("dd-MM-yyyy"));
            investigationCasePage.ClickSaveButton();

            // Set the Received Date to 2 days before todays date
            string investigationID = investigationCasePage.GetInvestigationCaseNumber();

            // Now close the investigation case
            investigationCasePage.ClickStartDialogButton();
            Table table = new Table(investigationCasePage.GetProcessSearchResultTable());
            table.ClickCell("Process Name", "INV: Close investigation case", "Created On");
            BaseWindow = driver.CurrentWindowHandle; 
            investigationCasePage.ClickDialogAddButton();

            
            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "INV:");
            INVPage iNVPage = new INVPage(driver);
            iNVPage.ClickNextButton();
            iNVPage.ClickNextButton();
            iNVPage.SetDate(DateTime.Today.ToString("dd-MM-yyyy"));
            iNVPage.ClickNextButton();
            iNVPage.ClickFinishButton();

            driver = driver.SwitchTo().Window(BaseWindow);

            // Now Investigation case is closed.. Verify that this case is not seen in active Investigation case
            investigationCasePage.ClickPageTitle();
            investigationCasePage.ClickSaveCloseButton();

            investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);

            investigationCaseSearchPage.SetPageFilterList("Inactive Investigation Cases");
            investigationCaseSearchPage.SetSearchRecord(investigationID);
            table = new Table(investigationCaseSearchPage.GetSearchResultTable());
            table.ClickCellValue("Case Number", investigationID, "Case Number");

            Thread.Sleep(100);

            investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickPageTitle();
            
            // Verify the Duration field... 
            Assert.AreEqual(1, investigationCasePage.GetDurationDaysValue(), "Duration field showing wrong days");

            // Reopen the Investigation case
            investigationCasePage.ClickStartDialogButton();

            table = new Table(investigationCasePage.GetProcessSearchResultTable());
            table.ClickCell("Process Name", "INV: Re-open investigation case", "Created On");
            BaseWindow = driver.CurrentWindowHandle; 
            investigationCasePage.ClickDialogAddButton();

            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "INV:");
            iNVPage = new INVPage(driver);
            iNVPage.ClickNextButton();
            iNVPage.ClickNextButton();
            iNVPage.ClickFinishButton();

            driver = driver.SwitchTo().Window(BaseWindow);

            investigationCasePage.ClickPageTitle();
            investigationCasePage.ClickSaveCloseButton();

            investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);

            investigationCaseSearchPage.SetPageFilterList("Inactive Investigation Cases");
            investigationCaseSearchPage.SetSearchRecord(investigationID);
            table = new Table(investigationCaseSearchPage.GetSearchResultTable());
            table.ClickCellValue("Case Number", investigationID, "Case Number");

            Thread.Sleep(100);

            investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickPageTitle();

            Assert.AreEqual(-1, investigationCasePage.GetDurationDaysValue(), "Duration field showing wrong days");
        }

        [TestMethod]
        [TestProperty("TestcaseID", "3314")]
        public void ATC3314_CRMInvestigationCaseManagementActivityNotVisibleInActivityMenu()
        {
            User user = this.environment.GetUser(SecurityRole.Investigations);

            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.ClickNewActivityIMG();

            Assert.IsFalse(homePage.CheckNewActivityContents("Case Management Activity"));

            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();

            homePage.ClickInvestigationsCasesRibbonButton();

            InvestigationCaseSearchPage investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationCaseSearchPage.ClickNewInvestigationCaseButton();

            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickSaveButton();

            Thread.Sleep(1000);

            investigationCasePage.ClickActivitiesAddButton();

            Assert.IsFalse(investigationCasePage.CheckAddNewActivityContents("Case Management Activity"));

        }
        [TestMethod]
        [TestProperty("TestcaseID", "3311")]
        [TestProperty("TestType", "Regression")]
        public void ATC3311_CRMInvestigationVerifyActionFollowupDateRemoval()
        {

            User user = this.environment.GetUser(SecurityRole.InvestigationsOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsCasesRibbonButton();

            InvestigationCaseSearchPage investigationsCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationsCaseSearchPage.ClickNewInvestigationCaseButton();

            // Create new case and validate Status and Substatus fields
            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickSaveButton();
            String caseNumber = investigationCasePage.GetInvestigationCaseNumber();
            investigationCasePage.ClickSaveCloseButton();

            // Search for the newly created Investigation Case
            investigationsCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationsCaseSearchPage.SetInvestigationSearchText(caseNumber);
            Table table = new Table(investigationsCaseSearchPage.GetSearchResultTable());
            table.ClickCellValue("Case Number", caseNumber, "Case Number");

            // Validate Action Date and Followup date fields are removed
            investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickPageTitle();

            Assert.IsFalse(investigationCasePage.VerifyElementExists("#rta_action_date_c"), "Action Date Present with CssValue rta_action_date_c");
            Assert.IsFalse(investigationCasePage.VerifyElementExists("#rta_inv_action_dateid_c"), "Action Date Present with CssValue rta_inv_action_dateid_c");
            Assert.IsFalse(investigationCasePage.VerifyElementExists("#rta_investigation_action_dateid_c"), "Action Date Present with CssValue rta_investigation_action_dateid_c");
            Assert.IsFalse(investigationCasePage.VerifyElementExists("#action_date_c"), "Action Date Present with CssValue action_date_c");
            Assert.IsFalse(investigationCasePage.VerifyElementExists("#rta_action_dateid_c"), "Action Date Present with CssValue rta_action_dateid_c");

            Assert.IsFalse(investigationCasePage.VerifyElementExists("#rta_followup_date_c"), "Follow up date field present with CssValue: rta_followup_date_c");
            Assert.IsFalse(investigationCasePage.VerifyElementExists("#rta_inv_followup_dateid_c"), "Follow up date field present with CssValue: rta_inv_followup_dateid_c");
            Assert.IsFalse(investigationCasePage.VerifyElementExists("#rta_investigation_followup_dateid_c"), "Follow up date field present with CssValue: rta_investigation_followup_dateid_c");
            Assert.IsFalse(investigationCasePage.VerifyElementExists("#followup_date_c"), "Follow up date field present with CssValue: followup_date_c");
            Assert.IsFalse(investigationCasePage.VerifyElementExists("#rta_followup_dateid_c"), "Follow up date field present with CssValue: rta_followup_dateid_c");
        }

        [TestMethod]
        [TestProperty("TestcaseID", "5428")]
        [TestProperty("TestType", "Regression")]
        public void ATC5428_CRMInvestigationVerifySupportOfficerQueue()
        {

            User user = this.environment.GetUser(SecurityRole.InvestigationsOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            String HomeWindow = driver.CurrentWindowHandle;
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsCasesRibbonButton();

            InvestigationCaseSearchPage investigationsCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationsCaseSearchPage.ClickNewInvestigationCaseButton();

            // Create new case and validate Status and Substatus fields
            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickSaveButton();
            String caseNumber = investigationCasePage.GetInvestigationCaseNumber();
            String BaseWindow = driver.CurrentWindowHandle;
            
            // Add the case to Support Officers Queue
            investigationCasePage.ClickAddToQueueButton();
            investigationCasePage.SetQueue("Investigations Support Officers");
            investigationCasePage.ClickDialogAddButton();

            driver = driver.SwitchTo().Window(BaseWindow);
            investigationCasePage.ClickPageTitle();
            investigationCasePage.ClickSaveCloseButton();

            // Verify the Support Officers queue
            driver.SwitchTo().Window(HomeWindow);
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsQueuesRibbonButton();

            // Verify the "Support Officers queue"
            QueueSearchPage investigationQueueSearchPage = new QueueSearchPage(driver);
            investigationQueueSearchPage.SetSearchRecord(caseNumber);
            investigationQueueSearchPage.SetQueue("Investigations Support Officers");

            Table table = new Table (investigationQueueSearchPage.GetSearchResultTable());
            Assert.AreEqual(1, table.GetRowCount(), "Investigation Case is not added to Support Officers Queue");
        }
        [TestMethod]
        [TestProperty("TestcaseID", "5429")]
        [TestProperty("TestType", "Regression")]
        public void ATC5429_CRMInvestigationVerifyInitialAssessmentQueue()
        {

            User user = this.environment.GetUser(SecurityRole.Investigations);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            String HomeWindow = driver.CurrentWindowHandle;
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsCasesRibbonButton();

            InvestigationCaseSearchPage investigationsCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationsCaseSearchPage.ClickNewInvestigationCaseButton();

            // Create new case and validate Status and Substatus fields
            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickSaveButton();
            String caseNumber = investigationCasePage.GetInvestigationCaseNumber();
            String BaseWindow = driver.CurrentWindowHandle;

            // Add the case to Support Officers Queue
            investigationCasePage.ClickAddToQueueButton();
            investigationCasePage.SetQueue("Investigations Initial Assessment");
            investigationCasePage.ClickDialogAddButton();

            driver = driver.SwitchTo().Window(BaseWindow);
            investigationCasePage.ClickSaveCloseButton();

            // Close the current window and login with Investigation Officer
            driver.Close();
            driver = null;

            this.TestSetup();

            user = this.environment.GetUser(SecurityRole.InvestigationsOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            // Verify the Support Officers queue
            homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsQueuesRibbonButton();

            // Verify the "Initial Assessment queue"
            QueueSearchPage investigationQueueSearchPage = new QueueSearchPage(driver);
            investigationQueueSearchPage.SetPageFilterList("All Items");
            investigationQueueSearchPage.SetQueue("Investigations Initial Assessment");  // Need to check with Paul
            Table table = new Table(investigationQueueSearchPage.GetHeaderSearchResultTable());
            table.ClickTableColumnHeader("Entered Queue");
            table.ClickTableColumnHeader("Entered Queue");
            table = new Table(investigationQueueSearchPage.GetSearchResultTable());
            table.GetCellValue("Title", caseNumber, "Title");

        }


        //[TestMethod]
        //[TestProperty("TestcaseID", "6361")]
        //[TestProperty("TestType", "Regression")]
        //public void ATC6361_CRMInvestigationValidateCasePendingMandatoryTasks()
        //{
        //    //Login in as role
        //    User user = this.environment.GetUser(SecurityRole.InvestigationsOfficer);

        //    new LoginDialog().Login(user.Id, user.Password);

        //    HomePage homePage = new HomePage(driver);
        //    homePage.HoverCRMRibbonTab();
        //    homePage.ClickInvestigationsRibbonButton();
        //    homePage.HoverInvestigationsRibbonTab();
        //    homePage.ClickInvestigationsCasesRibbonButton();

        //    // Create new investigation case
        //    InvestigationCaseSearchPage investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
        //    investigationCaseSearchPage.ClickNewInvestigationCaseButton();

        //    InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
        //    string BaseWindow = driver.CurrentWindowHandle;
        //    investigationCasePage.ClickPageTitle();
        //    investigationCasePage.SetSubStatus("Case pending");
        //    investigationCasePage.ClickPageTitle();
        //    investigationCasePage.ClickSaveButton();
        //    string investigationID = investigationCasePage.GetInvestigationCaseNumber();

        //    investigationCasePage.ClickSaveCloseButton();

        //    // Search for the Investigation Case ID
        //    investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
        //    investigationCaseSearchPage.SetSearchRecord(investigationID);

        //    Table table = new Table(investigationCaseSearchPage.GetSearchResultTable());
        //    table.ClickCellValue("Case Number", investigationID, "Case Number");

        //    driver = driver.SwitchTo().Window(BaseWindow);
        //    investigationCasePage.ClickPageTitle();
        //    investigationCasePage.SetPageFilterList("All Activities");
        //    Thread.Sleep(1000);
        //    table = new Table(investigationCasePage.GetActivitiesSearchResultTable());

        //    // Verify that we are having 2 completed and 2 open activities
        //    StringAssert.Contains(table.GetCellContainsValue("Subject", investigationID + ": Set future date", "Activity Status"), "Open");
        //    Assert.AreEqual(1, table.GetRowCount() - 1, "Additional tasks are created for Case pending status");
        //}



        [TestMethod]
        [TestProperty("TestcaseID", "4420")]
        [TestProperty("TestType", "Regression")]
        public void ATC4420_CRMInvestigationCreateNewClientViaInvestigationGroup()
        {

            //Login in as role
            User user = this.environment.GetUser(SecurityRole.Investigations);

            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickClientsRibbonButton();

            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);
            string baseWindow = driver.CurrentWindowHandle;
            clientsSearchPage.ClickNewClientButton();

            // Add new client name 
            ClientPage clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();

            // Fill in mandatory fields
            clientPage.SetClientType("Organisation");
            clientPage.SetOrganizationName("TC Investigation");
            clientPage.ClickSaveButton();

            string clientID = clientPage.GetClientID();

            clientPage.ClickSaveCloseButton();

            // Create another client
            driver = driver.SwitchTo().Window(baseWindow);

            clientsSearchPage.ClickNewClientButton();

            // Add second client
            clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();

            // Fill in mandatory fields
            clientPage.SetClientType("Organisation");
            clientPage.SetOrganizationName("TC Investigation second");
            clientPage.ClickSaveButton();

            string clientID2 = clientPage.GetClientID();

            StringAssert.Equals(clientID.Substring(0, 1), clientID2.Substring(0, 1));
            StringAssert.Equals(clientID.Substring(0, 1), "C"); //"Client ID is not starting with C"

            int clientNo1 = Convert.ToInt32(clientID.Substring(1));
            int clientNo2 = Convert.ToInt32(clientID2.Substring(1));

            Console.WriteLine(clientNo1);
            Console.WriteLine(clientNo2);

            if (clientNo2 - clientNo1 < 1)
            {
                throw new Exception("Client ID is not properly generated!!!");
            }
        }

        [TestMethod]
        [TestProperty("TestcaseID", "3308")]
        public void ATC3308_CRMInvestigationRecordPhoneCallActivityOnInvestigationsCase()
        {
            //Login in as role
            User user = this.environment.GetUser(SecurityRole.InvestigationsOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            String HomeWindow = driver.CurrentWindowHandle;
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsCasesRibbonButton();

            // Create new investigation case
            InvestigationCaseSearchPage investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationCaseSearchPage.ClickNewInvestigationCaseButton();

            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            string BaseWindow = driver.CurrentWindowHandle; //Records the current window handle
            investigationCasePage.ClickSaveButton();
            String investigationID = investigationCasePage.GetInvestigationCaseNumber();
            Thread.Sleep(1000);
            investigationCasePage.ClickActivitiesAddButton();
            investigationCasePage.ClickAddActivity("Phone Call");
            Thread.Sleep(2000);
                               
            //Add new Phone Call
            String subject = "New Phone Call";
            driver = investigationCasePage.SwitchNewBrowser(driver, BaseWindow);
            PhoneCallPage phoneCall = new PhoneCallPage(driver);
            Thread.Sleep(100);
            phoneCall.ClickPageTitle();
            phoneCall.SetSelectSubjectValue("Bond balance enquiry");
            phoneCall.SetSubject(subject);
            phoneCall.SetRecipient("BLAIR TEST");
            phoneCall.ClickSaveButton();
            phoneCall.ClickMarkCompleteButton();

            //Verify Phone Call details
            driver = driver.SwitchTo().Window(BaseWindow);
            investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickPageTitle();
            Table table = new Table(investigationCasePage.GetActivitiesSearchResultTable());
            Thread.Sleep(1000);
            Assert.AreEqual(subject,table.GetCellValue("Activity Type","Phone Call","Subject"));
            Assert.AreEqual("Completed", table.GetCellValue("Subject", subject, "Activity Status"));

            //Re-open Phone Call
            table.SelectTableRow("Activity Type", "Phone Call");
            phoneCall = new PhoneCallPage(driver);
            phoneCall.ClickStartDialogButton();
            table = new Table(phoneCall.GetProcessSearchResultTable());
            table.ClickCell("Process Name", "Re-open Phone Call Activity", "Created On");
            phoneCall.ClickDialogAddButton();
            Thread.Sleep(1000);
            
            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "Re-open Phone call Activity");

            ReOpenCall reOpenPhoneCall = new ReOpenCall(driver);
            reOpenPhoneCall.ClickNextButton();
            reOpenPhoneCall.ClickFinishButton();

            driver = driver.SwitchTo().Window(HomeWindow);
            driver.Navigate().Back();

            //Verify Reopned Phone call
            driver = driver.SwitchTo().Window(BaseWindow);
            investigationCasePage = new InvestigationCasePage(driver);
            table = new Table(investigationCasePage.GetActivitiesSearchResultTable());
            Thread.Sleep(1000);
            Assert.AreEqual("Open", table.GetCellValue("Subject", subject, "Activity Status"));

            table.SelectTableRow("Activity Type", "Phone Call");

            //Amend Phone call
            phoneCall = new PhoneCallPage(driver);
            String newSubject = "Amend Phone Call";
            phoneCall.SetSubject(newSubject);
            phoneCall.ClickSaveCloseButton();

            //Verify Amended Phone call details
            driver = driver.SwitchTo().Window(BaseWindow);
            investigationCasePage = new InvestigationCasePage(driver);
            table = new Table(investigationCasePage.GetActivitiesSearchResultTable());
            Thread.Sleep(1000);
            Assert.AreEqual(newSubject, table.GetCellValue("Activity Type", "Phone Call", "Subject"));
            table.SelectTableRow("Activity Type", "Phone Call");

            //Cancel Phone Call
            phoneCall = new PhoneCallPage(driver);
            phoneCall.ClickPageTitle();
            phoneCall.ClickClosePhoneCallButton();
            phoneCall.ConfirmDeactivation("Canceled");
            phoneCall.ClickConfirmDeactivationCloseButton();
            driver = driver.SwitchTo().Window(HomeWindow);
            driver.Navigate().Back();

            //Verify Canceled phone call details
            driver = driver.SwitchTo().Window(BaseWindow);
            investigationCasePage = new InvestigationCasePage(driver);
            table = new Table(investigationCasePage.GetActivitiesSearchResultTable());
            Assert.AreEqual("Canceled", table.GetCellValue("Subject", newSubject, "Activity Status"));            
                                            


        }
        [TestMethod]
        [TestProperty("TestcaseID", "4431")]
        [TestProperty("TestType", "Regression")]
        public void ATC4431_CRMInvestigationVerifyPhysicalAddressWarningMessage()
        {

            //Login in as role
            User user = this.environment.GetUser(SecurityRole.Investigations);

            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickClientsRibbonButton();

            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.ClickNewClientButton();

            // Add new client name 
            ClientPage clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();

            // Fill in mandatory fields
            clientPage.SetClientType("Organisation");
            clientPage.SetOrganizationName("TC Investigation");
            clientPage.ClickSaveButton();
            clientPage.ClickPageTitle();

            string clientID = clientPage.GetClientID();

            // Verify that "Physical Address is blank, please select an address." error message is displayed
            Assert.IsTrue(clientPage.VerifyWarningMessagePresent("rta_physicaladdressid"), "Physical Address blank message NOT displayed");
            StringAssert.Contains(clientPage.GetWarningMessage("rta_physicaladdressid"), "Physical Address is blank, please select an address.");

            // Enter valid Physical address
            string BaseWindow = driver.CurrentWindowHandle;

            // Set the postal address
            clientPage.ClickCreateNewClientAddressButton("rta_physicaladdressid");

            
            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "Client Address");
            ClientNewAddressPage clientNewAddressPage = new ClientNewAddressPage(driver);
            clientNewAddressPage.SetAddressDetails("Australian Physical", 10, "GRACELAND");

            driver.SwitchTo().Window(BaseWindow);
            clientPage = new ClientPage(driver);

            //Verify Address value
            Assert.AreEqual("10 GRACELAND", clientPage.GetAddressValue("rta_physicaladdressid"));
            Assert.IsFalse(clientPage.VerifyWarningMessagePresent("rta_physicaladdressid"), "Physical Address blank message DISPLAYED!!!!!");
        }

        [TestMethod]
        [TestProperty("TestcaseID", "4448")]
        [TestProperty("TestType", "Regression")]
        public void ATC4448_CRMInvestigationVerifyFilterOnInvestigatorName()
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

            string clientName = MyRange.Cells[MyRow, InvestigationSchema.GetColumnIndex(ColumnName.CLIENT_NAME)].Value;

            //Login in as role
            User user = this.environment.GetUser(SecurityRole.Investigations);

            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            string HomeWindow = driver.CurrentWindowHandle;
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsCasesRibbonButton();

            // Create investigation case 1 with Investigator as IMSTestU06"
            InvestigationCaseSearchPage investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
            string BaseWindow = driver.CurrentWindowHandle;
            investigationCaseSearchPage.ClickNewInvestigationCaseButton();

            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            string user06 = "IMSTestU06";
            string user04 = "IMSTestU04";

            investigationCasePage.ClickPageTitle();
            Assert.IsTrue(investigationCasePage.GetInvestigatorSearchElementText(user06));
            investigationCasePage.ClickSaveButton();
            string investigationID1 = investigationCasePage.GetInvestigationCaseNumber();
            Console.WriteLine(investigationID1);
            investigationCasePage.ClickSaveCloseButton();

            // Create investigation case 2 with Investigator as IMSTestU06"
            driver = driver.SwitchTo().Window(BaseWindow);
            investigationCaseSearchPage.ClickNewInvestigationCaseButton();

            investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickPageTitle();
            Assert.IsTrue(investigationCasePage.GetInvestigatorSearchElementText(user06));
            investigationCasePage.ClickSaveButton();
            string investigationID2 = investigationCasePage.GetInvestigationCaseNumber();
            Console.WriteLine(investigationID2);
            investigationCasePage.ClickSaveCloseButton();
            Thread.Sleep(1000);

            // Create one general case where owner = same common investigator from above point
            driver = driver.SwitchTo().Window(HomeWindow);
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsGeneralCasesRibbonButton();

            InvestigationGeneralCaseSearchPage investigationsGeneralCasesPage = new InvestigationGeneralCaseSearchPage(driver);
            investigationsGeneralCasesPage.ClickNewGeneralCaseButton();

            InvestigationGeneralCasePage investigationGeneralCasePage = new InvestigationGeneralCasePage(driver);
            investigationGeneralCasePage.ClickPageTitle();
            investigationGeneralCasePage.SetTitle("Client D");
            investigationGeneralCasePage.SetClientName(clientName);
            investigationGeneralCasePage.SetType("Enquiry");
            investigationGeneralCasePage.SetInvestigatorSearchElementText(user06);
            investigationGeneralCasePage.ClickSaveButton();

            string generalCaseID = investigationGeneralCasePage.GetGeneralCaseNumber();
            Console.WriteLine(generalCaseID);
            investigationGeneralCasePage.ClickSaveCloseButton();

            // Ensure an investigation master case exists with the same common {Investigator} as above
            driver = driver.SwitchTo().Window(HomeWindow);
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsMasterCasesRibbonButton();

            InvestigationMasterCaseSearchPage investigationMasterCasesSearchPage = new InvestigationMasterCaseSearchPage(driver);
            investigationMasterCasesSearchPage.ClickNewButton();

            InvestigationMasterCasePage investigationMasterCasePage = new InvestigationMasterCasePage(driver);
            investigationMasterCasePage.SetInvestigatorValue(user06);
            investigationMasterCasePage.ClickSaveButton();
            string masterCasrID = investigationMasterCasePage.GetInvestigationMasterCaseNumber();
            Console.WriteLine(masterCasrID);
            // Close the current window and login with Investigation Business Admin
            driver.Close();
            driver = null;

            this.TestSetup();

            user = this.environment.GetUser(SecurityRole.InvestigationsBusinessAdmin);
            new LoginDialog().Login(user.Id, user.Password);

            // Verify the Support Officers queue
            homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsCasesRibbonButton();

            // Ensure at least one other investigations case exists with a different {Investigator} but {Owner} = same common investigator as above
            investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationCaseSearchPage.ClickNewInvestigationCaseButton();

            investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickPageTitle();

            Assert.IsTrue(investigationCasePage.GetInvestigatorSearchElementText(user04));
            investigationCasePage.ClickSaveButton();
            string investigationID3 = investigationCasePage.GetInvestigationCaseNumber();
            Console.WriteLine(investigationID3);
            investigationCasePage.ClickSaveCloseButton();

            investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationCaseSearchPage.SetSearchRecord("IMSTestU06");

            Table table = new Table(investigationCaseSearchPage.GetSearchResultTable());
            table.ClickCellContainsValue("Investigator", "IMSTestU06", "Investigator");

            // Verify that User oage is getting displayed
            UserPage investigationUserPage = new UserPage(driver);
            investigationUserPage.ClickPageTitle();
            StringAssert.Contains(investigationUserPage.GetFullName(), "IMSTestU06");

            //Navigate to client phone numbers
            homePage.HoverClientXRibbonTab(user06);
            homePage.ClickInvestigatorXCasesRibbonButton();

            // Verify the cases displayed for Investigation Business Admin
            investigationUserPage = new UserPage(driver);
            StringAssert.Contains(investigationUserPage.GetPageFilterList(), "Investigation Case Associated View");

            // Verify Investigation Case 1 is displayed
            investigationUserPage = new UserPage(driver);
            investigationUserPage.SetSearchRecord(investigationID1);
            table = new Table(investigationUserPage.GetSearchResultTable());
            StringAssert.Contains(table.GetCellContainsValue("Case Number", investigationID1, "Investigator"), user06);
            StringAssert.Contains(table.GetCellContainsValue("Case Number", investigationID1, "Owner"), user04);

            // Verify Investigation Case 2 is displayed
            investigationUserPage = new UserPage(driver);
            investigationUserPage.SetSearchRecord(investigationID2);
            table = new Table(investigationUserPage.GetSearchResultTable());
            StringAssert.Contains(table.GetCellContainsValue("Case Number", investigationID2, "Investigator"), user06);
            StringAssert.Contains(table.GetCellContainsValue("Case Number", investigationID2, "Owner"), user04);

            // Verify General case is not displayed
            investigationUserPage = new UserPage(driver);
            investigationUserPage.SetSearchRecord(generalCaseID);
            table = new Table(investigationUserPage.GetSearchResultTable());
            Assert.IsTrue(table.GetNoRecordsInTable(), "General case is DISPLAYED!!!!");

            // Verify Master case is not displayed
            investigationUserPage = new UserPage(driver);
            investigationUserPage.SetSearchRecord(masterCasrID);
            table = new Table(investigationUserPage.GetSearchResultTable());
            Assert.IsTrue(table.GetNoRecordsInTable(), "Master case is DISPLAYED!!!!");

            // Verify Master case is not displayed
            investigationUserPage = new UserPage(driver);
            investigationUserPage.SetSearchRecord(investigationID3);
            table = new Table(investigationUserPage.GetSearchResultTable());
            Assert.IsTrue(table.GetNoRecordsInTable(), "Investigation case with different Investigator is DISPLAYED!!!!");

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion 
        }

        [TestMethod]
        [TestProperty("TestcaseID", "5286")]
        public void ATC5286_CRMInvestigationMasterCaseAssociatedCaseStatusChange()
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

            string clientName = MyRange.Cells[MyRow, InvestigationSchema.GetColumnIndex(ColumnName.CLIENT_NAME)].Value;

            //Login in as role
            User user = this.environment.GetUser(SecurityRole.InvestigationsOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            String HomeWindow = driver.CurrentWindowHandle;
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsCasesRibbonButton();

            // Create new investigation case
            InvestigationCaseSearchPage investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationCaseSearchPage.ClickNewInvestigationCaseButton();

            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickSaveButton();
            String investigationID1 = investigationCasePage.GetInvestigationCaseNumber();
            Thread.Sleep(1000);
            investigationCasePage.ClickSaveCloseButton();

            investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationCaseSearchPage.ClickNewInvestigationCaseButton();

            investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickSaveButton();
            String investigationID2 = investigationCasePage.GetInvestigationCaseNumber();
            Thread.Sleep(1000);
            investigationCasePage.ClickSaveCloseButton();
            investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);

            driver = driver.SwitchTo().Window(HomeWindow);
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsMasterCasesRibbonButton();

            InvestigationMasterCaseSearchPage investigationMasterCasesSearchPage = new InvestigationMasterCaseSearchPage(driver);
            string BaseWindow = driver.CurrentWindowHandle; 

            investigationMasterCasesSearchPage.ClickNewButton();

            InvestigationMasterCasePage investigationMasterCasePage = new InvestigationMasterCasePage(driver);
            string MasterWindow = driver.CurrentWindowHandle;
            investigationMasterCasePage.SetClientValue(clientName);
            investigationMasterCasePage.ClickSaveButton();

            String investigationMasterID = investigationMasterCasePage.GetInvestigationMasterCaseNumber();

            investigationMasterCasePage.ClickInvestigationCaseAddButton();
            investigationMasterCasePage.SetInvestigationCaseNumberToAssociateMaster(investigationID1);

            investigationMasterCasePage.ClickInvestigationCaseAddButton();
            investigationMasterCasePage.SetInvestigationCaseNumberToAssociateMaster(investigationID2);

            // Confirm user can view the current status and sub status of each linked Investigation Case
            Table table = new Table(investigationMasterCasePage.GetInvestigationCasesSearchResultTable());
            Assert.AreEqual("New case", table.GetCellValue("Case Number", investigationID1, "Investigation Status"));
            Assert.AreEqual("Creation", table.GetCellValue("Case Number", investigationID1, "Investigation Sub Status"));

            Assert.AreEqual("New case", table.GetCellValue("Case Number", investigationID2, "Investigation Status"));
            Assert.AreEqual("Creation", table.GetCellValue("Case Number", investigationID2, "Investigation Sub Status"));

            table.ClickCell("Case Number", investigationID2, "Case Number");

            investigationMasterCasePage.ClickInvestigationCaseAssociatedView();

            driver = driver.SwitchTo().Window(BaseWindow);

            // Confirm 'Investigation Case Associated View' displayed
            Assert.AreEqual("Investigation Case Associated View", investigationMasterCasePage.GetCurrentView());

            // Select associated Investigation Case to update 
            investigationMasterCasePage.SwitchToMasterCasePageFrame();

            table = new Table(investigationMasterCasePage.GetInvestigationCasesAssociatedViewTable());
            table.ClickCellValue("Case Number", investigationID2, "Case Number");

            investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickPageTitle();
            string CaseWindow = driver.CurrentWindowHandle; 
            investigationCasePage.ClickRunWorkflowButton();

            // From 'Run Workflow' select 'Update Status: New Case - Creation' process

            table = new Table(investigationCasePage.GetProcessSearchResultTable());
            table.ClickCell("Process Name", "Update Status: New Case - Creation", "Created On");
            
            investigationCasePage.ClickDialogAddButton();
            Thread.Sleep(400);
            investigationCasePage.ClickConfirmApplicationOfWindow(BaseWindow);

            driver = driver.SwitchTo().Window(CaseWindow);
            Thread.Sleep(400);
            investigationCasePage.ClickSaveCloseButton();

            investigationMasterCasePage = new InvestigationMasterCasePage(driver);
            investigationMasterCasePage.ClickSaveCloseButton();

            investigationMasterCasesSearchPage = new InvestigationMasterCaseSearchPage(driver);
            investigationMasterCasesSearchPage.SearchRecord(investigationMasterID);
            table = new Table(investigationMasterCasesSearchPage.GetSearchResultTable());
            table.ClickCellContainsValueEnterRow("Master Case ID", investigationMasterID, "Master Case ID");

            investigationMasterCasePage = new InvestigationMasterCasePage(driver);
            table = new Table(investigationMasterCasePage.GetInvestigationCasesSearchResultTable());   
            table.ClickCellValue("Case Number", investigationID2, "Case Number");

            // Confirm that Status - sub statuses updated to 'New Case - Initial assessment'

            investigationCasePage = new InvestigationCasePage(driver);
            StringAssert.Contains(investigationCasePage.GetStatus(), "New case");
            StringAssert.Contains(investigationCasePage.GetSubStatus(), "Initial assessment");

            table = new Table(investigationCasePage.GetActivitiesHeaderTable());
            table.ClickTableColumnHeader("Subject");

            table = new Table(investigationCasePage.GetActivitiesSearchResultTable());

            // Confirm mandatory tasks for Investigation Cases updated are created for that status and sub status

            StringAssert.Contains(table.GetCellContainsValue("Subject", 1), investigationID2 + ": Add parties");
            StringAssert.Contains(table.GetCellContainsValue("Subject", 2), investigationID2 + ": Allocate");
            StringAssert.Contains(table.GetCellContainsValue("Subject", 3), investigationID2 + ": Carry out initial assessment");
            StringAssert.Contains(table.GetCellContainsValue("Subject", 4), investigationID2 + ": Scan documents");
            StringAssert.Contains(table.GetCellContainsValue("Subject", 5), investigationID2 + ": Submit for initial assessment");

            Assert.AreEqual(5, table.GetRowCount() - 1, "Additional activities are displayed!!!!");

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion 
        }

        [TestMethod]
        [TestProperty("TestcaseID", "3310")]
        [TestProperty("TestType", "Regression")]
        public void ATC3310_CRMInvestigationAllCasesViewColumns()
        {
            //Login in as role
            User user = this.environment.GetUser(SecurityRole.Investigations);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            String HomeWindow = driver.CurrentWindowHandle;
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsCasesRibbonButton();

            // Select the 'All Investigation Cases' view from the table
            InvestigationCaseSearchPage investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationCaseSearchPage.SetPageFilterList("All Investigation Cases");


            Table table = new Table(investigationCaseSearchPage.GetHeaderSearchResultTable());
            Assert.AreEqual(1, table.GetColumnHeaderIndex("Case Number"), "Case Number column is not seen or is not 1st column!!");
            Assert.AreEqual(2, table.GetColumnHeaderIndex("Master Case"), "Master Case column is not seen or is not 2ndt column!!");
            Assert.AreEqual(3, table.GetColumnHeaderIndex("Investigation Status"), "Investigation Status column is not seen or is not 3rd column!!");
            Assert.AreEqual(4, table.GetColumnHeaderIndex("Investigation Sub Status"), "Investigation Sub Status column is not seen or is not 4th column!!");
            Assert.AreEqual(5, table.GetColumnHeaderIndex("Priority"), "Priority column is not seen or is not 5th column!!");
            Assert.AreEqual(6, table.GetColumnHeaderIndex("Received Date"), "Received Date column is not seen or is not 6th column!!");
            Assert.AreEqual(7, table.GetColumnHeaderIndex("Finalised Date"), "Finalised Date column is not seen or is not 7th column!!");
            Assert.AreEqual(8, table.GetColumnHeaderIndex("Duration (days)"), "Duration (days) column is not seen or is not 8th column!!");
            Assert.AreEqual(9, table.GetColumnHeaderIndex("Investigator"), "Investigator column is not seen or is not 9th column!!");
            Assert.AreEqual(10, table.GetColumnHeaderIndex("Owner"), "Owner column is not seen or is not 10th column!!");

            Assert.AreEqual(0, table.GetColumnHeaderIndex("Action Date"), "Action Date column is seen!!");
            Assert.AreEqual(0, table.GetColumnHeaderIndex("Follow Up Date"), "Follow Up Date column is seen!!");
            Assert.AreEqual(0, table.GetColumnHeaderIndex("Title"), "Title column is seen!!");
        }

        [TestMethod]
        [TestProperty("TestcaseID", "3326")]
        [TestProperty("TestType", "Regression")]
        public void ATC3326_CRMInvestigationUserAccessToRecordOutcomeOfAllegedOffence()
        {
            string[] allowedUserName = { "Investig ationsOfficer", "Investigations", 
                                           "InvestigationsManager", "InvestigationsBusinessAdmin", 
                                           "InvestigationsOfficer", "ESOForPES", "ExecutiveManagerForPES" };
            string[] notAllowedUserName = { "GeneralStaff", "RBSOfficer","IMSBusinessSupportStaff" ,"ResearchOfficers","RecordKeepingOfficers"};
            bool LastIteration = false;

            // CAN
            User user = this.environment.GetUser(SecurityRole.InvestigationOfficer);
            for (int i = 0; i < allowedUserName.Length; i++)
            {

                //Login in as role
                switch (i)
                {
                    case 0:
                        user = this.environment.GetUser(SecurityRole.InvestigationOfficer);
                        break;
                    case 1:
                        user = this.environment.GetUser(SecurityRole.Investigations);
                        break;
                    case 2:
                        user = this.environment.GetUser(SecurityRole.InvestigationsManager);
                        break;
                    case 3:
                        user = this.environment.GetUser(SecurityRole.InvestigationsBusinessAdmin);
                        break;
                    case 4:
                        user = this.environment.GetUser(SecurityRole.InvestigationsOfficer);
                        break;
                    case 5:
                        user = this.environment.GetUser(SecurityRole.ESOForPES);
                        break;
                    case 6:
                        user = this.environment.GetUser(SecurityRole.ExecutiveManagerForPES);
                        LastIteration = true;
                        break;
                }
                new LoginDialog().Login(user.Id, user.Password);

                HomePage homePage = new HomePage(driver);
                String HomeWindow = driver.CurrentWindowHandle;
                homePage.HoverCRMRibbonTab();

                homePage.ClickInvestigationsRibbonButton();
                homePage.HoverInvestigationsRibbonTab();
                homePage.ClickInvestigationsCasesRibbonButton();

                // Select the 'All Investigation Cases' view from the table
                InvestigationCaseSearchPage investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
                investigationCaseSearchPage.ClickNewInvestigationCaseButton();

                InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
                investigationCasePage.ClickSaveButton();
                string investigationID = investigationCasePage.GetInvestigationCaseNumber();
                investigationCasePage.ClickSaveCloseButton();

                driver = driver.SwitchTo().Window(HomeWindow);
                homePage.HoverCRMRibbonTab();
                homePage.ClickInvestigationsRibbonButton();
                homePage.HoverInvestigationsRibbonTab();
                homePage.ClickRightScrollRibbonButton();
                homePage.ClickAllegedOffencesButton();

                AllegendOffensesSearchPage allegedOffencesSearchPage = new AllegendOffensesSearchPage(driver);

                allegedOffencesSearchPage.ClickNewAllegedOffenceButton();

                AllegedOffencePage allegedOffencesPage = new AllegedOffencePage(driver);
                allegedOffencesPage.SetInvestigationCaseValue(investigationID);
                allegedOffencesPage.SetProvisionValue("RTRA 116(1)");
                allegedOffencesPage.ClickSaveButton();
                string allegedoffenceId = allegedOffencesPage.GetReferenceNumber();
                allegedOffencesPage.ClickSaveCloseButton();

                allegedOffencesSearchPage = new AllegendOffensesSearchPage(driver);
                allegedOffencesSearchPage.SetInvestigationSearchText(allegedoffenceId);
                Table table = new Table(allegedOffencesSearchPage.GetSearchResultTable());
                Assert.AreEqual(1, table.GetRowCount(), "Allegend Offense creation failed!!!!!");

                if (LastIteration == false)
                {
                    driver.Quit();
                    driver = null;
                    this.TestSetup();
                }
                // CAN end
            }

            // CAN NOT -1
            LastIteration = false;
        

            for (int i = 0; i <= 2; i++)
            {
                //Login in as role
                switch (i)
                {
                    case 0:
                        user = this.environment.GetUser(SecurityRole.GeneralStaff);
                        break;
                    case 1:
                        user = this.environment.GetUser(SecurityRole.RBSOfficer);
                        break;
                    case 2:
                        user = this.environment.GetUser(SecurityRole.IMSBusinessSupportStaff);
                        LastIteration = true;
                        break;
                }
                new LoginDialog().Login(user.Id, user.Password);

                HomePage homePage = new HomePage(driver);
                String HomeWindow = driver.CurrentWindowHandle;
                homePage.HoverCRMRibbonTab();
                Assert.IsFalse(homePage.VerifyInvestigationsRibbonButtonPresent(),String.Format("Investigation Ribbon Button is seen for {0}",user.Id));

                if (LastIteration == false)
                {
                    driver.Quit();
                    driver = null;
                    this.TestSetup();
                }
            }

            // CAN NOT -2
            LastIteration = false;
            for (int i = 3; i == 4; i++)
            {
                //Login in as role
                switch (i)
                {
                    case 3:
                        user = this.environment.GetUser(SecurityRole.ResearchOfficers);
                        break;
                    case 4:
                        user = this.environment.GetUser(SecurityRole.RecordKeepingOfficers);
                        LastIteration = true;
                        break;
                }
                new LoginDialog().Login(user.Id, user.Password);

                HomePage homePage = new HomePage(driver);
                String HomeWindow = driver.CurrentWindowHandle;
                homePage.HoverCRMRibbonTab();
                homePage.ClickInvestigationsRibbonButton();
                homePage.HoverInvestigationsRibbonTab();
                homePage.ClickInvestigationsCasesRibbonButton();

                // Select the 'All Investigation Cases' view from the table
                InvestigationCaseSearchPage investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);

                Assert.IsFalse(investigationCaseSearchPage.VerifyNewInvestigationCaseButtonPresent(), String.Format("New Investigation Case Button is seen for {0}", user.Id));

                if (LastIteration == false)
                {
                    driver.Quit();
                    driver = null;
                    this.TestSetup();
                }
            }
        }



        [TestMethod]
        [TestProperty("TestcaseID", "3329")]
        [TestProperty("TestType", "Regression")]
        public void ATC3329_CRMInvestigationVerifyProactiveDisputeResolutionFieldRemoval()
        {

            User user = this.environment.GetUser(SecurityRole.InvestigationsOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsCasesRibbonButton();

            InvestigationCaseSearchPage investigationsCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationsCaseSearchPage.ClickNewInvestigationCaseButton();

            // Create new case and validate Status and Substatus fields
            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickSaveButton();
            String caseNumber = investigationCasePage.GetInvestigationCaseNumber();
            investigationCasePage.ClickSaveCloseButton();

            // Search for the newly created Investigation Case
            investigationsCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationsCaseSearchPage.SetInvestigationSearchText(caseNumber);
            Table table = new Table(investigationsCaseSearchPage.GetSearchResultTable());
            table.ClickCellValue("Case Number", caseNumber, "Case Number");

            // Validate Action Date and Followup date fields are removed
            investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickPageTitle();

            // Proactive/Reative   
            Assert.IsFalse(investigationCasePage.VerifyElementExists("#rta_proactive_reactive_c"), "Proactive/Reative field Present with CssValue rta_proactive_reactive_c");
            Assert.IsFalse(investigationCasePage.VerifyElementExists("#rta_inv_proactive_reactiveid_c"), "Proactive/Reative field Present with CssValue rta_inv_proactive_reactiveid_c");
            Assert.IsFalse(investigationCasePage.VerifyElementExists("#rta_investigation_proactive_reactive_c"), "Proactive/Reative field Present with CssValue rta_investigation_proactive_reactive_c");
            Assert.IsFalse(investigationCasePage.VerifyElementExists("#rta_investigation_proactive_reactiveid_c"), "Proactive/Reative field Present with CssValue rta_investigation_proactive_reactiveid_c");
            Assert.IsFalse(investigationCasePage.VerifyElementExists("#proactive_reactive_c"), "Proactive/Reative field Present with CssValue proactive_reactive_c");
            Assert.IsFalse(investigationCasePage.VerifyElementExists("#rta_proactive_reactive_c"), "Proactive/Reative field Present with CssValue rta_proactive_reactive_c");
            Assert.IsFalse(investigationCasePage.VerifyElementExists("#rta_proactive_reactiveid_c"), "Proactive/Reative field Present with CssValue rta_proactive_reactiveid_c");

            // Dispute Resolution Case 
            Assert.IsFalse(investigationCasePage.VerifyElementExists("#rta_dispute_resolution_case_c"), "Dispute Resolution Case field present with CssValue: rta_dispute_resolution_case_c");
            Assert.IsFalse(investigationCasePage.VerifyElementExists("#rta_inv_dispute_resolution_caseid_c"), "Dispute Resolution Case field present with CssValue: rta_inv_dispute_resolution_caseid_c");
            Assert.IsFalse(investigationCasePage.VerifyElementExists("#rta_investigation_dispute_resolution_caseid_c"), "Dispute Resolution Case field present with CssValue: rta_investigation_dispute_resolution_caseid_c");
            Assert.IsFalse(investigationCasePage.VerifyElementExists("#dispute_resolution_case_c"), "Dispute Resolution Case field present with CssValue: dispute_resolution_case_c");
            Assert.IsFalse(investigationCasePage.VerifyElementExists("#rta_dispute_resolution_caseid_c"), "Dispute Resolution Case field present with CssValue: rta_dispute_resolution_caseid_c");
        }

        [TestMethod]
        [TestProperty("TestcaseID", "3323")]
        [TestProperty("TestType", "Regression")]
        public void ATC3323_CRMInvestigationVerifyPenaltyUnitsAreNotLocked()
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

            string clientName = MyRange.Cells[MyRow, InvestigationSchema.GetColumnIndex(ColumnName.CLIENT_NAME)].Value;

            User user = this.environment.GetUser(SecurityRole.InvestigationsOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsPenaltyInfringementNoticesRibbonButton();

            InvestigationPenaltyINoticeSearchPage invPenaltyINoticeSearchPage = new InvestigationPenaltyINoticeSearchPage(driver);
            invPenaltyINoticeSearchPage.ClickNewPenaltyNoticeButton();

            // Create new case and validate Status and Substatus fields
            InvestigationPenaltyINoticePage invPenaltyINoticePage = new InvestigationPenaltyINoticePage(driver);
            invPenaltyINoticePage.ClickPageTitle();
            invPenaltyINoticePage.SetClientName(clientName);
            invPenaltyINoticePage.SetIssuedAgainstField("Person");
            invPenaltyINoticePage.ClickSaveButton();
            invPenaltyINoticePage.ClickPageTitle();

            // Verify Penalty Units, Per Unit Amount and Penalty Amount fields are not locked
            Assert.IsFalse(invPenaltyINoticePage.CheckPenaltyUnitsLocked(), "Penalty Units fields is not locked");
            Assert.IsFalse(invPenaltyINoticePage.CheckPerUnitAmountLocked(), "Per Unit Amount fields is not locked");
            Assert.IsFalse(invPenaltyINoticePage.CheckPenaltyAmountFieldLocked(), "Penalty Amount fields is not locked");

            // Confirm the Penalty Units, Per Unit Amount and Penalty Amount fields have no complex interdependent validations.   
            StringAssert.Contains(invPenaltyINoticePage.GetPenaltyUnitsFieldText(), "--");
            StringAssert.Contains(invPenaltyINoticePage.GetPerUnitAmountFieldText(), "--");
            StringAssert.Contains(invPenaltyINoticePage.GetPenaltyAmountFieldText(), "--");

            invPenaltyINoticePage.SetPenaltyUnits("1000");
            invPenaltyINoticePage.SetPerUnitAmount("1234");
            invPenaltyINoticePage.SetPenaltyAmountFieldText("5678");

            invPenaltyINoticePage.ClickSaveButton();
            Thread.Sleep(500);

            StringAssert.Contains(invPenaltyINoticePage.GetPenaltyUnitsFieldText(), "1,000");
            StringAssert.Contains(invPenaltyINoticePage.GetPerUnitAmountFieldText(), "1,234");
            StringAssert.Contains(invPenaltyINoticePage.GetPenaltyAmountFieldText(), "5,678");


            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion 
        }

        [TestMethod]
        [TestProperty("TestcaseID", "3340")]
        [TestProperty("TestType", "Regression")]
        public void ATC3340_CRMInvestigationAllegedOffenseStatusChangesToBelief()
        {
            User user = this.environment.GetUser(SecurityRole.InvestigationsOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            string HomeWindow = driver.CurrentWindowHandle;
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsCasesRibbonButton();

            InvestigationCaseSearchPage investigationsCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationsCaseSearchPage.ClickNewInvestigationCaseButton();

            // Create new case and validate Status and Substatus fields
            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickSaveButton();
            String investigationID = investigationCasePage.GetInvestigationCaseNumber();
            investigationCasePage.ClickSaveCloseButton();

            driver = driver.SwitchTo().Window(HomeWindow);
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickRightScrollRibbonButton();
            homePage.ClickAllegedOffencesButton();

            AllegendOffensesSearchPage allegedOffencesSearchPage = new AllegendOffensesSearchPage(driver);

            allegedOffencesSearchPage.ClickNewAllegedOffenceButton();

            AllegedOffencePage allegedOffencesPage = new AllegedOffencePage(driver);
            allegedOffencesPage.SetInvestigationCaseValue(investigationID);
            allegedOffencesPage.SetProvisionValue("RTRA 116(1)");
            allegedOffencesPage.ClickSaveButton();
            StringAssert.Contains(allegedOffencesPage.GetStatusReason(), "Suspicion");

            string OffenceDate = DateTime.Now.AddYears(-1).ToString("d/MM/yyyy");
            allegedOffencesPage.SetOffenceDateValue(OffenceDate);
            string todaysDate = DateTime.Now.ToString("d/MM/yyyy");
            allegedOffencesPage.SetBeliefFormedDateValue(todaysDate);
            allegedOffencesPage.ClickSaveButton();

            StringAssert.Contains(allegedOffencesPage.GetStatusReason(), "Belief");
            string StatutoryDate = allegedOffencesPage.GetStatutoryLimitationValue();
            string allegedoffenceId = allegedOffencesPage.GetReferenceNumber();
            allegedOffencesPage.ClickSaveCloseButton();

            allegedOffencesSearchPage = new AllegendOffensesSearchPage(driver);
            allegedOffencesSearchPage.SetInvestigationSearchText(investigationID);
            Table table = new Table(allegedOffencesSearchPage.GetSearchResultTable());
            table.SelectTableRow("Status Reason", "Belief");

            allegedOffencesPage = new AllegedOffencePage(driver);
            StringAssert.Contains(allegedOffencesPage.GetStatusReason(), "Belief");
            StringAssert.Contains(allegedOffencesPage.GetStatutoryLimitationValue(), StatutoryDate);
            StringAssert.Contains(allegedOffencesPage.GetBefliefFormedDateValue(), todaysDate);
        }

        [TestMethod]
        [TestProperty("TestcaseID", "3345")]
        [TestProperty("TestType", "Regression")]
        public void ATC3345_CRMInvestigationVerifyUserCantEditMasterCases()
        {
            string[] NoAccessUser = { "GeneralStaff", "IMS business systems support staff" };
            User user = this.environment.GetUser(SecurityRole.GeneralStaff);
            bool lastIteration = false;
            for (int i = 0; i < 2; i++)
            {
                switch (i)
                { 
                    case 0:
                        user = this.environment.GetUser(SecurityRole.GeneralStaff);
                        break;
                    case 1:
                        user = this.environment.GetUser(SecurityRole.IMSBusinessSupportStaff);
                        lastIteration = true;
                        break;
                }
                new LoginDialog().Login(user.Id, user.Password);

                HomePage homePage = new HomePage(driver);
                string HomeWindow = driver.CurrentWindowHandle;


                // Case -1: Navigate to Investigation Master Case via Advanced Find and try to edit it.
                homePage.ClickAdvancedfindIMG();
                Thread.Sleep(400);
                driver = UICommon.SwitchToNewBrowserWithTitle(driver, HomeWindow, "Advanced Find");

                AdvancedFindPage advancedFindPage = new AdvancedFindPage(driver);
                Assert.IsFalse(advancedFindPage.VerifyLookForListItemPresent("Investigation Master Case"), "Investigation Master Case option present!!!");
                advancedFindPage.CloseWindow();

                driver = driver.SwitchTo().Window(HomeWindow);

                // Case -2: INVALID Navigate to Investigation Master Case via Home tiles and try to edit it.

                // Case -3 & 4: Navigate to Investigation Master Case via creating it through the home ribbon and try to edit it.
                homePage.HoverCRMRibbonTab();
                Assert.IsFalse(homePage.VerifyInvestigationsRibbonButtonPresent(), "Investigation Ribbon button Present!!!!");

                if (lastIteration == false)
                {
                    driver.Quit();
                    driver = null;
                    this.TestSetup();
                }
            }

        }

        [TestMethod]
        [TestProperty("TestcaseID", "3324")]
        public void ATC3324_CRMRecordoutcomeofnonlodgementofbondbreach()
        {

            string allegedoffenceId;
            string todayDate = DateTime.Now.ToString("d/MM/yyyy");
            string investigationID;

            //Login in as role
            User user = this.environment.GetUser(SecurityRole.Investigations);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsCasesRibbonButton();

            InvestigationCaseSearchPage investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationCaseSearchPage.ClickNewInvestigationCaseButton();

            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickSaveButton();
            investigationID = investigationCasePage.GetInvestigationCaseNumber();
            investigationCasePage.ClickSaveCloseButton();


            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickRightScrollRibbonButton();
            homePage.ClickAllegedOffencesButton();

            AllegendOffensesSearchPage allegedOffencesSearchPage = new AllegendOffensesSearchPage(driver);

            allegedOffencesSearchPage.ClickNewAllegedOffenceButton();

            AllegedOffencePage allegedOffencesPage = new AllegedOffencePage(driver);
            allegedOffencesPage.SetInvestigationCaseValue(investigationID);
            allegedOffencesPage.SetProvisionValue("RTRA 116(1)");
            allegedOffencesPage.SetOffenceDateValue("1/01/2015");
            allegedOffencesPage.SetBeliefFormedDateValue(todayDate);
            allegedOffencesPage.ClickSaveButton();
            allegedoffenceId = allegedOffencesPage.GetReferenceNumber();
            allegedOffencesPage.ClickSaveCloseButton();

            allegedOffencesSearchPage = new AllegendOffensesSearchPage(driver);
            allegedOffencesSearchPage.SetInvestigationSearchText(investigationID);
            Table table = new Table(allegedOffencesSearchPage.GetSearchResultTable());
            table.SelectTableRow("Status Reason", "Belief");

            allegedOffencesPage = new AllegedOffencePage(driver);
            allegedOffencesPage.SetOffenceDateValue("");
            StringAssert.Contains(allegedOffencesPage.GetOffenceDateValue(), "--");
            allegedOffencesPage.SetBeliefFormedDateValue("");
            StringAssert.Contains(allegedOffencesPage.GetStatusReason(), "Suspicion");

            allegedoffenceId = allegedOffencesPage.GetReferenceNumber();
            allegedOffencesPage.ClickSaveCloseButton();

            allegedOffencesSearchPage = new AllegendOffensesSearchPage(driver);
            allegedOffencesSearchPage.SetInvestigationSearchText(allegedoffenceId);
            table = new Table(allegedOffencesSearchPage.GetSearchResultTable());
            table.SelectContainsTableRow("Investigation Case", investigationID);

            allegedOffencesPage = new AllegedOffencePage(driver);
            StringAssert.Contains(allegedOffencesPage.GetStatusReason(), "Suspicion");
            StringAssert.Contains(allegedOffencesPage.GetOffenceDateValue(), "--");
            StringAssert.Contains(allegedOffencesPage.GetStatutoryLimitationValue(), "--");

            allegedOffencesPage.SetOffenceDateValue("01/01/2015");
            allegedOffencesPage.ClickSaveButton();

            StringAssert.Contains(allegedOffencesPage.GetStatutoryLimitationValue(), "1/01/2016");

            StringAssert.Contains(allegedOffencesPage.GetBefliefFormedDateValue(), "");

        }
        [TestMethod]
        [TestProperty("TestcaseID", "3316")]
        [TestProperty("TestType", "Regression")]
        public void ATC3316_CRMInvestigationVerifyRelatedActivitiesOwnerUpdation()
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

            string clientName = MyRange.Cells[MyRow, InvestigationSchema.GetColumnIndex(ColumnName.CLIENT_NAME)].Value;

            User user = this.environment.GetUser(SecurityRole.InvestigationOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            // Create Investigation Case
            HomePage homePage = new HomePage(driver);
            string HomeWindow = driver.CurrentWindowHandle;
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsCasesRibbonButton();

            InvestigationCaseSearchPage investigationsCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationsCaseSearchPage.ClickNewInvestigationCaseButton();
  
            // Create new case
            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            string BaseWindow = driver.CurrentWindowHandle;
            investigationCasePage.ClickSaveButton();
            String investigationID = investigationCasePage.GetInvestigationCaseNumber();

            //Email  
            investigationCasePage.ClickActivitiesAddButton();
            investigationCasePage.ClickAddActivity("Email");
            Thread.Sleep(1000);
            driver = investigationCasePage.SwitchNewBrowser(driver, BaseWindow);
            Thread.Sleep(2000);
            EmailPage emailPage = new EmailPage(driver);
            emailPage.ClickPageTitle();
            emailPage.SetToValueText(clientName);
            Thread.Sleep(500);
            emailPage.SetSubjectValueText("Test 3316 Email");
            emailPage.ClickSaveCloseIMG();
            Thread.Sleep(2000);
            driver = driver.SwitchTo().Window(BaseWindow);

            //Fax 
            investigationCasePage.ClickPageTitle();
            investigationCasePage.ClickActivitiesAddButton();
            investigationCasePage.ClickAddActivity("Fax");
            Thread.Sleep(1000);
            driver = investigationCasePage.SwitchNewBrowser(driver, BaseWindow);
            Thread.Sleep(2000);
            FaxPage faxPage = new FaxPage(driver);
            faxPage.ClickPageTitle();
            faxPage.SetSubjectValue("Test 3316 Fax");
            faxPage.ClickSaveCloseButton();
            Thread.Sleep(2000);
            driver = driver.SwitchTo().Window(BaseWindow);

            //Letter
            investigationCasePage.ClickPageTitle();
            investigationCasePage.ClickActivitiesAddButton();
            investigationCasePage.ClickAddActivity("Letter");
            Thread.Sleep(1000);
            driver = investigationCasePage.SwitchNewBrowser(driver, BaseWindow);
            Thread.Sleep(1000);
            LetterPage letterPage = new LetterPage(driver);
            letterPage.ClickPageTitle();
            letterPage.SetSubjectValue("Test 3316 Letter");
            letterPage.ClickSaveButton();
            letterPage.ClickSaveCloseButton();
            Thread.Sleep(2000);
            driver = driver.SwitchTo().Window(BaseWindow);

            //Phone Call  
            investigationCasePage.ClickPageTitle();
            investigationCasePage.ClickActivitiesAddButton();
            investigationCasePage.ClickAddActivity("Phone Call");
            Thread.Sleep(1000);
            driver = investigationCasePage.SwitchNewBrowser(driver, BaseWindow);
            Thread.Sleep(2000);
            PhoneCallPage phoneCall = new PhoneCallPage(driver);
            Thread.Sleep(100);
            phoneCall.ClickPageTitle();
            phoneCall.SetSelectSubjectValue("Bond balance enquiry");
            phoneCall.SetSubject("Test 3316 Phone Call");
            Thread.Sleep(500);
            phoneCall.SetRecipient(clientName);
            phoneCall.ClickSaveCloseButton();
            Thread.Sleep(2000);
            driver = driver.SwitchTo().Window(BaseWindow);

            // Client Management Activity  
            investigationCasePage.ClickPageTitle();
            investigationCasePage.ClickActivitiesAddButton();
            investigationCasePage.ClickAddActivity("Client Management Activity");
            Thread.Sleep(1000);
            driver = investigationCasePage.SwitchNewBrowser(driver, BaseWindow);
            Thread.Sleep(2000);
            ClientManagementActivityPage clientManagementActivityPage = new ClientManagementActivityPage(driver);
            clientManagementActivityPage.ClickPageTitle();
            clientManagementActivityPage.SetSubjectValue("Test 3316 Client Management Activity");
            clientManagementActivityPage.ClickSaveCloseButton();
            Thread.Sleep(2000);
            driver = driver.SwitchTo().Window(BaseWindow);

            // Front Counter Contact
            investigationCasePage.ClickPageTitle();
            investigationCasePage.ClickActivitiesAddButton();
            investigationCasePage.ClickAddActivity("Front Counter Contact");
            Thread.Sleep(2000);
            driver = investigationCasePage.SwitchNewBrowser(driver, BaseWindow);
            Thread.Sleep(1000);
            FrontCounterContactPage frontCounterContactPage = new FrontCounterContactPage(driver);
            frontCounterContactPage.SetSubjectValue("Test 3316 Front Counter Contact");
            frontCounterContactPage.ClickSaveCloseButton();
            Thread.Sleep(2000);
            driver = driver.SwitchTo().Window(BaseWindow);

            // Task             
            investigationCasePage.ClickPageTitle();
            investigationCasePage.ClickActivitiesAddButton();
            investigationCasePage.ClickAddTaskButton("Task");
            Thread.Sleep(1000);
            driver = investigationCasePage.SwitchNewBrowser(driver, BaseWindow);
            Thread.Sleep(500);
            TaskPage taskPage = new TaskPage(driver);
            taskPage.ClickPageTitle();
            taskPage.SetSelectSubjectValue("Bond existence");
            taskPage.ClickSaveCloseButton();
            Thread.Sleep(2000);
            driver = driver.SwitchTo().Window(BaseWindow);

            // Recurring Appointment    
            investigationCasePage.ClickPageTitle();
            investigationCasePage.ClickActivitiesAddButton();
            investigationCasePage.ClickAddActivity("Appointment");
            Thread.Sleep(1000);
            driver = investigationCasePage.SwitchNewBrowser(driver, BaseWindow);
            Thread.Sleep(2000);
            AppointmentPage appointmentPage = new AppointmentPage(driver);
            string AppWindow = driver.CurrentWindowHandle;
            appointmentPage.ClickPageTitle();
            appointmentPage.SetStartRange(DateTime.Today.AddDays(1).ToString("dd/MM/yyyy"));
            appointmentPage.ClickPageTitle();
            appointmentPage.ClickRecurrenceButton();
            appointmentPage.ClickSetButton();
            driver = driver.SwitchTo().Window(AppWindow);
            appointmentPage = new AppointmentPage(driver);
            appointmentPage.ClickPageTitle();
            appointmentPage.SetSubjectValue("Test 3316 Recurring Appointment");
            appointmentPage.ClickSaveCloseButton();
            Thread.Sleep(2000);
            driver = driver.SwitchTo().Window(BaseWindow);

            // Appointment 
            investigationCasePage.ClickPageTitle();
            investigationCasePage.ClickActivitiesAddButton();
            investigationCasePage.ClickAddActivity("Appointment");
            Thread.Sleep(1000);
            driver = investigationCasePage.SwitchNewBrowser(driver, BaseWindow);
            Thread.Sleep(2000);
            appointmentPage = new AppointmentPage(driver);
            appointmentPage.ClickPageTitle();
            appointmentPage.SetSubjectValue("Test 3316 Appointment");
            appointmentPage.ClickSaveCloseButton();
            Thread.Sleep(2000);
            driver = driver.SwitchTo().Window(BaseWindow);

            // Change the owner value
            investigationCasePage.ClickPageTitle();
            investigationCasePage.SetOwnerValue("IMSTestU04");
            investigationCasePage.ClickSaveCloseButton();
            investigationsCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationsCaseSearchPage.SetSearchRecord(investigationID);
            Table table = new Table(investigationsCaseSearchPage.GetSearchResultTable());
            table.ClickCellValue("Case Number", investigationID, "Case Number");

            driver = driver.SwitchTo().Window(BaseWindow);
            investigationCasePage.ClickPageTitle();
            investigationCasePage.ClickSeeRecordsAssociatedWithThisViewButton("Activities");
            Thread.Sleep(1000);

            investigationCasePage.SetActivitiesSearchText("Test 3316 Email");
            table = new Table(investigationCasePage.GetActivitiesAssociatedViewTable());
            StringAssert.Contains(table.GetCellContainsValue("Subject", "Test 3316 Email", "Owner"), "IMSTestU04");

            investigationCasePage.SetActivitiesSearchText("Test 3316 Fax");
            table = new Table(investigationCasePage.GetActivitiesAssociatedViewTable());
            StringAssert.Contains(table.GetCellContainsValue("Subject", "Test 3316 Fax", "Owner"), "IMSTestU04");

            investigationCasePage.SetActivitiesSearchText("Test 3316 Letter");
            table = new Table(investigationCasePage.GetActivitiesAssociatedViewTable());
            StringAssert.Contains(table.GetCellContainsValue("Subject", "Test 3316 Letter", "Owner"), "IMSTestU04");

            investigationCasePage.SetActivitiesSearchText("Test 3316 Phone Call");
            table = new Table(investigationCasePage.GetActivitiesAssociatedViewTable());
            StringAssert.Contains(table.GetCellContainsValue("Subject", "Test 3316 Phone Call", "Owner"), "IMSTestU04");

            investigationCasePage.SetActivitiesSearchText("Test 3316 Client Management Activity");
            table = new Table(investigationCasePage.GetActivitiesAssociatedViewTable());
            StringAssert.Contains(table.GetCellContainsValue("Subject", "Test 3316 Client Management Activity", "Owner"), "IMSTestU04");

            investigationCasePage.SetActivitiesSearchText("Test 3316 Front Counter Contact");
            table = new Table(investigationCasePage.GetActivitiesAssociatedViewTable());
            StringAssert.Contains(table.GetCellContainsValue("Subject", "Test 3316 Front Counter Contact", "Owner"), "IMSTestU04");

            investigationCasePage.SetActivitiesSearchText(investigationID + ": Bond existence");
            table = new Table(investigationCasePage.GetActivitiesAssociatedViewTable());
            StringAssert.Contains(table.GetCellContainsValue("Subject", investigationID + ": Bond existence", "Owner"), "IMSTestU04");

            investigationCasePage.SetActivitiesSearchText("Test 3316 Appointment");
            table = new Table(investigationCasePage.GetActivitiesAssociatedViewTable());
            StringAssert.Contains(table.GetCellContainsValue("Subject", "Test 3316 Appointment", "Owner"), "IMSTestU04");

            investigationCasePage.SetActivitiesSearchText("Test 3316 Recurring Appointment");
            table = new Table(investigationCasePage.GetActivitiesAssociatedViewTable());
            StringAssert.Contains(table.GetCellContainsValue("Subject", "Test 3316 Recurring Appointment", "Owner"), "IMSTestU04");

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion 
        }

        [TestMethod]
        [TestProperty("TestcaseID", "3309")]
        [TestProperty("TestType", "Regression")]
        public void ATC3309_CRMInvestigationVerifyOfficerCanRecordPhoneCalls()
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

            string clientName = MyRange.Cells[MyRow, InvestigationSchema.GetColumnIndex(ColumnName.CLIENT_NAME)].Value;

            User user = this.environment.GetUser(SecurityRole.InvestigationOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            // Create Investigation Case
            HomePage homePage = new HomePage(driver);
            string HomeWindow = driver.CurrentWindowHandle;
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsCasesRibbonButton();

            // Verify that Investigation Officer is able to record Phone call for Investigation Case
            InvestigationCaseSearchPage investigationsCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationsCaseSearchPage.ClickNewInvestigationCaseButton();

            // Create new case
            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            string BaseWindow = driver.CurrentWindowHandle;
            investigationCasePage.ClickSaveButton();
            String investigationID = investigationCasePage.GetInvestigationCaseNumber();

            //Phone Call  
            investigationCasePage.ClickPageTitle();
            investigationCasePage.ClickActivitiesAddButton();
            investigationCasePage.ClickAddActivity("Phone Call");
            driver = investigationCasePage.SwitchNewBrowserWithTitle(driver, BaseWindow, "Phone Call");    
      
            PhoneCallPage phoneCall = new PhoneCallPage(driver);
            phoneCall.ClickPageTitle();
            phoneCall.SetSelectSubjectValue("Bond balance enquiry");
            phoneCall.SetSubject("Test 3309 Phone Call");
            phoneCall.SetRecipient(clientName);
            phoneCall.ClickSaveCloseButton();
            driver = driver.SwitchTo().Window(BaseWindow);
            investigationCasePage.CheckForErrors();
            investigationCasePage.ClickPageTitle();
            Table table = new Table(investigationCasePage.GetActivitiesSearchResultTable());
            StringAssert.Contains(table.GetCellContainsValue("Subject", "Test 3309 Phone Call", "Owner"), "IMSTestU03");

            table.ClickCellValue("Subject", "Test 3309 Phone Call", "Subject");
            phoneCall = new PhoneCallPage(driver);
            phoneCall.SwitchFrame();
            StringAssert.Contains(phoneCall.GetSelectSubjectValue(), "Bond balance enquiry");
            StringAssert.Contains(phoneCall.GetSubjectValue(), "Test 3309 Phone Call");
            StringAssert.Contains(phoneCall.GetRecipientValue(), clientName);
            StringAssert.Contains(phoneCall.GetSenderValue(), user.Id);

            // Verify that Investigation Officer is able to record Phone call for Client
            driver = driver.SwitchTo().Window(HomeWindow);
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();

            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.ClickNewClientButton();

            ClientPage clientPage = new ClientPage(driver);
            BaseWindow = driver.CurrentWindowHandle;
            clientPage.ClickPageTitle();
            
            clientPage.PopulateNewClient("Client Phone Record");
            clientPage.ClickSaveButton();

            clientPage.ClickActivitiesAddButton();
            clientPage.ClickAddActivity("Phone Call");
            Thread.Sleep(3000);
            driver = clientPage.SwitchNewBrowser(driver, BaseWindow, "Phone Call");
  
            phoneCall = new PhoneCallPage(driver);
   
            phoneCall.ClickPageTitle();
            phoneCall.SetSelectSubjectValue("Bond balance enquiry");
            phoneCall.SetSubject("Test 3309 Client Phone Call");
 
            phoneCall.SetRecipient(clientName);
            phoneCall.ClickSaveCloseButton();
            driver = driver.SwitchTo().Window(BaseWindow);
            phoneCall.CheckForErrors();
            clientPage.ClickPageTitle();
            table = new Table(clientPage.GetActivitiesTable());
            StringAssert.Contains(table.GetCellContainsValue("Subject", "Test 3309 Client Phone Call", "Owner"), "IMSTestU03");

            table.ClickCellValue("Subject", "Test 3309 Client Phone Call", "Subject");
            phoneCall = new PhoneCallPage(driver);
            StringAssert.Contains(phoneCall.GetSelectSubjectValue(), "Bond balance enquiry");
            StringAssert.Contains(phoneCall.GetSubjectValue(), "Test 3309 Client Phone Call");
            StringAssert.Contains(phoneCall.GetRecipientValue(), clientName);
            StringAssert.Contains(phoneCall.GetSenderValue(), user.Id);

            // Verify that Investigation Officer is able to record Phone call for General Case
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsGeneralCasesRibbonButton();

            InvestigationGeneralCaseSearchPage invGeneralCaseSearchPage = new InvestigationGeneralCaseSearchPage(driver);
            invGeneralCaseSearchPage.ClickNewGeneralCaseButton();

            InvestigationGeneralCasePage investigationGeneralCasePage = new InvestigationGeneralCasePage(driver);
            BaseWindow = driver.CurrentWindowHandle;
            investigationGeneralCasePage.ClickPageTitle();
            investigationGeneralCasePage.SetTitle("New General Case");
            investigationGeneralCasePage.SetClientName(clientName);
            investigationGeneralCasePage.SetType("Complaint");
            investigationGeneralCasePage.ClickSaveButton();

            investigationGeneralCasePage.ClickActivitiesAddButton();
            investigationGeneralCasePage.ClickAddActivity("Phone Call");

  
            driver = investigationGeneralCasePage.SwitchNewBrowserWithTitle(driver, BaseWindow, "Phone Call");

            phoneCall = new PhoneCallPage(driver);

            phoneCall.ClickPageTitle();
            phoneCall.SetSelectSubjectValue("Bond balance enquiry");
            phoneCall.SetSubject("Test 3309 General Case Phone Call");
         
            phoneCall.SetRecipient(clientName);
            phoneCall.ClickSaveCloseButton();
            driver = driver.SwitchTo().Window(BaseWindow);
            phoneCall.CheckForErrors();
            investigationGeneralCasePage.ClickPageTitle();

             table = new Table(investigationGeneralCasePage.GetActivitiesTable());
            StringAssert.Contains(table.GetCellContainsValue("Subject", "Test 3309 General Case Phone Call", "Owner"), "IMSTestU03");

            table.ClickCellValue("Subject", "Test 3309 General Case Phone Call", "Subject");
            phoneCall = new PhoneCallPage(driver);
            StringAssert.Contains(phoneCall.GetSelectSubjectValue(), "Bond balance enquiry");
            StringAssert.Contains(phoneCall.GetSubjectValue(), "Test 3309 General Case Phone Call");
            StringAssert.Contains(phoneCall.GetRecipientValue(), clientName);
            StringAssert.Contains(phoneCall.GetSenderValue(), user.Id);

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion        
        }

        [TestMethod]
        [TestProperty("TestcaseID", "3312")]
        [TestProperty("TestType", "Regression")]
        public void ATC3312_CRMInvestigationCaseCloseAndReopen()
        {

            //There is only 1
            string InvestigationWindowHandle; // for driver
            string ClientWindowHandle; // for ClientDriver

            string ClientName = "CLIENT CASE" + UICommon.GetRandomString(3);
           

            User user = this.environment.GetUser(SecurityRole.InvestigationOfficer);
            new LoginDialog().Login(user.Id, user.Password);
            
            // Create Investigation Case
            HomePage homePageInvestigation = new HomePage(driver);
            InvestigationWindowHandle = driver.CurrentWindowHandle;
            homePageInvestigation.HoverCRMRibbonTab();
            homePageInvestigation.ClickInvestigationsRibbonButton();
            homePageInvestigation.HoverInvestigationsRibbonTab();
            homePageInvestigation.ClickInvestigationsCasesRibbonButton();
            
            // Verify that Investigation Officer is able to record Phone call for Investigation Case
            InvestigationCaseSearchPage investigationsCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationsCaseSearchPage.ClickNewInvestigationCaseButton();

            // Create new Investigation case
            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickSaveButton();
            String investigationID = investigationCasePage.GetInvestigationCaseNumber();


            IWebDriver ClientDriver = null;
            // Spawn a new window and open a Client record
            if (Properties.Settings.Default.BROWSER == BrowserType.Ie)
            {
                ClientDriver = new BrowserContext().WebDriver;
                this.environment = TestEnvironment.GetTestEnvironment();
            }
            
            ClientDriver.Navigate().GoToUrl(this.environment.Url);
            new LoginDialog().Login(user.Id, user.Password);

            // Create new Client Profile
            HomePage homePageClient = new HomePage(ClientDriver);
            ClientWindowHandle = ClientDriver.CurrentWindowHandle;
            homePageClient.HoverCRMRibbonTab();
            homePageClient.ClickClientServicesRibbonButton();
            homePageClient.HoverClientServicesRibbonTab();
            homePageClient.ClickClientsRibbonButton();

            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(ClientDriver);
            clientsSearchPage.ClickNewClientButton();

            ClientPage clientPage = new ClientPage(ClientDriver);
            clientPage.ClickPageTitle();
            clientPage.PopulateNewClient(ClientName);
            clientPage.ClickSaveCloseButton();         
            Thread.Sleep(3000);
            clientsSearchPage = new ClientsSearchPage(ClientDriver);

            // ADD the Client from the second open window as a Case Party to the Investigation Case
            driver.SwitchTo().Window(InvestigationWindowHandle);
            investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickPageTitle();
            investigationCasePage.ClickAddCasePartyRecordButton();

            investigationCasePage.SwitchNewBrowserWithTitle(driver, InvestigationWindowHandle, "Case Party");

            CasePartyPage casePartyPage = new CasePartyPage(driver);
            casePartyPage.ClickPageTitle();
            casePartyPage.SetClientName(ClientName);
            casePartyPage.ClickSaveCloseButton();
            Thread.Sleep(3000);

            driver = driver.SwitchTo().Window(InvestigationWindowHandle);
            investigationCasePage.ClickSaveCloseButton();

            // Verify new Case Management Activity is shown in the Investigation Case which documents that the Case Party was added and by whom.
            investigationsCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationsCaseSearchPage.SetSearchRecord(investigationID);
            Table table = new Table(investigationsCaseSearchPage.GetSearchResultTable());
            table.ClickCellValue("Case Number", investigationID, "Case Number");

            investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickPageTitle();
            table = new Table(investigationCasePage.GetActivitiesSearchResultTable());
            StringAssert.Contains(table.GetCellContainsValue("Subject", "Case party " + ClientName + " added to case " + investigationID, "Activity Status"), "Completed");
            StringAssert.Contains(table.GetCellContainsValue("Subject", "Case party " + ClientName + " added to case " + investigationID, "Modified By"), "SRCRM5-TEAdmin Last Name");
            

            //Client window
            ClientDriver = ClientDriver.SwitchTo().Window(ClientWindowHandle);
            clientsSearchPage = new ClientsSearchPage(ClientDriver);
            clientsSearchPage.SetClientSearchText(ClientName);

            table = new Table(clientsSearchPage.GetSearchResultTable());
            table.ClickCellValue("Full Name", ClientName, "Full Name");

            clientPage = new ClientPage(ClientDriver);
            clientPage.ClickSeeRecordsAssociatedWithThisViewButton("Activities");
          
            clientPage.SwitchToFrame();
            
            clientPage.SetPageFilterList("All Activities", ClientDriver);
            clientPage.SetFilterOnList("All", ClientDriver);
            table = new Table(clientPage.GetActivitiesAssociatedViewTable(ClientDriver));
            StringAssert.Contains(table.GetCellValue("Subject", "Case party " + ClientName + " added to case " + investigationID, "Activity Status"), "Completed");
            StringAssert.Contains(table.GetCellValue("Subject", "Case party " + ClientName + " added to case " + investigationID, "Modified By"), "SRCRM5-TEAdmin Last Name");
            

            // Investigation Case window, SELECT Start Dialog and close the Investigation case
            driver.SwitchTo().Window(InvestigationWindowHandle);
            investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickPageTitle();
            investigationCasePage.ClickStartDialogButton();

            table = new Table(investigationCasePage.GetProcessSearchResultTable());
            table.ClickCell("Process Name", "INV: Close investigation case", "Created On");
            investigationCasePage.ClickDialogAddButton();
            
            driver = UICommon.SwitchToNewBrowserWithTitle(driver, InvestigationWindowHandle, "INV:");

            INVPage iNVPage = new INVPage(driver);
            iNVPage.ClickNextButton();
            iNVPage.ClickNextButton();
            iNVPage.ClickNextButton();
            iNVPage.ClickFinishButton();
            Thread.Sleep(3000);

            driver = driver.SwitchTo().Window(InvestigationWindowHandle);
            investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickSaveCloseButton();

            // Investigation Case is shown as Inactive. A new Case Management Activity is shown in the Investigation Case which documents that the case was closed and by whom.
            investigationsCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationsCaseSearchPage.SetPageFilterList("All Investigation Cases");

            table = new Table(investigationsCaseSearchPage.GetHeaderSearchResultTable());
            table.ClickTableColumnHeader("Case Number");

            table = new Table(investigationsCaseSearchPage.GetSearchResultTable());
            table.ClickCellValue("Case Number", investigationID, "Case Number");

            investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickPageTitle();
            StringAssert.Contains(investigationCasePage.GetStatus(), "Closed");
            table = new Table(investigationCasePage.GetActivitiesSearchResultTable());
            StringAssert.Contains(table.GetCellContainsValue("Subject", investigationID + " case closed: Act applies-Insufficient evidence", "Activity Status"), "Completed");
            StringAssert.Contains(table.GetCellContainsValue("Subject", investigationID + " case closed: Act applies-Insufficient evidence", "Modified By"), "IMSTestU03");

            
            // A new Case Management Activity is shown in the Client which documents that the case was closed and by whom.
            ClientDriver.SwitchTo().Window(ClientWindowHandle);
            homePageClient.HoverCRMRibbonTab();
            homePageClient.ClickClientServicesRibbonButton();
            homePageClient.HoverClientServicesRibbonTab();
            homePageClient.ClickClientsRibbonButton();

            clientsSearchPage = new ClientsSearchPage(ClientDriver);
            clientsSearchPage.SetClientSearchText(ClientName);

            table = new Table(clientsSearchPage.GetSearchResultTable());
            table.ClickCellValue("Full Name", ClientName, "Full Name");

            clientPage = new ClientPage(ClientDriver);
            clientPage.ClickSeeRecordsAssociatedWithThisViewButton("Activities");

            clientPage.SwitchToFrame();

            Thread.Sleep(1000);
            clientPage.SetPageFilterList("All Activities", ClientDriver);
            clientPage.SetFilterOnList("All", ClientDriver);
            table = new Table(clientPage.GetActivitiesAssociatedViewTable(ClientDriver));
            Console.WriteLine(table.GetRowCount());
            StringAssert.Contains(table.GetCellValue("Subject", investigationID + " case closed: Act applies-Insufficient evidence", "Activity Status"), "Completed");
            StringAssert.Contains(table.GetCellValue("Subject", investigationID + " case closed: Act applies-Insufficient evidence", "Modified By"), "IMSTestU03");

            // In the Investigation Case, SELECT Start Dialog and Re-Open the case
            driver.SwitchTo().Window(InvestigationWindowHandle);
            investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickPageTitle();
            investigationCasePage.ClickStartDialogButton();

            table = new Table(investigationCasePage.GetProcessSearchResultTable());
            table.ClickCell("Process Name", "INV: Re-open investigation case", "Created On");
            investigationCasePage.ClickDialogAddButton();

            driver = UICommon.SwitchToNewBrowserWithTitle(driver, InvestigationWindowHandle, "INV:");

            iNVPage = new INVPage(driver);
            iNVPage.ClickNextButton();
            iNVPage.ClickNextButton();
            iNVPage.ClickFinishButton();
            
            driver = driver.SwitchTo().Window(InvestigationWindowHandle);
            investigationCasePage = new InvestigationCasePage(driver);
            homePageInvestigation.HoverInvestigationsRibbonTab();
            homePageInvestigation.ClickInvestigationsCasesRibbonButton();
            investigationsCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationsCaseSearchPage.SetSearchRecord(investigationID);    

            table = new Table(investigationsCaseSearchPage.GetSearchResultTable());
            table.ClickCellValue("Case Number", investigationID, "Case Number");

            // Investigation Case is shown as Active. A new Case Management Activity is shown in the Investigation Case which documents that the case was closed and by whom.
            investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickPageTitle();
            StringAssert.Contains(investigationCasePage.GetStatus(), "Case finalisation");
            table = new Table(investigationCasePage.GetActivitiesSearchResultTable());
            StringAssert.Contains(table.GetCellContainsValue("Subject", investigationID + " case re-opened: Case finalisation - Case closed discussion", "Activity Status"), "Completed");
            StringAssert.Contains(table.GetCellContainsValue("Subject", investigationID + " case re-opened: Case finalisation - Case closed discussion", "Modified By"), "IMSTestU03");

            // A new Case Management Activity is shown in the Client which documents that the case was reopened and by whom.
            ClientDriver.SwitchTo().Window(ClientWindowHandle);
            homePageClient.HoverClientServicesRibbonTab();
            homePageClient.ClickClientsRibbonButton();
            clientsSearchPage = new ClientsSearchPage(ClientDriver);
            clientsSearchPage.SetClientSearchText(ClientName);

            table = new Table(clientsSearchPage.GetSearchResultTable());
            table.ClickCellValue("Full Name", ClientName, "Full Name");

            clientPage = new ClientPage(ClientDriver);
            clientPage.ClickSeeRecordsAssociatedWithThisViewButton("Activities");
            
            clientPage.SwitchToFrame();
            Thread.Sleep(1000);
            clientPage.SetPageFilterList("All Activities", ClientDriver);
            clientPage.SetFilterOnList("All", ClientDriver);
            table = new Table(clientPage.GetActivitiesAssociatedViewTable(ClientDriver));
            StringAssert.Contains(table.GetCellValue("Subject", investigationID + " case re-opened: Case finalisation - Case closed discussion", "Activity Status"), "Completed");
            StringAssert.Contains(table.GetCellValue("Subject", investigationID + " case re-opened: Case finalisation - Case closed discussion", "Modified By"), "IMSTestU03");

        }

        [TestMethod]
        [TestProperty("TestcaseID", "3313")]
        [TestProperty("TestType", "Regression")]
        public void ATC3313_CRMInvestigationAddAndDeactivateCaseParty()
        {

            User user = this.environment.GetUser(SecurityRole.InvestigationOfficer);
            new LoginDialog().Login(user.Id, user.Password);
            string ClientName = "CLIENT CASE" + UICommon.GetRandomString(3);

            // Create Investigation Case
            HomePage homePage = new HomePage(driver);
            string HomeWindowInv = driver.CurrentWindowHandle;
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsCasesRibbonButton();

            // Verify that Investigation Officer is able to record Phone call for Investigation Case
            InvestigationCaseSearchPage investigationsCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationsCaseSearchPage.ClickNewInvestigationCaseButton();

            // Create new Investigation case
            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickSaveButton();
            String investigationID = investigationCasePage.GetInvestigationCaseNumber();

            string InvWindow = driver.CurrentWindowHandle;

            IWebDriver ClientDriver = null;
            // Spawn a new window and open a Client record
            if (Properties.Settings.Default.BROWSER == BrowserType.Ie)
            {
                ClientDriver = new BrowserContext().WebDriver;
                this.environment = TestEnvironment.GetTestEnvironment();
            }

            ClientDriver.Navigate().GoToUrl(this.environment.Url);
            new LoginDialog().Login(user.Id, user.Password);

            // Create new Client Profile
            homePage = new HomePage(ClientDriver);
            string HomeWindowClient = ClientDriver.CurrentWindowHandle;
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();


            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(ClientDriver);
            Trace.Listeners.Add(new TextWriterTraceListener("TextWriterOutput.log", "myListener"));
            Trace.TraceInformation("Add Client Case Party start:" + DateTime.Now.ToString("ddMMyyyyhhmmssffff"));
            Trace.Flush();
            
            clientsSearchPage.ClickNewClientButton();

            ClientPage clientPage = new ClientPage(ClientDriver);
            clientPage.ClickPageTitle();
            clientPage.PopulateNewClient(ClientName);
            clientPage.ClickSaveCloseButton();
            string ClientWindow = ClientDriver.CurrentWindowHandle;
                
            // ADD the Client from the second open window as a Case Party to the Investigation Case
            driver.SwitchTo().Window(InvWindow);
            investigationCasePage = new InvestigationCasePage(driver);
            string BaseWindow = driver.CurrentWindowHandle;
            investigationCasePage.ClickAddCasePartyRecordButton();

            investigationCasePage.SwitchNewBrowserWithTitle(driver, BaseWindow, "Case Party");

                  
            CasePartyPage casePartyPage = new CasePartyPage(driver);
            casePartyPage.ClickPageTitle();
            casePartyPage.ClickPartyType();
            Assert.IsTrue(casePartyPage.VerifyCasePartyTypeExists("Bond Contributor"), "Bond Contributor option not seen!!!");
            Assert.IsTrue(casePartyPage.VerifyCasePartyTypeExists("Tenant"), "Tenant option not seen!!!");
            Assert.IsTrue(casePartyPage.VerifyCasePartyTypeExists("Resident"), "Resident option not seen!!!");
            Assert.IsTrue(casePartyPage.VerifyCasePartyTypeExists("Agent"), "Agent option not seen!!!");
            Assert.IsTrue(casePartyPage.VerifyCasePartyTypeExists("Lessor"), "Lessor option not seen!!!");
            Assert.IsTrue(casePartyPage.VerifyCasePartyTypeExists("Rooming Accommodation Provider"), "Rooming Accommodation Provider option not seen!!!");
            Assert.IsTrue(casePartyPage.VerifyCasePartyTypeExists("Student Accommodation Provider"), "Student Accommodation Provider option not seen!!!");
            Assert.IsTrue(casePartyPage.VerifyCasePartyTypeExists("Moveable Dwelling Provider"), "Moveable Dwelling Provider option not seen!!!");
            Assert.IsTrue(casePartyPage.VerifyCasePartyTypeExists("Community Housing Organisation"), "Community Housing Organisation option not seen!!!");
            Assert.IsTrue(casePartyPage.VerifyCasePartyTypeExists("Owner"), "Owner option not seen!!!");
            Trace.Listeners.Add(new TextWriterTraceListener("TextWriterOutput.log", "myListener"));
            Trace.TraceInformation("Add Case Party end:" + DateTime.Now.ToString("ddMMyyyyhhmmssffff"));
            Trace.Flush();
            casePartyPage.ClickPageTitle();
            casePartyPage.SetClientName(ClientName);
            casePartyPage.ClickSaveCloseButton();

            driver = driver.SwitchTo().Window(BaseWindow);
            investigationCasePage.CheckForErrors();
            investigationCasePage.ClickSaveCloseButton();
            
            // Verify new Case Management Activity is shown in the Investigation Case which documents that the Case Party was added and by whom.

            investigationsCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationsCaseSearchPage.SetSearchRecord(investigationID);

            Table table = new Table(investigationsCaseSearchPage.GetSearchResultTable());
            table.ClickCellValue("Case Number", investigationID, "Case Number");
            driver = driver.SwitchTo().Window(BaseWindow);
            investigationCasePage.ClickPageTitle();
            table = new Table(investigationCasePage.GetActivitiesSearchResultTable());
            StringAssert.Contains(table.GetCellContainsValue("Subject", "Case party " + ClientName + " added to case " + investigationID, "Activity Status"), "Completed");
            InvWindow = driver.CurrentWindowHandle;

            //Client window
            ClientDriver = ClientDriver.SwitchTo().Window(ClientWindow);
            clientsSearchPage = new ClientsSearchPage(ClientDriver);
            clientsSearchPage.SetClientSearchText(ClientName);

            table = new Table(clientsSearchPage.GetSearchResultTable());
            table.ClickCellValue("Full Name", ClientName, "Full Name");

            clientPage = new ClientPage(ClientDriver);
            Thread.Sleep(3000);
            clientPage.ClickSeeRecordsAssociatedWithThisViewButton("Activities");
            clientPage.SwitchToFrame();
            clientPage.SetPageFilterList("All Activities", ClientDriver);
            clientPage.SetFilterOnList("All", ClientDriver);
            table = new Table(clientPage.GetActivitiesAssociatedViewTable(ClientDriver));
            StringAssert.Contains(table.GetCellValue("Subject", "Case party " + ClientName + " added to case " + investigationID, "Activity Status"), "Completed");
            ClientWindow = ClientDriver.CurrentWindowHandle;

            // DEACTIVATE the Case Party by opening the Case Party Associated View and running the Dialog.
            driver = driver.SwitchTo().Window(InvWindow);
            investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickPageTitle();
            table = new Table(investigationCasePage.GetCasePartyTable());
            table.ClickCellContainsValueEnterRow("Client", ClientName, "Party Type");

            casePartyPage = new CasePartyPage(driver);
            casePartyPage.ClickPageTitle();
            casePartyPage.ClickSaveButton();
            casePartyPage.ClickStartDialog();
            table = new Table(casePartyPage.GetProcessSearchResultTable());
            table.ClickCell("Process Name", "INV: Deactivate case party", "Created On");
            BaseWindow = driver.CurrentWindowHandle;
            casePartyPage.ClickDialogAddButton();

            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "INV:");

            INVPage INV = new INVPage(driver);
            INV.ClickNextButton();
            INV.ClickFinishButton();

            driver = driver.SwitchTo().Window(BaseWindow);
            casePartyPage.ClickSaveCloseButton();

            investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickPageTitle();

            table = new Table(investigationCasePage.GetActivitiesSearchResultTable());
            StringAssert.Contains(table.GetCellContainsValue("Subject", "Case party " + ClientName + " removed from case " + investigationID, "Activity Status"), "Completed");

            // A new Case Management Activity is shown in the Client which documents that the Case Party was removed and by whom.

            ClientDriver = ClientDriver.SwitchTo().Window(ClientWindow);
            ClientDriver = ClientDriver.SwitchTo().Window(HomeWindowClient);
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();

            clientsSearchPage = new ClientsSearchPage(ClientDriver);
            clientsSearchPage.SetClientSearchText(ClientName);

            table = new Table(clientsSearchPage.GetSearchResultTable());
            table.ClickCellValue("Full Name", ClientName, "Full Name");

            clientPage = new ClientPage(ClientDriver);
            clientPage.ClickPageTitle();
            clientPage.ClickSeeRecordsAssociatedWithThisViewButton("Activities");
            clientPage.SwitchToFrame();
            clientPage.SetPageFilterList("All Activities", ClientDriver);
            clientPage.SetFilterOnList("All", ClientDriver);
            table = new Table(clientPage.GetActivitiesAssociatedViewTable(ClientDriver));
            StringAssert.Contains(table.GetCellValue("Subject", "Case party " + ClientName + " removed from case " + investigationID, "Activity Status"), "Completed");
        }
    }
}

