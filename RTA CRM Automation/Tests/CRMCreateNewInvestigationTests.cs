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
using System.IO;
using System.Reflection;
using RTA.Automation.CRM.DataSource;
using System.Windows.Forms;
using System.Linq;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using AutoItX3Lib;
using RTA.Automation.CRM.Common;
using RTA.Automation.CRM.Pages.Investigations;
using System.Diagnostics;

namespace RTA.Automation.CRM.Tests
{
    [TestClass]
    public class CRMCreateNewInvestigationTests : BaseTest
    {

        

        [TestInitialize]
        public void DataSourceSetup()
        {
            //open excel
            MyApp = new Excel.Application();
            MyApp.Visible = false;

        }

        [TestMethod]
        [TestProperty("TestcaseID", "3368")]
        public void ATC3368_CRMCheckInvestigationCaseID()
        {
            User user = this.environment.GetUser(SecurityRole.Investigations);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsCasesRibbonButton();

            InvestigationCaseSearchPage investigationsCaseSearchPage = new InvestigationCaseSearchPage(driver);

            investigationsCaseSearchPage.ClickNewInvestigationCaseButton();

            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);

            investigationCasePage.ClickSaveButton();
            string caseId1 = investigationCasePage.GetInvestigationCaseNumber();
            int caseNum1 = int.Parse(caseId1.Substring(2));

            investigationCasePage.ClickNewInvestigationCaseButton();

            //InvestigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickSaveButton();

            string caseId2 = investigationCasePage.GetInvestigationCaseNumber();
            int caseNum2 = int.Parse(caseId2.Substring(2));

            if (caseId1.StartsWith("IN") && (caseId2.StartsWith("IN")))  //checks if Investigation case starts with "IN"
            {
                Assert.AreEqual(caseNum1, caseNum2 - 1);      //if caseNum2 is a single increment from caseNum1 then TRUE
            }
            else
            {
                Assert.Fail();
            }
        }


        [TestMethod]
        [TestProperty("TestcaseID", "6827")]
        public void ATC6827_CRMCheckInvestigatorField()
        {

            User user = this.environment.GetUser(SecurityRole.Investigations);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsCasesRibbonButton();

            InvestigationCaseSearchPage investigationsCaseSearchPage = new InvestigationCaseSearchPage(driver);

            investigationsCaseSearchPage.ClickNewInvestigationCaseButton();

            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            string user02 = "IMSTestU02";
            Assert.IsTrue(investigationCasePage.GetInvestigatorSearchElementText(user02));
            investigationCasePage.ClickSaveButton();
            string investigationCase = investigationCasePage.GetInvestigationCaseNumber();
            investigationCasePage.ClickSaveCloseButton();
            investigationsCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationsCaseSearchPage.SetInvestigationSearchText(investigationCase);

            Table table = new Table(investigationsCaseSearchPage.GetSearchResultTable());
            StringAssert.Contains(table.GetCellValue("Case Number", investigationCase, "Investigator"), user02);


        }

        [TestMethod]
        [TestProperty("TestcaseID", "9361")]
        public void ATC9361_CRMRegisterDocumentIntoInvestigationCase()
        {
            User user = this.environment.GetUser(SecurityRole.Investigations);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsCasesRibbonButton();

            InvestigationCaseSearchPage investigationsCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationsCaseSearchPage.ClickNewInvestigationCaseButton();
            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            investigationCasePage.ClickSaveButton();
            var investigationCase = investigationCasePage.GetInvestigationCaseNumber();

            homePage.HoverClientRibbonTab(investigationCase.ToUpperInvariant());
            Thread.Sleep(500);
            driver.FindElement(By.XPath("//a[@title='Documents' and @id='Node_navDocument']")).Click();

            var directoryName = UICommon.GetAlertMessage(driver).Split(new[]{"/"}, StringSplitOptions.RemoveEmptyEntries).Last();

            var baseWindow = driver.CurrentWindowHandle; //Records the current window handle
            var gridFrame = driver.SwitchTo().DefaultContent().SwitchTo().Frame("contentIFrame1").SwitchTo().Frame("areaDocumentFrame").SwitchTo().Frame("gridIframe");
            var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("mnuNewButton")));
            gridFrame.FindElement(By.Id("mnuNewButton")).Click();
            Thread.Sleep(1000);
            gridFrame.FindElement(By.XPath("//span[contains(text(), 'Investigation Case Document')]")).Click();
            Thread.Sleep(3000);
            SwitchWindow(baseWindow);

            var a = new Actions(driver);
            var fileUpload = driver.SwitchTo().Frame("spPageFrame").FindElement(By.Id("ctl00_PlaceHolderMain_UploadDocumentSection_ctl05_InputFile"));
            a.DoubleClick(fileUpload).Build().Perform();
            Thread.Sleep(2000);

            var myFile = Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal), "inv_case_doc.txt");
            if (File.Exists(myFile)) File.Delete(myFile);
            var sw = new StreamWriter(myFile);
            sw.WriteLine("Hello world!");
            sw.Flush();
            sw.Close();

            var ai = new AutoItX3();
            var winTitle = "Choose File to Upload";
            int found = ai.WinWait(winTitle, "", Properties.Settings.Default.SHORT_WAIT_SECONDS);
            
            if (found == 1)
            {
                ai.WinActivate(winTitle);
                ai.Send(myFile);
                ai.Send("{ENTER}");
                ai.WinWaitClose(winTitle);
            }
            else
            {
                throw new Exception("Unable to locate open file dialog");
            }

            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("ctl00_PlaceHolderMain_ctl03_RptControls_btnOK")));
            driver.FindElement(By.Id("ctl00_PlaceHolderMain_ctl03_RptControls_btnOK")).Click();

            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Ribbon.DocLibListForm.Edit.Commit.Publish-Large")));
            driver.FindElement(By.Id("Ribbon.DocLibListForm.Edit.Commit.Publish-Large")).Click();

            gridFrame = driver.SwitchTo().DefaultContent().SwitchTo().Frame("contentIFrame1").SwitchTo().Frame("areaDocumentFrame").SwitchTo().Frame("gridIframe");
            var table = gridFrame.FindElement(By.XPath("//div[@id='divDataArea']/div/table"));
            
            // verify document is in the table
            if (!table.IsTextInTable("inv_case_doc.txt"))
                throw new Exception("Unable to locate file in CRM");

            gridFrame.FindElement(By.Id("openSharepointButton")).Click();
            Thread.Sleep(5000);
            SwitchWindow(baseWindow);

            var spTable = driver.FindElement(By.Id("onetidDoclibViewTbl0"));

            // verify document is in the table
            if (!spTable.IsTextInTable("inv_case_doc.txt"))
                throw new Exception("Unable to locate file in Sharepoint");

            driver = driver.SwitchTo().Window(baseWindow);
        }

        private static void SwitchWindow(string baseWindow)
        {
            string NewWindow = ""; //prepares for the new window handle

            ReadOnlyCollection<string> handles = null;
            for (int i = 1; i < 60; i++)
            {
                if (driver.WindowHandles.Count == 1)
                { Thread.Sleep(1000); }
                else { break; }
            }
            handles = driver.WindowHandles;
            foreach (string handle in handles)
            {
                var Handles = handle;
                if (baseWindow != handle)
                {
                    NewWindow = handle;
                    driver = driver.SwitchTo().Window(NewWindow);
                    break;
                }
            }
        }

        [TestMethod]
        [TestProperty("TestcaseID", "4415")]
        public void ATC4415_CRMAdditionalStatusforAllegedoffence()
        {
            User user = this.environment.GetUser(SecurityRole.Investigations);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsCasesRibbonButton();

            InvestigationCaseSearchPage investigationsCaseSearchPage = new InvestigationCaseSearchPage(driver);

            investigationsCaseSearchPage.ClickNewInvestigationCaseButton();

            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);

            investigationCasePage.ClickSaveButton();
            string caseId = investigationCasePage.GetInvestigationCaseNumber();

            investigationCasePage.ClickAllegedOffencesTab();
            string BaseWindow = driver.CurrentWindowHandle;
            investigationCasePage.ClickAllegedOffencesAddButton();
            Thread.Sleep(3000);
            driver = investigationCasePage.SwitchNewBrowserWithTitle(driver, BaseWindow, "Alleged Offence:");

            AllegedOffencePage allegedOffencePage = new AllegedOffencePage(driver);
            allegedOffencePage.ClickPageTitle();
            allegedOffencePage.SetProvisionValue("RTRA 101(1): Rent in advance - maximum amounts required");
            allegedOffencePage.ClickPageTitle();
            allegedOffencePage.SetOutcomeValue("Act applies-Statutory limitation expired");
            allegedOffencePage.ClickSaveButton();
            StringAssert.Contains(allegedOffencePage.GetReferenceNumber(), "IN");
            allegedOffencePage.ClickSaveCloseButton();
            driver = driver.SwitchTo().Window(BaseWindow);

        }



        [TestMethod]
        [TestProperty("TestcaseID", "6710")]
        public void ATC6710_CRMMasterCaseInvestigationCaseOptionalTasks()
        {
            string INVNumber;
            User user = this.environment.GetUser(SecurityRole.Investigations);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsMasterCasesRibbonButton();

            InvestigationMasterCaseSearchPage investigationMasterCasesSearchPage = new InvestigationMasterCaseSearchPage(driver);
            investigationMasterCasesSearchPage.ClickNewButton();

            InvestigationMasterCasePage investigationMasterCasePage = new InvestigationMasterCasePage(driver);
            investigationMasterCasePage.ClickSaveButton();
            INVNumber = investigationMasterCasePage.GetInvestigationMasterCaseNumber();
            investigationMasterCasePage.ClickSaveCloseButton();
            investigationMasterCasesSearchPage = new InvestigationMasterCaseSearchPage(driver);
            investigationMasterCasesSearchPage.SetInvestigationSearchText(INVNumber);
            Table table = new Table(investigationMasterCasesSearchPage.GetSearchResultTable());
            table.ClickCellContainsValueEnterRow("Master Case ID", INVNumber, "Master Case ID");
            investigationMasterCasePage = new InvestigationMasterCasePage(driver);
            investigationMasterCasePage.ClickPageTitle();
            investigationMasterCasePage.ClickActivitiesAddButton();
            investigationMasterCasePage.ClickCRMToolbar();
            string BaseWindow = driver.CurrentWindowHandle; //Records the current window handle
            investigationMasterCasePage.ClickAddTaskButton("Task");
            Thread.Sleep(3000);
            //Enter Request Party details
            driver = investigationMasterCasePage.SwitchNewBrowserWithTitle(driver, BaseWindow, "Task");
           

            TaskPage taskPage = new TaskPage(driver);
            string[] subjects = new string[] {
                "Relevant address added", 
                "Allocate to investigator",
                "Relevant bond background searches",
                "Follow-up request to be sent to complainant(s)",
                "Request to be sent (compelled by law)",
                "Follow-up request to be sent to third party",
                "Visit complainant",
                "Visit other witness",
                "Visit subject premises",
                "Visit other premises",
                "Visit respondent",
                "Execute warrant(s)",
                "Follow-up request to be sent to respondent",
                "Acknowledgement to be sent to respondent",
                "Send closure letter to respondent(s)",
                "Send closure letter to complainant(s)",
                "Additional investigations required",
                "Seek executive approval to prosecute",
                "Issue PIN",
                "Awaiting PIN referral to SPER"
            };

            foreach (string i in subjects)
            {
                taskPage.ClickPageTitle();
                Thread.Sleep(2000);
                taskPage.SetSelectSubjectValue(i);
                StringAssert.Contains(taskPage.GetSubjectValue(), i);
                
            }
           
            taskPage.ClickSaveCloseButton();

            driver = driver.SwitchTo().Window(BaseWindow);


            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsCasesRibbonButton();

            InvestigationCaseSearchPage investigationCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationCaseSearchPage.ClickNewInvestigationCaseButton();

            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            Thread.Sleep(3000);
            investigationCasePage.ClickSaveButton();
            investigationCasePage.GetInvestigationCaseNumber();
            investigationCasePage.ClickActivitiesAddButton();
            investigationCasePage.ClickCRMToolbar();
            BaseWindow = driver.CurrentWindowHandle; //Records the current window handle
            investigationCasePage.ClickAddTaskButton("Task");
            Thread.Sleep(3000);
            //Enter Request Party details
            driver = investigationCasePage.SwitchNewBrowserWithTitle(driver, BaseWindow, "Task");


            taskPage = new TaskPage(driver);
            subjects = new string[] {
                "Relevant address added", 
                "Allocate to investigator",
                "Relevant bond background searches",
                "Follow-up request to be sent to complainant(s)",
                "Request to be sent (compelled by law)",
                "Follow-up request to be sent to third party",
                "Visit complainant",
                "Visit other witness",
                "Visit subject premises",
                "Visit other premises",
                "Visit respondent",
                "Execute warrant(s)",
                "Follow-up request to be sent to respondent",
                "Acknowledgement to be sent to respondent",               
                "Send closure letter to complainant(s)",
                "Send closure letter to respondent(s)",
                "Additional investigations required",
                "Seek executive approval to prosecute",
                "Issue PIN",
                "Awaiting PIN referral to SPER"
            };

            foreach (string i in subjects)
            {
                taskPage.ClickPageTitle();
                taskPage.SetSelectSubjectValue(i);
                StringAssert.Contains(taskPage.GetSubjectValue(), i);

            }

            taskPage.ClickSaveCloseButton();
            driver = driver.SwitchTo().Window(BaseWindow);

        }






        [TestMethod]
        [TestProperty("TestcaseID", "3367")]
        public void ATC3367a_CRMSendemailfromCRMNewActivity()
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

            User user = this.environment.GetUser(SecurityRole.Investigations);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsCasesRibbonButton();

            InvestigationCaseSearchPage investigationsCaseSearchPage = new InvestigationCaseSearchPage(driver);

            investigationsCaseSearchPage.ClickNewInvestigationCaseButton();

            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            string user02 = "IMSTestU02";
            Assert.IsTrue(investigationCasePage.GetInvestigatorSearchElementText(user02));
            investigationCasePage.ClickSaveButton();
            string investigationCase = investigationCasePage.GetInvestigationCaseNumber();
            investigationCasePage.ClickSaveCloseButton();
            investigationsCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationsCaseSearchPage.SetInvestigationSearchText(investigationCase);

            Table table = new Table(investigationsCaseSearchPage.GetSearchResultTable());
            StringAssert.Contains(table.GetCellValue("Case Number", investigationCase, "Investigator"), user02);

            MyRange.Cells[MyRow, InvestigationSchema.GetColumnIndex(ColumnName.IN_ID)].Value = investigationCase;

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }

        [TestMethod]
        [TestProperty("TestcaseID", "3367")]
        //[DeploymentItem()]
        public void ATC3367c_CRMSendemailfromCRMNewActivity()
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

            string investigationID = MyRange.Cells[MyRow, InvestigationSchema.GetColumnIndex(ColumnName.IN_ID)].Value.ToString();

            User user = this.environment.GetUser(SecurityRole.Investigations);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsCasesRibbonButton();

            InvestigationCaseSearchPage investigationsCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationsCaseSearchPage.SetInvestigationSearchText(investigationID);
            Table table = new Table(investigationsCaseSearchPage.GetSearchResultTable());
            table.ClickCellValue("Case Number", investigationID, "Case Number");


            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            StringAssert.Contains(investigationCasePage.GetInvestigationCaseNumber(), investigationID);
            table = new Table(investigationCasePage.GetActivitiesSearchResultTable());
            StringAssert.Contains(table.GetCellContainsValue("Subject", "Test 3367 CRM Email Creation RTA:", "Activity Status"), "Completed");


            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }

        [TestMethod]
        [TestProperty("TestcaseID", "3285")]
        [TestProperty("TestType", "Regression")]
        public void ATC3285_CRMInvestigationMasterCaseConfirmpresenceofFieldsonform()
        {
            string investigationMasterCaseNumber;
            User user = this.environment.GetUser(SecurityRole.Investigations);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsMasterCasesRibbonButton();

            InvestigationMasterCaseSearchPage investigationsMasterCaseSearchPage = new InvestigationMasterCaseSearchPage(driver);

            investigationsMasterCaseSearchPage.ClickNewButton();

            InvestigationMasterCasePage investigationMasterCasePage = new InvestigationMasterCasePage(driver);
            
            investigationMasterCasePage.SetUnknowInvestigatorValue("Test User");
            investigationMasterCasePage.SetClientValue("BLAIR TEST");
            //investigationMasterCasePage.SetOwnerValue("IMSTestU04");
            investigationMasterCasePage.SetInvestigatorValue("IMSTestU02");
            investigationMasterCasePage.ClickSaveButton();
            investigationMasterCaseNumber = investigationMasterCasePage.GetInvestigationMasterCaseNumber();
            investigationMasterCasePage.ClickSaveCloseButton();


        }

        [TestMethod]
        [TestProperty("TestcaseID", "3341")]
        [TestProperty("TestType", "Regression")]
        public void ATC3341_CRMInvestigationNewCaseDefaultsToNewCaseStatus()
        {
           
            User user = this.environment.GetUser(SecurityRole.Investigations);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsCasesRibbonButton();

            Trace.Listeners.Add(new TextWriterTraceListener("TextWriterOutput.log", "myListener"));
            Trace.TraceInformation("Create new investigation case start:" + DateTime.Now.ToString("ddMMyyyyhhmmssffff"));
            Trace.Flush();

            InvestigationCaseSearchPage investigationsCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationsCaseSearchPage.ClickNewInvestigationCaseButton();

            // Create new case and validate Status and Substatus fields
            InvestigationCasePage investigationCasePage = new InvestigationCasePage(driver);
            Assert.AreEqual("New case", investigationCasePage.GetStatus());
            Assert.AreEqual("Creation", investigationCasePage.GetSubStatus());
            investigationCasePage.ClickSaveButton();
            String caseNumber=investigationCasePage.GetInvestigationCaseNumber();
            investigationCasePage.ClickSaveCloseButton();

            // Search for the newly created Investigation Case
            investigationsCaseSearchPage = new InvestigationCaseSearchPage(driver);
            investigationsCaseSearchPage.SetInvestigationSearchText(caseNumber);
            Table table = new Table(investigationsCaseSearchPage.GetSearchResultTable());
            table.ClickCellValue("Case Number", caseNumber, "Case Number");

            // Validate the activity created for New Investigation case
            investigationCasePage = new InvestigationCasePage(driver);
            StringAssert.Contains(investigationCasePage.GetInvestigationCaseNumber(), caseNumber);
            table = new Table(investigationCasePage.GetActivitiesSearchResultTable());
            StringAssert.Contains(table.GetCellContainsValue("Subject", caseNumber+": Scan documents", "Activity Status"), "Open");
            StringAssert.Contains(table.GetCellContainsValue("Subject", caseNumber + ": Add parties", "Activity Status"), "Open");
            StringAssert.Contains(table.GetCellContainsValue("Subject", caseNumber + ": Submit for initial assessment", "Activity Status"), "Open");

            Trace.Listeners.Add(new TextWriterTraceListener("TextWriterOutput.log", "myListener"));
            Trace.TraceInformation("Create new investigation case end:" + DateTime.Now.ToString("ddMMyyyyhhmmssffff"));
            Trace.Flush();                     
        }
        [TestMethod]
        [TestProperty("TestcaseID", "6707")]
        [TestProperty("TestType", "Regression")]
        public void ATC6707_CRMInvestigationNewCaseTasksInOutlook()
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

            // Validate the activity created for New Investigation case
            investigationCasePage = new InvestigationCasePage(driver);
            StringAssert.Contains(investigationCasePage.GetInvestigationCaseNumber(), caseNumber);
            table = new Table(investigationCasePage.GetActivitiesSearchResultTable());
            StringAssert.Contains(table.GetCellContainsValue("Subject", caseNumber + ": Scan documents", "Activity Status"), "Open");
            StringAssert.Contains(table.GetCellContainsValue("Subject", caseNumber + ": Add parties", "Activity Status"), "Open");
            StringAssert.Contains(table.GetCellContainsValue("Subject", caseNumber + ": Submit for initial assessment", "Activity Status"), "Open");
            
            // Navigate to Outlook and verify the task  - Paul
        }

        [TestMethod]
        [TestProperty("TestcaseID", "6737")]
        public void ATC6737_CRMGeneralCaseClienttobeOptional()
        {
            User user = this.environment.GetUser(SecurityRole.Investigations);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsGeneralCasesRibbonButton();

            InvestigationGeneralCaseSearchPage investigationsGeneralCaseSearchPage = new InvestigationGeneralCaseSearchPage(driver);

            investigationsGeneralCaseSearchPage.ClickNewGeneralCaseButton();

            InvestigationGeneralCasePage investigationGeneralCasePage = new InvestigationGeneralCasePage(driver);
            investigationGeneralCasePage.ClickPageTitle();
            
            investigationGeneralCasePage.SetTitle("Test client N/A - " + DateTime.Now);
            investigationGeneralCasePage.SetClientName("-NOT APPLICABLE-");
            investigationGeneralCasePage.SetType("Complaint");
            investigationGeneralCasePage.ClickSaveButton();
            investigationGeneralCasePage.GetGeneralCaseNumber();


        }

        [TestMethod]
        [TestProperty("TestcaseID", "3346")]
        public void ATC3346_CRMEntityGeneralCaseCreateNew()
        {
            User user = this.environment.GetUser(SecurityRole.Investigations);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickInvestigationsRibbonButton();
            homePage.HoverInvestigationsRibbonTab();
            homePage.ClickInvestigationsGeneralCasesRibbonButton();

            InvestigationGeneralCaseSearchPage investigationsGeneralCaseSearchPage = new InvestigationGeneralCaseSearchPage(driver);

            investigationsGeneralCaseSearchPage.ClickNewGeneralCaseButton();

            InvestigationGeneralCasePage investigationGeneralCasePage = new InvestigationGeneralCasePage(driver);
            investigationGeneralCasePage.ClickPageTitle();

            investigationGeneralCasePage.SetTitle("Test client N/A - " + DateTime.Now);
            investigationGeneralCasePage.SetClientName("-NOT APPLICABLE-");
            investigationGeneralCasePage.SetType("Complaint");
            investigationGeneralCasePage.SetReceivedDate((DateTime.Now.AddDays(1)).ToString("dd/MM/yyyy"));
            StringAssert.Contains(investigationGeneralCasePage.GetReceivedDateErrorMessage(), "Invalid date entered for Received Date");
            investigationGeneralCasePage.SetReceivedDate((DateTime.Now).ToString("dd/MM/yyyy"));

            investigationGeneralCasePage.ClickSaveButton();
            investigationGeneralCasePage.GetGeneralCaseNumber();


        }




    }
}
