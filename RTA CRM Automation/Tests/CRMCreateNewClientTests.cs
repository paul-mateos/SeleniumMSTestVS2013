using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.TestTools.UITesting;
using RTA.Automation.CRM.UI;
using RTA.Automation.CRM.Environment;
using OpenQA.Selenium;
using RTA.Automation.CRM.Pages;
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


namespace RTA.Automation.CRM.Tests
{
    [TestClass]
    public class CRMCreateNewClientTests : BaseTest
    {

        [TestInitialize]
        public void DataSourceSetup()
        {
            //open excel
            MyApp = new Excel.Application();
            MyApp.Visible = false;

        }

        [TestMethod]
        [TestProperty("TestcaseID", "4397")]
        public void ATC4397_CRMCheckClientID()
        {

            //Login in as role
            User user = this.environment.GetUser(SecurityRole.Default);
            new LoginDialog().Login(user.Id, user.Password);



            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();

            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);

            clientsSearchPage.ClickNewClientButton();
            
            ClientPage clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();
            clientPage.PopulateNewClient("UPINA");
            clientPage.ClickSaveButton();
            string clientId1 = clientPage.GetClientID();
            int clientNum1 = int.Parse(clientId1.Substring(1));

            clientPage.ClickNewClientButton();

            clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();
            clientPage.PopulateNewClient("UPINAA");
            clientPage.ClickSaveButton();

            string clientId2 = clientPage.GetClientID();
            int clientNum2 = int.Parse(clientId2.Substring(1));

            if (clientId1.StartsWith("C") && (clientId2.StartsWith("C")))  //checks if clientID starts with "C"
            {
                Assert.AreEqual(clientNum1, clientNum2 - 1);      //if clientNum2 is a single increment from clientnum1 then TRUE
            }
            else
            {
                Assert.Fail();
            }

        }


        [TestMethod]
        [TestProperty("TestcaseID", "4401")]
        public void ATC4401_CRMClientNameSuffixListtest()
        {
            #region Start Up Excel
            MyBook = MyApp.Workbooks.Open(DatasourceDir + @"\Clients.xlsx", 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
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

            string clientName = (MyRange.Cells[MyRow, ClientsSchema.GetColumnIndex("CLIENT_NAME")].Value.ToString());
            
            //Login in as role
            User user = this.environment.GetUser(SecurityRole.GeneralStaff);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();

            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);

            Trace.Listeners.Add(new TextWriterTraceListener("TextWriterOutput.log", "myListener"));
            Trace.TraceInformation("Search for client start:" + DateTime.Now.ToString("ddMMyyyyhhmmssffff"));
            Trace.Flush();

            clientsSearchPage.SetClientSearchText(clientName);
            Table table = new Table(clientsSearchPage.GetSearchResultTable());
            
            StringAssert.Equals(table.GetCellValue("Full Name", clientName, "Full Name"), clientName);

            Trace.Listeners.Add(new TextWriterTraceListener("TextWriterOutput.log", "myListener"));
            Trace.TraceInformation("Search for client end:" + DateTime.Now.ToString("ddMMyyyyhhmmssffff"));
            Trace.Flush();

            table.ClickCellValue("Full Name", clientName, "Full Name");

            ClientPage clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();
            #region assert suffix
            Assert.IsTrue(clientPage.GetSuffixText("AC"));
            Assert.IsTrue(clientPage.GetSuffixText("AM"));
            Assert.IsTrue(clientPage.GetSuffixText("AO"));
            Assert.IsTrue(clientPage.GetSuffixText("BEM"));
            Assert.IsTrue(clientPage.GetSuffixText("BM"));
            Assert.IsTrue(clientPage.GetSuffixText("CH"));
            Assert.IsTrue(clientPage.GetSuffixText("COMDC"));
            Assert.IsTrue(clientPage.GetSuffixText("CV"));
            Assert.IsTrue(clientPage.GetSuffixText("DFM"));
            Assert.IsTrue(clientPage.GetSuffixText("DSC"));
            Assert.IsTrue(clientPage.GetSuffixText("ESQ"));
            Assert.IsTrue(clientPage.GetSuffixText("GC"));
            Assert.IsTrue(clientPage.GetSuffixText("JNR"));
            Assert.IsTrue(clientPage.GetSuffixText("JP"));
            Assert.IsTrue(clientPage.GetSuffixText("MBE"));
            Assert.IsTrue(clientPage.GetSuffixText("MC"));
            Assert.IsTrue(clientPage.GetSuffixText("MHA"));
            Assert.IsTrue(clientPage.GetSuffixText("MHR"));
            Assert.IsTrue(clientPage.GetSuffixText("MLA"));
            Assert.IsTrue(clientPage.GetSuffixText("MLC"));
            Assert.IsTrue(clientPage.GetSuffixText("MP"));
            Assert.IsTrue(clientPage.GetSuffixText("OAM"));
            Assert.IsTrue(clientPage.GetSuffixText("OBE"));
            Assert.IsTrue(clientPage.GetSuffixText("OC"));
            Assert.IsTrue(clientPage.GetSuffixText("OM"));
            Assert.IsTrue(clientPage.GetSuffixText("QC"));
            Assert.IsTrue(clientPage.GetSuffixText("SC"));
            Assert.IsTrue(clientPage.GetSuffixText("SNR"));
            Assert.IsTrue(clientPage.GetSuffixText("VC"));
            Assert.IsTrue(clientPage.GetSuffixText("DSM"));
            #endregion
            
            clientPage.ClickSaveButton();
            string clientID = clientPage.GetClientID();
            clientPage.ClickSaveCloseButton();

            clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.SetClientSearchText(clientName);
            table = new Table(clientsSearchPage.GetSearchResultTable());

            StringAssert.Equals(table.GetCellValue("Full Name", clientName, "Full Name"), clientName);
            table.ClickCellValue("Full Name", clientName, "Full Name");
            clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();
            Assert.IsTrue(clientPage.GetSuffixText("DSM"));
           
            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

        [TestMethod]
        [TestProperty("TestcaseID", "4402")]
        public void ATC4402_CRMClientNameTitleListtest()
        {
            #region Start Up Excel
            MyBook = MyApp.Workbooks.Open(DatasourceDir + @"\Clients.xlsx", 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
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
            string clientName = (MyRange.Cells[MyRow, ClientsSchema.GetColumnIndex("CLIENT_NAME")].Value.ToString());
            //Login in as role
            User user = this.environment.GetUser(SecurityRole.GeneralStaff);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();

            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);

            clientsSearchPage.SetClientSearchText(clientName);
            Table table = new Table(clientsSearchPage.GetSearchResultTable());


            StringAssert.Equals(table.GetCellValue("Full Name", clientName, "Full Name"), clientName);
            table.ClickCellValue("Full Name", clientName, "Full Name");

            ClientPage clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();
            #region assert suffix
            //DR FR MISS MJR MR MRS MS PROF
            Assert.IsTrue(clientPage.GetTitleText("DR"));
            Assert.IsTrue(clientPage.GetTitleText("FR"));
            Assert.IsTrue(clientPage.GetTitleText("MISS"));
            Assert.IsTrue(clientPage.GetTitleText("MJR"));
            Assert.IsTrue(clientPage.GetTitleText("MR"));
            Assert.IsTrue(clientPage.GetTitleText("MRS"));
            Assert.IsTrue(clientPage.GetTitleText("MS"));
            Assert.IsTrue(clientPage.GetTitleText("PROF"));
            #endregion
            
            clientPage.ClickSaveButton();
            string clientID = clientPage.GetClientID();
            clientPage.ClickSaveCloseButton();

            clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.SetClientSearchText(clientName);
            table = new Table(clientsSearchPage.GetSearchResultTable());

            StringAssert.Equals(table.GetCellValue("Full Name", clientName, "Full Name"), clientName);
            table.ClickCellValue("Full Name", clientName, "Full Name");
            clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();
            Assert.IsTrue(clientPage.GetTitleText("PROF"));

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }

        [TestMethod]
        [TestProperty("TestcaseID", "4405")]
        public void ATC4405_CRMBusinessHoursandAfterHoursphone()
        {
            #region Start Up Excel
            MyBook = MyApp.Workbooks.Open(DatasourceDir + @"\Clients.xlsx", 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
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

            string clientName = (MyRange.Cells[MyRow, ClientsSchema.GetColumnIndex("CLIENT_NAME")].Value.ToString());
            //Login in as role
            User user = this.environment.GetUser(SecurityRole.GeneralStaff);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();

            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);

            clientsSearchPage.SetClientSearchText(clientName);
            Table table = new Table(clientsSearchPage.GetSearchResultTable());


            StringAssert.Equals(table.GetCellValue("Full Name", clientName, "Full Name"), clientName);
            table.ClickCellValue("Full Name", clientName, "Full Name");



            //Navigate to client phone numbers
            homePage.HoverClientXRibbonTab(clientName);
            homePage.ClickClientXPhoneNumbersRibbonButton();

            //Add new phone numbers
            ClientPage clientPage = new ClientPage(driver);
            string BaseWindow = driver.CurrentWindowHandle;
            clientPage.ClickAddNewClientPhoneImage();

            //Enter payment reference details
            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "Client Phone Number");
            //Assert availability list
            ClientPhoneNumberPage clientPhoneNumberPage = new ClientPhoneNumberPage(driver);
            Assert.IsTrue(clientPhoneNumberPage.GetAvailabilityListItem("Business hours only"));
            Assert.IsTrue(clientPhoneNumberPage.GetAvailabilityListItem("After hours only"));
            Assert.IsTrue(clientPhoneNumberPage.GetAvailabilityListItem("Anytime"));

            //Create 3 new phone numbers
                //clientPhoneNumberPage.ClickTypeList();
                clientPhoneNumberPage.SetTypeListValue("Fixed Line");
                //clientPhoneNumberPage.ClickAreaCodeElement();
                clientPhoneNumberPage.SetAreaCodeValue("07");
                //clientPhoneNumberPage.ClickPhoneNumberElement();
                clientPhoneNumberPage.SetPhoneNumberValue("11111111");
                clientPhoneNumberPage.ClickSaveButton();
                //Assert new phone number has saved
                string phoneNumber = clientPhoneNumberPage.GetPhoneNumber();

                clientPhoneNumberPage.ClickNewButton();
                clientPhoneNumberPage.ClickPageTitle();
                clientPhoneNumberPage.SetClientNameList(clientName);
                //clientPhoneNumberPage.SetClientNameListValue("BLAIR TEST");
                //clientPhoneNumberPage.ClickTypeList();
                clientPhoneNumberPage.SetTypeListValue("Mobile");
                //clientPhoneNumberPage.ClickPhoneNumberElement();
                clientPhoneNumberPage.SetPhoneNumberValue("0422222222");
                clientPhoneNumberPage.ClickSaveButton();
                //Assert new phone number has saved
                string phoneNumber2 = clientPhoneNumberPage.GetPhoneNumber();

                clientPhoneNumberPage.ClickNewButton();
                clientPhoneNumberPage.ClickPageTitle();
                clientPhoneNumberPage.SetClientNameList(clientName);
                //clientPhoneNumberPage.SetClientNameListValue("BLAIR TEST");
                //clientPhoneNumberPage.ClickTypeList();
                clientPhoneNumberPage.SetTypeListValue("Fax");
                //clientPhoneNumberPage.ClickAreaCodeElement();
                clientPhoneNumberPage.SetAreaCodeValue("07");
                //clientPhoneNumberPage.ClickPhoneNumberElement();
                clientPhoneNumberPage.SetPhoneNumberValue("33333333");
                clientPhoneNumberPage.ClickSaveButton();
                //Assert new phone number has saved
                string phoneNumber3 = clientPhoneNumberPage.GetPhoneNumber();
              
            clientPhoneNumberPage.ClickSaveCloseButton();
            clientPhoneNumberPage.ClickSaveCloseButton();
            clientPhoneNumberPage.ClickSaveCloseButton();

            driver = driver.SwitchTo().Window(BaseWindow);

            clientPage.SetClientSearchText(clientName);
            table = new Table(clientPage.GetSearchResultTable());

            StringAssert.Equals(table.GetCellValue("Client", clientName, "Full Phone Number"), phoneNumber);

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

        [TestMethod]
        [TestProperty("TestcaseID", "4407")]
        public void ATC4407_CRMAlertsValidationforClient()
        {
            #region Start Up Excel
            MyBook = MyApp.Workbooks.Open(DatasourceDir + @"\Clients.xlsx", 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
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

            string clientName = (MyRange.Cells[MyRow, ClientsSchema.GetColumnIndex("CLIENT_NAME")].Value.ToString());

            //Login in as role
            User user = this.environment.GetUser(SecurityRole.GeneralStaff);
            new LoginDialog().Login(user.Id, user.Password);



            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();

            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);
            string testName = clientName;
            clientsSearchPage.SetClientSearchText(testName);
            Table table = new Table(clientsSearchPage.GetSearchResultTable());
            table.ClickCellValue("Full Name", testName, "Full Name");

            ClientPage clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();
            string BaseWindow = driver.CurrentWindowHandle;
            
            clientPage.ClickAddAlertElement();

            //Enter Alert details
            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "Alert");
            AlertPage alertPage = new AlertPage(driver);
            Assert.IsTrue(alertPage.GetAlertTypeText("Other"));

            alertPage.SetOtherAlertText("Other Alert Test Value");
            alertPage.SetDesctiptionText("Description Text Test Value");

            alertPage.ClickSaveIMG();
            StringAssert.Contains(alertPage.GetAlertNumber(), testName);

            Assert.IsTrue(alertPage.GetAlertTypeText("Receivership/Management"));

            //Assert that Other Alert is locked
            string controlMode = alertPage.GetOtherAlertControlMode();
            StringAssert.Equals(controlMode, "locked");
            alertPage.SetStartDateValue(DateTime.Now.ToString("dd/MM/yyyy"));
            alertPage.SetEndDateValue(DateTime.Now.ToString("dd/MM/yyyy"));
            alertPage.ClickSaveIMG();
            //Assert.IsTrue(alertPage.GetNoErrorMessageState());
            StringAssert.Contains(alertPage.GetAlertState(), "Active");

            
            alertPage.SetEndDateValue(DateTime.Now.AddDays(1).ToString("dd/MM/yyyy"));
            alertPage.ClickSaveIMG();
            //Assert.IsTrue(alertPage.GetNoErrorMessageState());

            
            alertPage.SetEndDateValue(DateTime.Now.AddDays(-1).ToString("dd/MM/yyyy"));
            alertPage.ClickSaveIMG();
            StringAssert.Contains(alertPage.GetStartDateErrorText(), "Start date must be earlier than or the same as end date");
            alertPage.ClickPageTitle();
            //alertPage.ClickStartDate();
            alertPage.SetStartDateValue(DateTime.Now.AddDays(-1).ToString("dd/MM/yyyy"));
            alertPage.ClickSaveIMG();
            //Assert.IsTrue(alertPage.GetNoErrorMessageState());

            alertPage.ClickDeactivateIMG();
            StringAssert.Contains(alertPage.GetAlertState(), "Inactive");
            driver.Close();
            driver.SwitchTo().Window(BaseWindow);

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }

        [TestMethod]
        [TestProperty("TestcaseID", "6828")]
        public void ATC6828_CRMUnknownClientfeatureEnabledforInvestigationsuser()
        {

            //Login in as role
            User user = this.environment.GetUser(SecurityRole.InvestigationsOfficer);
            new LoginDialog().Login(user.Id, user.Password);



            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();

            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);

            clientsSearchPage.ClickNewClientButton();

            ClientPage clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();
            StringAssert.Contains(clientPage.GetUnknownClientListValues(), "YesNo");
        }

        [TestMethod]
        [TestProperty("TestcaseID", "4413")]
        public void ATC4413_CRMClientAgreedtoEmailCorrespondencenewclientsettoyes()
        {

            //Login in as role
            User user = this.environment.GetUser(SecurityRole.Default);
            new LoginDialog().Login(user.Id, user.Password);



            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();

            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);

            clientsSearchPage.ClickNewClientButton();


            ClientPage clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();
            clientPage.ClickPreferencesTab();
            StringAssert.Contains(clientPage.GetEmailCorrespondenceValue(), "No");
            clientPage.SetEmailCorrespondenceValue("Yes");


            clientPage.PopulateNewClient("EmailCorrespondenceTest");
            clientPage.ClickSaveButton();
            string clientId = clientPage.GetClientID();

            clientPage.ClickSaveCloseButton();


        }
        
      
        [TestMethod]
        [TestProperty("TestcaseID", "3348")]
        public void ATC3348_CRMEntityClientNameCreateNew() {
            //3348: Entity: Client Name - Create New

            #region Start Up Excel
            MyBook = MyApp.Workbooks.Open(DatasourceDir + @"\Clients.xlsx", 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            MySheet = (Excel.Worksheet)MyBook.Sheets[Properties.Settings.Default.ENVIRONMENT.ToString()];
            MyRange = MySheet.UsedRange;

            //Get specific row for the data
            int testDataRows = MyRange.Rows.Count;
            System.Diagnostics.Debug.WriteLine("testDataRows Present: " + testDataRows.ToString());
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

            string clientName = (MyRange.Cells[MyRow, ClientsSchema.GetColumnIndex("CLIENT_NAME")].Value.ToString());
            String Title = (MyRange.Cells[MyRow, ClientsSchema.GetColumnIndex("TITLE")].Value.ToString());
            String GivenName = (MyRange.Cells[MyRow, ClientsSchema.GetColumnIndex("GIVEN_NAME")].Value.ToString());
            //String MiddleName = "Jackie";
            String MiddleName = (MyRange.Cells[MyRow, ClientsSchema.GetColumnIndex("MIDDLE_NAME")].Value.ToString());
            
            String FamilyName = (MyRange.Cells[MyRow, ClientsSchema.GetColumnIndex("FAMILY_NAME")].Value.ToString());
            //String Suffix = "AM";
            String Suffix = (MyRange.Cells[MyRow, ClientsSchema.GetColumnIndex("SUFFIX")].Value.ToString());
            



            //Login in as role
            User user = this.environment.GetUser(SecurityRole.InvestigationsManager);
            new LoginDialog().Login(user.Id, user.Password);
            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();
            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);

            clientsSearchPage.SetClientSearchText("BLAIR TEST");
            Table table = new Table(clientsSearchPage.GetSearchResultTable());

            StringAssert.Equals(table.GetCellValue("Full Name", "BLAIR TEST", "Full Name"), "BLAIR TEST");
            table.ClickCellValue("Full Name", "BLAIR TEST", "Full Name");
            //Navigate to Cient Name
            homePage.HoverClientXRibbonTab("BLAIR TEST");
            homePage.ClickClientXClientNamesRibbonButton();  

            //Add new Cient Name
            ClientPage clientPage = new ClientPage(driver);
            string BaseWindow = driver.CurrentWindowHandle;
            System.Diagnostics.Debug.WriteLine("BaseWindow handle :" + BaseWindow.ToString());
            clientPage.ClickAddNewClientName();

            

            //*****************This needs to be moved out of here********************************************
            string NewWindow = ""; //prepares for the new window handle

            
            ReadOnlyCollection<string> handles = null;
            //Check for Alert Present
            System.Diagnostics.Debug.WriteLine("Check for Alert Present");
            try
            {
                if (driver.SwitchTo().Alert() != null) {
                    System.Diagnostics.Debug.WriteLine("Alert Present: " + driver.SwitchTo().Alert().Text);
                    System.Diagnostics.Debug.WriteLine("alert type: " + driver.SwitchTo().Alert().GetType().ToString()); 
                    driver.SwitchTo().Alert().Accept(); // prepares Selenium to handle alert 
            }
            }
            catch (NoAlertPresentException)
            { 
                // no alert message
                System.Diagnostics.Debug.WriteLine("No Alert Present");
            }

            //switch to new window
            //for some reason UICommon.SwitchToNewBrowser does not detect there is a new window, so I'll reinstate the old code below
            // generally the old code detects the new window in under 10 sec. However, UICommon.SwitchToNewBrowser runs for 30sec and fails to find the new window, although it should find it. 
            //driver = UICommon.SwitchToNewBrowser(driver, BaseWindow);           

            //Check for New Window Present
            for (int i = 1; i < 10; i++)
            {
                if (driver.WindowHandles.Count == 1)
                { Thread.Sleep(1000); }
                else { break; }
            }
             
            handles = driver.WindowHandles;
            System.Diagnostics.Debug.WriteLine("window handle count :" + handles.Count);
            foreach (string handle in handles)
            {
                System.Diagnostics.Debug.WriteLine("window handle :" + handle.ToString());
                var Handles = handle;
                if (BaseWindow != handle)
                {
                    NewWindow = handle;

                    driver = driver.SwitchTo().Window(NewWindow);
                    System.Diagnostics.Debug.WriteLine("SwitchTo() New Window: " + NewWindow.ToString());
                    break;
                }
            }
            
            
            //Populate Client Name Page
            ClientNamePage clientNamePage = new ClientNamePage(driver);
            
            //select Client Title
            //clientNamePage.ClickTitleList();
            clientNamePage.SetTitleListValue(Title);
            //clientNamePage.ClickGivenName();
            clientNamePage.SetGivenNameValue(GivenName);
            //clientNamePage.ClickMiddleName();
            clientNamePage.SetMiddleNameValue(MiddleName);
            //clientNamePage.ClickFamilyName();
            clientNamePage.SetFamilyNameValue(FamilyName);
            //clientNamePage.ClickSuffixList();
            clientNamePage.SetSuffixListValue(Suffix);
            clientNamePage.ClickSaveButton();
            //switch back to Client Name Page window. This may not be necessary as we are already there.
            //switch to new window
            //driver = UICommon.SwitchToNewBrowser(driver, BaseWindow);

            string FormTitle_Expected = Title + " " + GivenName + " " + MiddleName + " " + FamilyName + " " + Suffix;
            //string FormTitle_Expected = Title + " " + GivenName + " " + FamilyName;
            
            //Get the saved FormTitle Text which gets displayed on the screen after the record is saved.
            //string FormTitle = clientNamePage.GetFormTitle("MISS Jan Jackie Rivers AM");
            string FormTitle_Actual = clientNamePage.GetFormTitle();
            System.Diagnostics.Debug.WriteLine("FormTitle_Actual: " + FormTitle_Actual);
            System.Diagnostics.Debug.WriteLine("FormTitle_Expected: " + FormTitle_Expected);
            //Assert clientName has saved. 
            Assert.IsTrue(FormTitle_Actual.Equals(FormTitle_Expected.ToUpper()), "Actual Form Title does not match the expected (saved) value: " + FormTitle_Expected.ToUpper());
            //Save and close the Client Name Page       
            clientNamePage.ClickSaveCloseButton();

        }

        [TestMethod]
        [TestProperty("TestcaseID", "3349")]
        public void ATC3349_CRMEntityClientIdentificationArtefactCreateNew()
        {
            //3349: Entity: Client Identification Artefact - Create New
            #region Start Up Excel
            MyBook = MyApp.Workbooks.Open(DatasourceDir + @"\Clients.xlsx", 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            MySheet = (Excel.Worksheet)MyBook.Sheets[Properties.Settings.Default.ENVIRONMENT.ToString()];
            MyRange = MySheet.UsedRange;

            //Get specific row for the data
            int testDataRows = MyRange.Rows.Count;
            System.Diagnostics.Debug.WriteLine("testDataRows Present: " + testDataRows.ToString());
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

            string clientName = (MyRange.Cells[MyRow, ClientsSchema.GetColumnIndex("CLIENT_NAME")].Value.ToString());
            String ClientIdProvided = (MyRange.Cells[MyRow, ClientsSchema.GetColumnIndex("ID_PROVIDED")].Value.ToString());
            //Login in as role
            User user = this.environment.GetUser(SecurityRole.InvestigationsManager);
            new LoginDialog().Login(user.Id, user.Password);
            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();
            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);

            clientsSearchPage.SetClientSearchText("BLAIR TEST");
            Table table = new Table(clientsSearchPage.GetSearchResultTable());

            StringAssert.Equals(table.GetCellValue("Full Name", "BLAIR TEST", "Full Name"), "BLAIR TEST");
            table.ClickCellValue("Full Name", "BLAIR TEST", "Full Name");
            //Navigate to Cient Name
            homePage.HoverClientXRibbonTab("BLAIR TEST");
            homePage.ClickClientXClientIdArtefactRibbonButton();

            //Add new Cient Name
            ClientPage clientPage = new ClientPage(driver);
            string BaseWindow = driver.CurrentWindowHandle;
            System.Diagnostics.Debug.WriteLine("BaseWindow handle :" + BaseWindow.ToString());
            clientPage.ClickAddNewClientIdArtefact();
            string Title = "Client Identification Artefact";
            
            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, Title);
            //Populate Client Identification Artefact Page
            ClientIdentificationArtefactPage clientIdentificationArtefactPagePage = new ClientIdentificationArtefactPage(driver);
            clientIdentificationArtefactPagePage.SetClientIdProvided(ClientIdProvided);

            clientIdentificationArtefactPagePage.ClickSaveButton();
            string FormTitle_Expected = clientName.ToUpper() + ": " + ClientIdProvided;
            //Get the saved FormTitle Text which gets displayed on the screen after the record is saved.
            string FormTitle_Actual = clientIdentificationArtefactPagePage.GetFormTitle(ClientIdProvided);
            System.Diagnostics.Debug.WriteLine("FormTitle_Actual: " + FormTitle_Actual);
            System.Diagnostics.Debug.WriteLine("FormTitle_Expected: " + FormTitle_Expected);
            //Assert Client Identification Artefact has saved. 
            Assert.IsTrue(FormTitle_Actual.Equals(FormTitle_Expected), "Actual Form Title does not match the expected (saved) value: " + FormTitle_Expected);
            //Save and close the Client Identification Artefact Page       
            clientIdentificationArtefactPagePage.ClickSaveCloseButton();


        }

        [TestMethod]
        [TestProperty("TestcaseID", "3351")]
        public void ATC3351_CRMEntityClientNameCreateNewWithInformation()
        {
            #region Start Up Excel
            MyBook = MyApp.Workbooks.Open(DatasourceDir + @"\Clients.xlsx", 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
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

            string clientName = (MyRange.Cells[MyRow, ClientsSchema.GetColumnIndex("CLIENT_NAME")].Value.ToString());
            string EmailID = (MyRange.Cells[MyRow, ClientsSchema.GetColumnIndex("EMAIL")].Value.ToString());

            //Login in as role
            User user = this.environment.GetUser(SecurityRole.Investigations);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            string HomeWindow = driver.CurrentWindowHandle;
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();

            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.ClickNewClientButton();

            // Add new client name 
            ClientPage clientPage = new ClientPage(driver);


            // Fill in mandatory fields
            clientPage.PopulateNewClient(clientName);

            // Fill in Email address
            clientPage.SetEmail1ID(EmailID);
            clientPage.ClickSaveButton();
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

            // Navigate to client phone numbers
            driver.SwitchTo().Window(HomeWindow);
            homePage.HoverClientXRibbonTab(clientName);
            homePage.ClickClientXPhoneNumbersRibbonButton();
            clientPage = new ClientPage(driver);
            BaseWindow = driver.CurrentWindowHandle;

            clientPage.ClickAddNewClientPhoneImage();

            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "Client Phone Number");

            ClientPhoneNumberPage clientPhoneNumberPage = new ClientPhoneNumberPage(driver);
            //clientPhoneNumberPage.ClickTypeList();
            clientPhoneNumberPage.SetTypeListValue("Fixed Line");
            //clientPhoneNumberPage.ClickAreaCodeElement();
            clientPhoneNumberPage.SetAreaCodeValue("07");
            //clientPhoneNumberPage.ClickPhoneNumberElement();
            clientPhoneNumberPage.SetPhoneNumberValue("11111111");
            clientPhoneNumberPage.ClickSaveCloseButton();

            driver = driver.SwitchTo().Window(BaseWindow);

            clientPage.SetClientSearchText(clientName);
            Table table = new Table(clientPage.GetSearchResultTable());

            StringAssert.Contains(table.GetCellValue("Client", clientName, "Full Phone Number"), "+61 7 1111 1111");

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }
        [TestMethod]
        [TestProperty("TestcaseID", "4419")]
        public void ATC4419_CRMEntityClientDOBFieldTest()
        {
            #region Start Up Excel
            MyBook = MyApp.Workbooks.Open(DatasourceDir + @"\Clients.xlsx", 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
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

            string clientName = (MyRange.Cells[MyRow, ClientsSchema.GetColumnIndex("CLIENT_NAME")].Value.ToString());

            //Login in as role
            User user = this.environment.GetUser(SecurityRole.Investigations);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();

            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.ClickNewClientButton();

            // Add new client name 
            ClientPage clientPage = new ClientPage(driver);
            string BaseWindow = driver.CurrentWindowHandle;

            // Fill in mandatory fields
            clientPage.PopulateNewClient(clientName);
            clientPage.ClickSaveButton();

            Assert.IsTrue(clientPage.CheckFamilyNameErrorPresent(), "Error in saving Family Name");

            // Set {Date of Birth} field to tomorrow.
            clientPage.SetDateOfBirthValue(DateTime.Now.AddDays(1).ToString("dd/MM/yyyy"));
            clientPage.ClickSaveButton();
            StringAssert.Contains(clientPage.GetBirthdayErrorText(), "Invalid Date of Birth entered, please correct.");

            // Set {Date of Birth} field to make client less than 16 years old.
            clientPage.SetDateOfBirthValue(DateTime.Now.AddYears(-5).ToString("dd/MM/yyyy"));
            clientPage.ClickSaveButton();
            StringAssert.Contains(clientPage.GetBirthdayErrorText(), "Invalid Date of Birth entered, please correct.");

            // Set {Date of Birth} field to make client have year of birth prior to 1900 (e.g. 1899).
            clientPage.SetDateOfBirthValue("01/01/1899");
            clientPage.ClickSaveButton();
            StringAssert.Contains(clientPage.GetBirthdayErrorText(), "The specified date format is invalid or the date is out of valid range. Enter a valid date in the format: d/MM/yyyy");

            // Set {Date of Birth} field to make client 16 years old.
            driver.Navigate().Refresh();
            clientPage = new ClientPage(driver);
            clientPage.SetDateOfBirthValue(DateTime.Now.AddYears(-16).ToString("dd/MM/yyyy"));
            clientPage.ClickSaveButton();
            Assert.IsFalse(clientPage.CheckDateOfBirthErrorPresent(), "Error in saving date of birth field to make client 16 years old.");

            // Update {Date of Birth} field to tomorrow.
            clientPage.SetDateOfBirthValue(DateTime.Now.AddDays(1).ToString("dd/MM/yyyy"));
            clientPage.ClickSaveButton();
            StringAssert.Contains(clientPage.GetBirthdayErrorText(), "Invalid Date of Birth entered, please correct.");

            // Update {Date of Birth} field to make client less than 16 years old.
            clientPage.SetDateOfBirthValue(DateTime.Now.AddYears(-6).ToString("dd/MM/yyyy"));
            clientPage.ClickSaveButton();
            StringAssert.Contains(clientPage.GetBirthdayErrorText(), "Invalid Date of Birth entered, please correct.");

            // Set {Date of Birth} field to make client have year of birth prior to 1900 (e.g. 1899).
            clientPage.SetDateOfBirthValue("01/01/1899");
            clientPage.ClickSaveButton();
            StringAssert.Contains(clientPage.GetBirthdayErrorText(), "The specified date format is invalid or the date is out of valid range. Enter a valid date in the format: d/MM/yyyy");

            // Remove error message
            clientPage.SetDateOfBirthValue("");
            clientPage.ClickSaveButton();

            // Update {Date of Birth} field to 01/01/1900
            driver.Navigate().Refresh();
            clientPage = new ClientPage(driver);
            clientPage.SetDateOfBirthValue("01/01/1900");
            clientPage.ClickSaveButton();
            Assert.IsFalse(clientPage.CheckDateOfBirthErrorPresent(), "Error in saving date of birth field to 01/01/1900");

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

        [TestMethod]
        [TestProperty("TestcaseID", "3347")]
        public void ATC3347_CRMEntityCreateNewClientPersonalMobileTest()
        {
            
            #region Start Up Excel
            MyBook = MyApp.Workbooks.Open(DatasourceDir + @"\Clients.xlsx", 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
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

            string clientName = (MyRange.Cells[MyRow, ClientsSchema.GetColumnIndex("CLIENT_NAME")].Value.ToString());
            
            //Login in as role
            User user = this.environment.GetUser(SecurityRole.Investigations);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();
            
           

            //Search for already existing client
            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);
            Trace.Listeners.Add(new TextWriterTraceListener("TextWriterOutput.log", "myListener"));
            Trace.TraceInformation("Search for client start:" + DateTime.Now.ToString("ddMMyyyyhhmmssffff"));
            Trace.Flush();
            clientsSearchPage.SetClientSearchText(clientName);
            
            Table table = new Table(clientsSearchPage.GetSearchResultTable());
            StringAssert.Equals(table.GetCellValue("Full Name", clientName, "Full Name"), clientName);
          
            Trace.Listeners.Add(new TextWriterTraceListener("TextWriterOutput.log", "myListener"));
            Trace.TraceInformation("Search for client end:" + DateTime.Now.ToString("ddMMyyyyhhmmssffff"));
            Trace.Flush();

            table.ClickCellValue("Full Name", clientName, "Full Name");

            ClientPage clientPage = new ClientPage(driver);
            String BaseWindow = driver.CurrentWindowHandle;

            // Provide new Personal preferred Mobile number
            clientPage.ClickAddNewPersonalPreferredClientNumber();

            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "Client Phone Number");

            ClientPhoneNumberPage clientPhoneNumberPage = new ClientPhoneNumberPage(driver);
            //clientPhoneNumberPage.ClickTypeList();
            clientPhoneNumberPage.SetTypeListValue("Mobile");
            //clientPhoneNumberPage.ClickPhoneNumberElement();
            clientPhoneNumberPage.SetPhoneNumberValue("0422 222 222");
            clientPhoneNumberPage.ClickSaveCloseButton();

            driver = driver.SwitchTo().Window(BaseWindow);
            clientPage.ClickSaveButton();

            // On Clients page verify that Personal preferred mobile number is correctly populated
            StringAssert.Equals(clientPage.GetPersonalPreferredMobileNumber(), "+61 422 222 222");

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }

        [TestMethod]
        [TestProperty("TestcaseID", "3337")]
        public void ATC3337_CRMEntityClientConfirmvalidationofABN()
        {

            //Login in as role
            User user = this.environment.GetUser(SecurityRole.SystemAdministrator);
            new LoginDialog().Login(user.Id, user.Password);



            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();

            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);

            clientsSearchPage.ClickNewClientButton();

            ClientPage clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();

            clientPage.SetABNACNValue("124324542747");
            clientPage.GetABNACNErrorText();

            clientPage.SetABNACNValue("7898281890n");
            clientPage.GetABNACNErrorText();

            clientPage.SetABNACNValue("78982818908");
            clientPage.GetABNACNErrorTextNotExist();

            clientPage.SetABNACNValue("005249981");
            clientPage.GetABNACNErrorTextNotExist();

            clientPage.SetARBNValue("123456789");
            clientPage.GetARBNErrorText();

            clientPage.SetARBNValue("12345678n");
            clientPage.GetARBNErrorText();

            clientPage.SetARBNValue("005249981");
            clientPage.GetARBNErrorTextNotExist();


        }

        [TestMethod]
        [TestProperty("TestcaseID", "3339")]
        public void ATC3339_CRMEntityClientEmailshouldnotbeautocapitalised()
        {

            //Login in as role
            User user = this.environment.GetUser(SecurityRole.SystemAdministrator);
            new LoginDialog().Login(user.Id, user.Password);



            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();

            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);

            clientsSearchPage.ClickNewClientButton();

            ClientPage clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();

            clientPage.setEmail1Address("test@test.com");
            clientPage.GetEmailAddress1ErrorTextNotExist();
            Assert.AreEqual(clientPage.GetBannerEmailValue(), "test@test.com", true, "Email comparison failed");

            clientPage.setEmail1Address("TEST@TEST.COM");
            clientPage.GetEmailAddress1ErrorTextNotExist();
            Assert.AreEqual(clientPage.GetBannerEmailValue(), "TEST@TEST.COM", true, "Email comparison failed");


            clientPage.setEmail1Address("test@TESTCOM");
            StringAssert.Contains(clientPage.GetEmail1ErrorText(), "Invalid text entered for Email 1");
            clientPage.setEmail1Address("test@test.com");
            clientPage.GetEmailAddress1ErrorTextNotExist();

            clientPage.setEmail1Address("test@.COM");
            StringAssert.Contains(clientPage.GetEmail1ErrorText(), "Invalid text entered for Email 1");
            clientPage.setEmail1Address("test@test.com");
            clientPage.GetEmailAddress1ErrorTextNotExist();

            clientPage.setEmail1Address("test@test.C");
            StringAssert.Contains(clientPage.GetEmail1ErrorText(), "Invalid text entered for Email 1");
            clientPage.setEmail1Address("test@test.com");
            clientPage.GetEmailAddress1ErrorTextNotExist();

        
        }


        [TestMethod]
        [TestProperty("TestcaseID", "3350")]
        public void ATC3350_CRMEntityClientAddressCreateNew()
        {
            #region Start Up Excel
            MyBook = MyApp.Workbooks.Open(DatasourceDir + @"\Clients.xlsx", 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            MySheet = (Excel.Worksheet)MyBook.Sheets[Properties.Settings.Default.ENVIRONMENT.ToString()];
            MyRange = MySheet.UsedRange;

            //Get specific row for the data
            int testDataRows = MyRange.Rows.Count;
            int MyRow = 0;
            for (int i = 2; i <= testDataRows; i++)
            {
                if (MyRange.Cells[i, 1].Value.ToString() == "3350")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string address = (MyRange.Cells[MyRow, ClientsSchema.GetColumnIndex("ADDRESS")].Value.ToString());

            //Login in as role
            User user = this.environment.GetUser(SecurityRole.SystemAdministrator);
            new LoginDialog().Login(user.Id, user.Password);



            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();

            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);
            
            Trace.Listeners.Add(new TextWriterTraceListener("TextWriterOutput.log", "myListener"));
            Trace.TraceInformation("Add New Address start:" + DateTime.Now.ToString("ddMMyyyyhhmmssffff"));
            Trace.Flush();
            
            clientsSearchPage.ClickNewClientButton();

            ClientPage clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();
            clientPage.SetClientType("Person");
            clientPage.SetFamilyName("TESTING ADDRESS");
            clientPage.ClickSaveButton();
            string clientID = clientPage.GetClientID();
            homePage.HoverClientXRibbonTab("TESTING ADDRESS");
            homePage.ClickClientXAddressesRibbonButton();
            string BaseWindow = driver.CurrentWindowHandle;

            clientPage = new ClientPage(driver);
            Thread.Sleep(3000);
            clientPage.ClickAddNewClientAddressImage();
            driver = clientPage.SwitchNewBrowser(driver, BaseWindow, "Client Address:");
            ClientNewAddressPage clientNewAddressPage = new ClientNewAddressPage(driver);
            clientNewAddressPage.ClickPageTitle();
            clientNewAddressPage.SetAddressDetails("*" + address);
            clientNewAddressPage.ClickSaveButton();
            clientNewAddressPage.GetCleintAddress();
            
            Trace.Listeners.Add(new TextWriterTraceListener("TextWriterOutput.log", "myListener"));
            Trace.TraceInformation("Add New Address End:" + DateTime.Now.ToString("ddMMyyyyhhmmssffff"));
            Trace.Flush();

            clientNewAddressPage.ClickSaveAndClose();

            driver = driver.SwitchTo().Window(BaseWindow);
            clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();
            clientPage = new ClientPage(driver);
            clientPage.SetPostalAddress("*" + address);
            clientPage.ClickSaveCloseButton();
            clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.SetClientSearchText(clientID);
            Table table = new Table(clientsSearchPage.GetSearchResultTable());
            Assert.AreEqual(table.GetCellValue("RTA Client Id", clientID, "Postal Address"), address, "Address comparison failed");

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }

        [TestMethod]
        [TestProperty("TestcaseID", "3350p")]
        public void ATC3350p_CRMEntityClientAddressCreateNew()
        {
            #region Start Up Excel
            MyBook = MyApp.Workbooks.Open(DatasourceDir + @"\Clients.xlsx", 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            MySheet = (Excel.Worksheet)MyBook.Sheets[Properties.Settings.Default.ENVIRONMENT.ToString()];
            MyRange = MySheet.UsedRange;

            //Get specific row for the data
            int testDataRows = MyRange.Rows.Count;
            int MyRow = 0;
            for (int i = 2; i <= testDataRows; i++)
            {
                if (MyRange.Cells[i, 1].Value.ToString() == "3350")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string address = (MyRange.Cells[MyRow, ClientsSchema.GetColumnIndex("ADDRESS")].Value.ToString());

            //Login in as role
            User user = this.environment.GetUser(SecurityRole.SystemAdministrator);
            new LoginDialog().Login(user.Id, user.Password);



            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();

            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);

            Trace.Listeners.Add(new TextWriterTraceListener("TextWriterOutput.log", "myListener"));
            Trace.TraceInformation("Add New Address start:" + DateTime.Now.ToString("ddMMyyyyhhmmssffff"));
            Trace.Flush();

            clientsSearchPage.ClickNewClientButton();

            ClientPage clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();
            clientPage.SetClientType("Person");
            clientPage.SetFamilyName("TESTING ADDRESS");
            clientPage.ClickSaveButton();
            string clientID = clientPage.GetClientID();
            homePage.HoverClientXRibbonTab("TESTING ADDRESS");
            homePage.ClickClientXAddressesRibbonButton();
            string BaseWindow = driver.CurrentWindowHandle;

            clientPage = new ClientPage(driver);
            clientPage.ClickAddNewClientAddressImageIRSIT();
            driver = clientPage.SwitchNewBrowser(driver, BaseWindow, "Client Address:");
            ClientNewAddressPage clientNewAddressPage = new ClientNewAddressPage(driver);
            clientNewAddressPage.SetAddressDetails("*" + address);
            clientNewAddressPage.ClickSaveButton();
            clientNewAddressPage.GetCleintAddress();

            Trace.Listeners.Add(new TextWriterTraceListener("TextWriterOutput.log", "myListener"));
            Trace.TraceInformation("Add New Address End:" + DateTime.Now.ToString("ddMMyyyyhhmmssffff"));
            Trace.Flush();

            clientNewAddressPage.ClickSaveAndClose();

            driver = driver.SwitchTo().Window(BaseWindow);
            clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();
            clientPage = new ClientPage(driver);
            clientPage.ClickSaveButton();
            clientPage.SetPostalAddress("*" + address);
            clientPage.ClickSaveCloseButton();

            clientsSearchPage = new ClientsSearchPage(driver);
            Thread.Sleep(3000);
            clientsSearchPage.SetClientSearchText(clientID);
            Table table = new Table(clientsSearchPage.GetSearchResultTable());
            Assert.AreEqual(table.GetCellValue("RTA Client Id", clientID, "Postal Address"), address, "Address comparison failed");

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }

        [TestMethod]
        [TestProperty("TestcaseID", "4421")]
        public void ATC4421_CRMClientPostalAddressWarning()
        {
            //Login in as role
            User user = this.environment.GetUser(SecurityRole.SystemAdministrator);
            new LoginDialog().Login(user.Id, user.Password);

            //(Template) - Create new client (person) via Client Services group
            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();

            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);

            clientsSearchPage.ClickNewClientButton();

            ClientPage clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();

            //Complete mandatory fields
            clientPage.SetClientType("Person");
            clientPage.SetFamilyName("TC POSTALADDRESSWARN");

            //Inspect screen for warning message regarding non-population of {Postal Address}.
            Assert.IsTrue(clientPage.VerifyWarningMessagePresent("rta_postaladdressid"), "Postal Address blank message NOT displayed");
            StringAssert.Contains(clientPage.GetWarningMessage("rta_postaladdressid"), "Postal Address is blank, please select an address.");

            //Attempt to save record with of {Postal Address} unpopulated. Record saves
            clientPage.ClickSaveButton();

            string clientID = clientPage.GetClientID();
            clientPage.ClickSaveCloseButton();

            //Reenter record.Warning message is redisplayed.
            clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.SetClientSearchText(clientID);
            
            Table table = new Table(clientsSearchPage.GetSearchResultTable());
            table.ClickCellValue("RTA Client Id", clientID, "Full Name");

            clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();

            Assert.IsTrue(clientPage.VerifyWarningMessagePresent("rta_postaladdressid"), "Postal Address blank message NOT displayed");
            StringAssert.Contains(clientPage.GetWarningMessage("rta_postaladdressid"), "Postal Address is blank, please select an address.");
        }

        [TestMethod]
        [TestProperty("TestcaseID", "4430")]
        public void ATC4430_CRMClientPhysicalAddressWarning()
        {
            //Login in as role
            User user = this.environment.GetUser(SecurityRole.SystemAdministrator);
            new LoginDialog().Login(user.Id, user.Password);

            //(Template) - Create new client (person) via Client Services group
            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();

            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);

            clientsSearchPage.ClickNewClientButton();

            ClientPage clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();

            //Complete mandatory fields
            clientPage.SetClientType("Person");
            clientPage.SetFamilyName("TC PHYSICALADDRESSWARN");

            //Inspect screen for warning message regarding non-population of {Postal Address}.
            Assert.IsTrue(clientPage.VerifyWarningMessagePresent("rta_physicaladdressid"), "Physical Address blank message NOT displayed");
            StringAssert.Contains(clientPage.GetWarningMessage("rta_physicaladdressid"), "Physical Address is blank, please select an address.");

            //Attempt to save record with of {Physical Address} unpopulated. Record saves
            clientPage.ClickSaveButton();

            string clientID = clientPage.GetClientID();
            clientPage.ClickSaveCloseButton();

            //Reenter record.Warning message is redisplayed.
            clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.SetClientSearchText(clientID);

            Table table = new Table(clientsSearchPage.GetSearchResultTable());
            table.ClickCellValue("RTA Client Id", clientID, "Full Name");

            clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();

            Assert.IsTrue(clientPage.VerifyWarningMessagePresent("rta_physicaladdressid"), "Physical Address blank message NOT displayed");
            StringAssert.Contains(clientPage.GetWarningMessage("rta_physicaladdressid"), "Physical Address is blank, please select an address.");
        }

        
        [TestMethod]
        [TestProperty("TestcaseID", "4439")]
        public void ATC4439_ClientAlertsValidation()
        {
            //Login in as role
            User user = this.environment.GetUser(SecurityRole.Investigations);
            new LoginDialog().Login(user.Id, user.Password);

            //Navigate to: Client Services group > Clients tile
            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();

            //Double-click on a Organisation record.
            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.ClickNewClientButton();

            ClientPage clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();

            clientPage.SetClientType("Organisation");
            clientPage.SetOrganizationName("TC AlertValidation");
            clientPage.ClickSaveButton();
            string clientID = clientPage.GetClientID();
            clientPage.ClickSaveCloseButton();

            clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.SetClientSearchText(clientID);

            Table table = new Table(clientsSearchPage.GetSearchResultTable());
            table.ClickCellValue("RTA Client Id", clientID, "Full Name");

            //Inspect Current Alerts section.Click [+] button.
            clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();
            string BaseWindow = driver.CurrentWindowHandle;
            clientPage.ClickAddAlertElement();

            //Set {Alert Type} to 'Other'.{Other Alert} field becomes mandatory and is enabled.
            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "Alert");
            AlertPage alertPage = new AlertPage(driver);
            alertPage.SetAlertTypeElement("Other");
            Assert.IsTrue(alertPage.VerifyOtherAlertMandatoryFieldIconPresent());

            //Enter free text into {Other Alert} field and {Description}.
            alertPage.SetOtherAlertText("Other Alert One");
            alertPage.SetDesctiptionText("Description Text Test Value One");

            //Saved.Displays as current alert.As per entered.
            alertPage.ClickSaveCloseIMG();
            driver = driver.SwitchTo().Window(BaseWindow);    
            clientPage = new ClientPage(driver);
            
            Table alertTable = new Table(clientPage.GetCurrentAlertsTable());
            StringAssert.Contains(alertTable.GetCellValue("Other Alert", "Other Alert One", "Alert Type"),"Other","Validating whether other alert displayed in current alerts");

            BaseWindow = driver.CurrentWindowHandle;
            //Create another Alert:Set {Start date}.Set {End date} to today.Save record.
            clientPage.ClickAddAlertElement();
            Thread.Sleep(5000);

            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "Alert");
            alertPage = new AlertPage(driver);
            alertPage.SetAlertTypeElement("Other");

            Assert.IsTrue(alertPage.VerifyOtherAlertMandatoryFieldIconPresent());

            alertPage.SetOtherAlertText("Other Alert Two");
            alertPage.SetDesctiptionText("Description Text Test Value Two");
            alertPage.SetStartDateValue(DateTime.Now.ToString("dd/MM/yyyy"));
            alertPage.SetEndDateValue(DateTime.Now.ToString("dd/MM/yyyy"));

            alertPage.ClickSaveCloseIMG();
            driver.SwitchTo().Window(BaseWindow);

            //Saved.Displays as current alert.As per entered.
            clientPage = new ClientPage(driver);
            alertTable = new Table(clientPage.GetCurrentAlertsTable());
            StringAssert.Contains(alertTable.GetCellValue("Other Alert", "Other Alert Two", "Alert Type"), "Other", "Validating whether other alert displayed in current alerts");

            //Modify record:Set {End date} to yesterday.Save record.        
            IWebElement alertElement = alertTable.GetCellElementContainsValue("Other Alert", "Other Alert Two", "Other Alert");
            UICommon.DoubleClickElement(alertElement, driver);

            alertPage = new AlertPage(driver);
            alertPage.SetStartDateValue(DateTime.Now.AddDays(-1).ToString("dd/MM/yyyy"));
            alertPage.SetEndDateValue(DateTime.Now.AddDays(-1).ToString("dd/MM/yyyy"));
            alertPage.ClickSaveCloseIMG();

            //Saved.NOT displayed as current alert.As per entered.
            clientPage = new ClientPage(driver);
            alertTable = new Table(clientPage.GetCurrentAlertsTable());
            Assert.IsFalse(alertTable.MatchingCellFound("Other Alert","Other Alert Two"), "Validating the alert with end date as yeterday not displayed in current alerts");

            //Create more then four current Alerts.
            BaseWindow = driver.CurrentWindowHandle;
            clientPage.ClickAddAlertElement();
            Thread.Sleep(3000);
          
            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "Alert");
            alertPage = new AlertPage(driver);
            alertPage.SetAlertTypeElement("Other");

            Assert.IsTrue(alertPage.VerifyOtherAlertMandatoryFieldIconPresent());

            alertPage.SetOtherAlertText("Other Alert Two");
            alertPage.SetDesctiptionText("Description Text Test Value Two");
            alertPage.SetStartDateValue(DateTime.Now.ToString("dd/MM/yyyy"));
            alertPage.SetEndDateValue(DateTime.Now.ToString("dd/MM/yyyy"));

            alertPage.ClickSaveCloseIMG();
            driver = driver.SwitchTo().Window(BaseWindow);

            clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();

            BaseWindow = driver.CurrentWindowHandle;
            clientPage.ClickAddAlertElement();
            Thread.Sleep(3000);

            //Enter Alert details
            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "Alert");
            alertPage = new AlertPage(driver);
            alertPage.SetAlertTypeElement("Other");

            Assert.IsTrue(alertPage.VerifyOtherAlertMandatoryFieldIconPresent());

            alertPage.SetOtherAlertText("Other Alert Three");
            alertPage.SetDesctiptionText("Description Text Test Value Three");
            alertPage.SetStartDateValue(DateTime.Now.ToString("dd/MM/yyyy"));
            alertPage.SetEndDateValue(DateTime.Now.ToString("dd/MM/yyyy"));

            alertPage.ClickSaveCloseIMG();

            driver = driver.SwitchTo().Window(BaseWindow);
            
            clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();
            BaseWindow = driver.CurrentWindowHandle;
            clientPage.ClickAddAlertElement();
            Thread.Sleep(3000);
           
            //Enter Alert details
            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "Alert");
            alertPage = new AlertPage(driver);
            alertPage.SetAlertTypeElement("Other");

            Assert.IsTrue(alertPage.VerifyOtherAlertMandatoryFieldIconPresent());

            alertPage.SetOtherAlertText("Other Alert Four");
            alertPage.SetDesctiptionText("Description Text Test Value Four");
            alertPage.SetStartDateValue(DateTime.Now.ToString("dd/MM/yyyy"));
            alertPage.SetEndDateValue(DateTime.Now.ToString("dd/MM/yyyy"));

            alertPage.ClickSaveCloseIMG();
            driver = driver.SwitchTo().Window(BaseWindow);
           
            //Section is dynamic and allows for multi-page display.
            clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();
            clientPage.ClickNextPageAlertTable();

            alertTable = new Table(clientPage.GetCurrentAlertsTable());
            Assert.IsTrue(alertTable.MatchingCellFound("Other Alert", "Other Alert Four"), "Validating the fourth alert displayed in the next page of alert table");
        }

        [TestMethod]
        [TestProperty("TestcaseID", "4433")]
        public void ATC4433_CommonTriggerPointCopyAddressFromParent()
        {
            //Login to CRM as General staff user.
            User user = this.environment.GetUser(SecurityRole.GeneralStaff);
            new LoginDialog().Login(user.Id, user.Password);

            //Navigate to: Client Services group > Clients tile
            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();

            //Create Parent Organization with Postal address
            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.ClickNewClientButton();

            ClientPage clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();
            string parentOrganisation = "TCParentOrganization"+ UICommon.GetRandomString(5);

            clientPage.SetClientType("Organisation");
            clientPage.SetOrganizationName(parentOrganisation);
            clientPage.ClickSaveButton();

            string BaseWindow = driver.CurrentWindowHandle;
            clientPage.ClickCreateNewClientAddressButton("rta_postaladdressid");           
            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "Client Address");
            ClientNewAddressPage clientNewAddressPage = new ClientNewAddressPage(driver);
            clientNewAddressPage.SetAddressDetails("Australian Physical", 10, "GRACELAND");

            driver = driver.SwitchTo().Window(BaseWindow);
            clientPage = new ClientPage(driver);
            Assert.AreEqual("10 GRACELAND", clientPage.GetAddressValue("rta_postaladdressid"));

            clientPage.ClickSaveCloseButton();
            
            //Create child with parent organization has postal address 
            clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.ClickNewClientButton();

            clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();

            string childOrganisation = "TCChildOrganization"+ UICommon.GetRandomString(5);
            clientPage.SetClientType("Organisation");
            clientPage.SetOrganizationName(childOrganisation);
            clientPage.SetParentOrganization(parentOrganisation.ToUpper());
            clientPage.ClickSaveButton();
            string clientID = clientPage.GetClientID();
           
            //Select "Start Dialog" item from top entity menu.
            clientPage.ClickStartDialogButton();

            /*Run dialog "Copy an address from parent organisation":
             *For client where parent organisation has postal address but no physical address
             *Select YES for Postal address and YES for Physical addres
             */
            Table table = new Table(clientPage.GetProcessSearchResultTable());
            table.ClickCell("Process Name", "Copy an address from parent organisation", "Created On");

            BaseWindow = driver.CurrentWindowHandle;
            clientPage.ClickDialogAddButton();

            Thread.Sleep(2000);
            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "Copy an address");
            CopyAddressPage copyAddressPage = new CopyAddressPage(driver);

            copyAddressPage.SetCopyPhysicalAddressYes();
            copyAddressPage.SetCopyPostalAddressYes();

            //User is notified that no Physical address is copied.
            copyAddressPage.ClickNextButton();
            Assert.AreEqual(copyAddressPage.GetErrorMessage(), "The parent organisation does not have a physical address");

            /*For client where parent organisation has postal address but no physical address
             *Select NO for Postal address and YES for Physical address
             */
            copyAddressPage.ClickPreviousButton();
            copyAddressPage.SetCopyPostalAddressNo();
            copyAddressPage.SetCopyPhysicalAddressYes();
            copyAddressPage.ClickNextButton();
            ////User is notified that no Physical address is copied.
            Assert.AreEqual(copyAddressPage.GetErrorMessage(), "The parent organisation does not have a physical address");

            /* For client where parent organisation has postal address but no physical address
             * Select YES for Postal address and NO for Physical address
             */
            copyAddressPage.ClickPreviousButton();
            copyAddressPage.SetCopyPostalAddressYes();
            copyAddressPage.SetCopyPhysicalAddressNo();
            copyAddressPage.ClickNextButton();

            //Copy is successful.
            StringAssert.Contains(copyAddressPage.GetFinishMessage(), "This is the end of the dialog. Click Finish to close it");
            copyAddressPage.ClickFinishButton();

            driver = driver.SwitchTo().Window(BaseWindow);
            clientPage.ClickSaveCloseButton();
            Thread.Sleep(5000);
            clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.SetClientSearchText(clientID);

            Table searchTable = new Table(clientsSearchPage.GetSearchResultTable());
            searchTable.ClickCellValue("RTA Client Id", clientID, "Full Name");

            clientPage = new ClientPage(driver);
            Assert.AreEqual(clientPage.GetPostalAddress(), "10 GRACELAND");
        }

        [TestMethod]
        [TestProperty("TestcaseID", "3338")]
        public void ATC3338_CRMClientFilterPhoneNumbers()
        {
            //Login in as role
            User user = this.environment.GetUser(SecurityRole.GeneralStaff);
            new LoginDialog().Login(user.Id, user.Password);
            string clientName = "FILTER PHONE" + UICommon.GetRandomString(3); 

            HomePage homePage = new HomePage(driver);
            string HomeWindow = driver.CurrentWindowHandle;
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();


            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.ClickNewClientButton();
            ClientPage clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();
            clientPage.SetFamilyName(clientName);
            //clientPage.PopulateNewClient(clientName);
            clientPage.ClickSaveButton();
 
            //Navigate to client phone numbers
            driver = driver.SwitchTo().Window(HomeWindow);
            homePage.HoverClientXRibbonTab(clientName);
            homePage.ClickClientXPhoneNumbersRibbonButton();

            //Add new phone numbers
            clientPage = new ClientPage(driver);
            string BaseWindow = driver.CurrentWindowHandle;
            clientPage.ClickAddNewClientPhoneImage();

            //Enter payment reference details
            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "Client Phone Number");

            //Assert availability list
            ClientPhoneNumberPage clientPhoneNumberPage = new ClientPhoneNumberPage(driver);

            //Create Fixed Line Number
            //clientPhoneNumberPage.ClickTypeList();
            clientPhoneNumberPage.SetTypeListValue("Fixed Line");
            //clientPhoneNumberPage.ClickAreaCodeElement();
            clientPhoneNumberPage.SetAreaCodeValue("07");
            //clientPhoneNumberPage.ClickPhoneNumberElement();
            clientPhoneNumberPage.SetPhoneNumberValue("11111111");
            clientPhoneNumberPage.ClickSaveButton();
            //Assert new phone number has saved
            string FixedLineNumber = clientPhoneNumberPage.GetPhoneNumber();

            clientPhoneNumberPage.ClickNewButton();
            clientPhoneNumberPage.ClickPageTitle();
            clientPhoneNumberPage.SetClientNameList(clientName);
            //clientPhoneNumberPage.ClickTypeList();
            clientPhoneNumberPage.SetTypeListValue("Mobile");
            //clientPhoneNumberPage.ClickPhoneNumberElement();
            clientPhoneNumberPage.SetPhoneNumberValue("0422222222");
            clientPhoneNumberPage.ClickSaveButton();
            //Assert new phone number has saved
            string MobileNumber = clientPhoneNumberPage.GetPhoneNumber();

            clientPhoneNumberPage.ClickNewButton();
            clientPhoneNumberPage.ClickPageTitle();
            clientPhoneNumberPage.SetClientNameList(clientName);
            //clientPhoneNumberPage.ClickTypeList();
            clientPhoneNumberPage.SetTypeListValue("Fax");
            //clientPhoneNumberPage.ClickAreaCodeElement();
            clientPhoneNumberPage.SetAreaCodeValue("07");
            //clientPhoneNumberPage.ClickPhoneNumberElement();
            clientPhoneNumberPage.SetPhoneNumberValue("33333333");
            clientPhoneNumberPage.ClickSaveButton();
            //Assert new phone number has saved
            string FaxNumber =  clientPhoneNumberPage.GetPhoneNumber();

            clientPhoneNumberPage.ClickNewButton();
            clientPhoneNumberPage.ClickPageTitle();
            clientPhoneNumberPage.SetClientNameList(clientName);
            //clientPhoneNumberPage.ClickTypeList();
            clientPhoneNumberPage.SetTypeListValue("Pager");
            //clientPhoneNumberPage.ClickAreaCodeElement();
            clientPhoneNumberPage.SetAreaCodeValue("07");
            //clientPhoneNumberPage.ClickPhoneNumberElement();
            clientPhoneNumberPage.SetPhoneNumberValue("44444444");
            clientPhoneNumberPage.ClickSaveButton();
            //Assert new phone number has saved
            string PagerNumber =clientPhoneNumberPage.GetPhoneNumber();

            clientPhoneNumberPage.ClickSaveCloseButton();
            clientPhoneNumberPage.ClickSaveCloseButton();
            clientPhoneNumberPage.ClickSaveCloseButton();
            clientPhoneNumberPage.ClickSaveCloseButton();

            driver = driver.SwitchTo().Window(HomeWindow);
            homePage.ClickClientXRibbonTab(clientName);
            clientPage = new ClientPage(driver);

            // Select Personal Mobile field -> only mobile phone number for client can be selected
            clientPage.ClickPageTitle();
            clientPage.OpenMobileNumbeDropDown();
            Assert.IsTrue(clientPage.VerifyMobileNumberOptionPresent(MobileNumber), "Mobile Phone number seen for Personal Mobile Field !!!");
            Assert.IsFalse(clientPage.VerifyMobileNumberOptionPresent(FixedLineNumber), "Fixed Line number seen for Personal Mobile Field !!!");
            Assert.IsFalse(clientPage.VerifyMobileNumberOptionPresent(FaxNumber), "Fax number seen for Personal Mobile Field !!!");
            Assert.IsFalse(clientPage.VerifyMobileNumberOptionPresent(PagerNumber), "Pager number seen for Personal Mobile Field !!!");

            // Personal Preferred field -> only mobile,fax, fixed line or pager phone number for client can be selected
            clientPage.ClickPageTitle();
            clientPage.OpenPersonalPreferredNumbeDropDown();
            Assert.IsTrue(clientPage.VerifyPersonalPreferredNumberOptionPresent(MobileNumber), "Mobile Phone number NOT seen for Personal Preferred field !!!");
            Assert.IsTrue(clientPage.VerifyPersonalPreferredNumberOptionPresent(FixedLineNumber), "Fixed Line number NOT seen for Personal Preferred field !!!");
            Assert.IsTrue(clientPage.VerifyPersonalPreferredNumberOptionPresent(FaxNumber), "Fax number NOT seen for Personal Preferred field !!!");
            Assert.IsTrue(clientPage.VerifyPersonalPreferredNumberOptionPresent(PagerNumber), "Pager number NOT seen for Personal Preferred field !!!");

            // Work Preferred field -> only mobile,fax, fixed line or pager phone number for client can be selected
            clientPage.ClickPageTitle();
            clientPage.OpenWorkPreferredNumbeDropDown();
            Assert.IsTrue(clientPage.VerifyWorkPreferredOptionPresent(MobileNumber), "Mobile Phone number NOT seen for Work Preferred field !!!");
            Assert.IsTrue(clientPage.VerifyWorkPreferredOptionPresent(FixedLineNumber), "Fixed Line number NOT seen for Work Preferred field !!!");
            Assert.IsTrue(clientPage.VerifyWorkPreferredOptionPresent(FaxNumber), "Fax number NOT seen for Work Preferred field !!!");
            Assert.IsTrue(clientPage.VerifyWorkPreferredOptionPresent(PagerNumber), "Pager number NOT seen for Work Preferred field !!!");

            // Home / Main Phone field -> only mobile,fax, fixed line or pager phone number for client can be selected
            clientPage.ClickPageTitle();
            clientPage.OpenHomeMainPhoneNumbeDropDown();
            Assert.IsTrue(clientPage.VerifyHomeMainPhoneOptionPresent(MobileNumber), "Mobile Phone number NOT seen for Home / Main Phone field !!!");
            Assert.IsTrue(clientPage.VerifyHomeMainPhoneOptionPresent(FixedLineNumber), "Fixed Line number NOT seen for Home / Main Phone field !!!");
            Assert.IsTrue(clientPage.VerifyHomeMainPhoneOptionPresent(FaxNumber), "Fax number NOT seen for Home / Main Phone field !!!");
            Assert.IsTrue(clientPage.VerifyHomeMainPhoneOptionPresent(PagerNumber), "Pager number NOT seen for Home / Main Phone field !!!");

            // Fax field  -> only fax or fixed line phone number for client can be selected
            clientPage.ClickPageTitle();
            clientPage.OpenFaxPhoneNumbeDropDown();
            Assert.IsFalse(clientPage.VerifyFaxOptionPresent(MobileNumber), "Mobile Phone number seen for Fax field !!!");
            Assert.IsTrue(clientPage.VerifyFaxOptionPresent(FixedLineNumber), "Fixed Line number NOT seen for Fax field !!!");
            Assert.IsTrue(clientPage.VerifyFaxOptionPresent(FaxNumber), "Fax number NOT seen for Fax field !!!");
            Assert.IsFalse(clientPage.VerifyFaxOptionPresent(PagerNumber), "Pager number seen for Fax field !!!");
        }

        [TestMethod]
        [TestProperty("TestcaseID", "5513")]
        public void ATC5513_CRMClientUnknownClientUpdation()
        {
            //Login in as Investigations Officer role
            User user = this.environment.GetUser(SecurityRole.InvestigationsOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            string HomeWindow = driver.CurrentWindowHandle;
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();

            Trace.Listeners.Add(new TextWriterTraceListener("TextWriterOutput.log", "myListener"));
            Trace.TraceInformation("Create new client start:" + DateTime.Now.ToString("ddMMyyyyhhmmssffff"));
            Trace.Flush();

            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.ClickNewClientButton();
            ClientPage clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();
            clientPage.SetUnknownClientListValues("Yes");
            clientPage.SetGivenName("Given Name");
            clientPage.SetMiddleName("Middle Name");
            clientPage.ClickSaveButton();
            string ClientID = clientPage.GetClientID();
            clientPage.ClickSaveCloseButton();

            Trace.Listeners.Add(new TextWriterTraceListener("TextWriterOutput.log", "myListener"));
            Trace.TraceInformation("Create new client end:" + DateTime.Now.ToString("ddMMyyyyhhmmssffff"));
            Trace.Flush();


            // Select VIEW: Active Clients - UNKNOWN-
            clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.SetPageFilterList("Active Clients -UNKNOWN-");

            Table table = new Table(clientsSearchPage.GetHeaderSearchResultTable());
            table.ClickTableColumnHeader("RTA Client Id");
            table.ClickTableColumnHeader("RTA Client Id");

            // allow the client record to be created for the first name and populate the client surname with a standard token that is easy to identify and search on
            table = new Table(clientsSearchPage.GetSearchResultTable());
            StringAssert.Contains(table.GetCellValue("RTA Client Id", ClientID, "Full Name"), "GIVEN NAME MIDDLE NAME -UNKNOWN-");
            table.ClickCellValue("RTA Client Id", ClientID, "Full Name");

            clientPage = new ClientPage(driver);
            //  allow Investigation User to update the client surname.
            clientPage.ClickPageTitle();
            clientPage.SetUnknownClientListValues("No");
            clientPage.SetFamilyName("Family Name");
            clientPage.ClickSaveCloseButton();

            Trace.Listeners.Add(new TextWriterTraceListener("TextWriterOutput.log", "myListener"));
            Trace.TraceInformation("Search by Client ID start:" + DateTime.Now.ToString("ddMMyyyyhhmmssffff"));
            Trace.Flush();

            clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.SetPageFilterList("Active Clients");
            clientsSearchPage.SetClientSearchText(ClientID);
            table = new Table(clientsSearchPage.GetSearchResultTable());
            StringAssert.Contains(table.GetCellValue("RTA Client Id", ClientID, "Full Name"), "GIVEN NAME MIDDLE NAME FAMILY NAME");

            Trace.Listeners.Add(new TextWriterTraceListener("TextWriterOutput.log", "myListener"));
            Trace.TraceInformation("Search by Client ID end:" + DateTime.Now.ToString("ddMMyyyyhhmmssffff"));
            Trace.Flush();


            table.ClickCellValue("RTA Client Id", ClientID, "Full Name");
            clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();
            clientPage.SetUnknownClientListValues("Yes");
            clientPage.ClickSaveCloseButton();

            driver.Quit();
            driver = null;

            this.TestSetup();

            user = this.environment.GetUser(SecurityRole.RecordKeepingOfficers);
            new LoginDialog().Login(user.Id, user.Password);

            homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();

            clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.SetPageFilterList("Active Clients -UNKNOWN-");

            table = new Table(clientsSearchPage.GetHeaderSearchResultTable());
            table.ClickTableColumnHeader("RTA Client Id");
            table.ClickTableColumnHeader("RTA Client Id");
            table = new Table(clientsSearchPage.GetSearchResultTable());
            table.ClickCellValue("RTA Client Id", ClientID, "Full Name");
            
            clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();

            // Verify that Family name and Unknown Client Fields are locked
            Assert.IsTrue(clientPage.VerifyElementLocked("rta_unknownclient"),"Unknown Field is NOT locked !!!!!");
            Assert.IsTrue(clientPage.VerifyElementLocked("lastname"), "Lastename Field is NOT locked !!!!!");
        }
        [TestMethod]
        [TestProperty("TestcaseID", "5505")]
        public void ATC5505_CRMClientUpdatePhoneNumbers()
        {
            //Login in as role
            User user = this.environment.GetUser(SecurityRole.GeneralStaff);
            new LoginDialog().Login(user.Id, user.Password);
            string clientName = "CLIENT PHONE" + UICommon.GetRandomString(3);

            HomePage homePage = new HomePage(driver);
            string HomeWindow = driver.CurrentWindowHandle;
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();


            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.ClickNewClientButton();
            ClientPage clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();
            clientPage.SetFamilyName(clientName);
            //clientPage.PopulateNewClient(clientName);
            clientPage.ClickSaveButton();

            //Navigate to client phone numbers
            driver = driver.SwitchTo().Window(HomeWindow);
            homePage.HoverClientXRibbonTab(clientName);
            homePage.ClickClientXPhoneNumbersRibbonButton();

            //Add new phone numbers
            clientPage = new ClientPage(driver);
            string BaseWindow = driver.CurrentWindowHandle;
            clientPage.ClickAddNewClientPhoneImage();

            //Enter payment reference details
            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "Client Phone Number");

            //Assert availability list
            ClientPhoneNumberPage clientPhoneNumberPage = new ClientPhoneNumberPage(driver);

            //Create Fixed Line Number
            //clientPhoneNumberPage.ClickTypeList();
            clientPhoneNumberPage.SetTypeListValue("Fixed Line");
            //clientPhoneNumberPage.ClickAreaCodeElement();
            clientPhoneNumberPage.SetAreaCodeValue("07");
            //clientPhoneNumberPage.ClickPhoneNumberElement();
            clientPhoneNumberPage.SetPhoneNumberValue("11111111");
            clientPhoneNumberPage.ClickSaveButton();
            //Assert new phone number has saved
            string FixedLineNumber =  clientPhoneNumberPage.GetPhoneNumber();

            clientPhoneNumberPage.ClickNewButton();
            clientPhoneNumberPage.ClickPageTitle();
            clientPhoneNumberPage.SetClientNameList(clientName);
            //clientPhoneNumberPage.ClickTypeList();
            clientPhoneNumberPage.SetTypeListValue("Mobile");
            //clientPhoneNumberPage.ClickPhoneNumberElement();
            clientPhoneNumberPage.SetPhoneNumberValue("0422222222");
            clientPhoneNumberPage.ClickSaveButton();
            //Assert new phone number has saved
            string MobileNumber = clientPhoneNumberPage.GetPhoneNumber();

            clientPhoneNumberPage.ClickNewButton();
            clientPhoneNumberPage.ClickPageTitle();
            clientPhoneNumberPage.SetClientNameList(clientName);
            //clientPhoneNumberPage.ClickTypeList();
            clientPhoneNumberPage.SetTypeListValue("Fax");
            //clientPhoneNumberPage.ClickAreaCodeElement();
            clientPhoneNumberPage.SetAreaCodeValue("07");
            //clientPhoneNumberPage.ClickPhoneNumberElement();
            clientPhoneNumberPage.SetPhoneNumberValue("33333333");
            clientPhoneNumberPage.ClickSaveButton();
            //Assert new phone number has saved
            string FaxNumber =  clientPhoneNumberPage.GetPhoneNumber();

            clientPhoneNumberPage.ClickNewButton();
            clientPhoneNumberPage.ClickPageTitle();
            clientPhoneNumberPage.SetClientNameList(clientName);
            //clientPhoneNumberPage.ClickTypeList();
            clientPhoneNumberPage.SetTypeListValue("Pager");
            //clientPhoneNumberPage.ClickAreaCodeElement();
            clientPhoneNumberPage.SetAreaCodeValue("07");
            //clientPhoneNumberPage.ClickPhoneNumberElement();
            clientPhoneNumberPage.SetPhoneNumberValue("44444444");
            clientPhoneNumberPage.ClickSaveButton();
            //Assert new phone number has saved
            string PagerNumber =clientPhoneNumberPage.GetPhoneNumber();

            clientPhoneNumberPage.ClickSaveCloseButton();
            clientPhoneNumberPage.ClickSaveCloseButton();
            clientPhoneNumberPage.ClickSaveCloseButton();
            clientPhoneNumberPage.ClickSaveCloseButton();

            // Verify that Activities table is not showing any activity
            driver = driver.SwitchTo().Window(HomeWindow);
            homePage.ClickClientXRibbonTab(clientName);
            clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();

            Table table = new Table(UICommon.GetSearchResultTable(driver));

            Assert.IsTrue(table.GetNoRecordsInTable(), "Client Phone Number Updation Activity Getting displayed !!!!!");

            // Update the Personal Mobile, Personal Preferred Contact Number, Work Preferred Contact Number and Fax and save the record
            clientPage.SetMobileNumber(MobileNumber);
            clientPage.SetPersonalPreferredMobileNumber(FixedLineNumber);
            clientPage.SetWorkPreferredNumber(PagerNumber);
            clientPage.SetFaxNumber(FaxNumber);

            clientPage.ClickSaveCloseButton();

            clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.SetClientSearchText(clientName);
            table = new Table(clientsSearchPage.GetSearchResultTable());
            table.ClickCellValue("Full Name", clientName, "Full Name");

            clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();
            Thread.Sleep(500);
            table = new Table(clientPage.GetActivitiesTable());
            Assert.IsTrue(table.MatchingCellFound("Subject", "Client Profile Phone Number updated"), "Client Phone Number Updation Activity NOT displayed !!!!!");
            StringAssert.Contains(table.GetCellValue("Subject", "Client Profile Phone Number updated", "Actual End"), DateTime.Now.ToString("dd/MM/yyyy"));
        }

        [TestMethod]
        [TestProperty("TestcaseID", "4444")]
        public void ATC4444_CRMClientEmailDelete()
        {
            //Login as CRM default user.
            User user = this.environment.GetUser(SecurityRole.GeneralStaff);
            new LoginDialog().Login(user.Id, user.Password);

            /*Data preparation step 
            Organisation Client is needed with following fields populated:
            - Email 1
            - Email 2*/

            string clientName = "CLIENTEMAILDEL" + UICommon.GetRandomString(3);

            HomePage homePage = new HomePage(driver);
            string HomeWindow = driver.CurrentWindowHandle;
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();

            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.ClickNewClientButton();
            ClientPage clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();
            clientPage.SetClientType("Organisation");
            clientPage.SetOrganizationName(clientName);
            clientPage.ClickSaveButton();

            string email1 = clientName + "1@gmail.com";
            string email2 = clientName + "2@gmail.com";

            clientPage.SetEmail1ID(clientName + "1@gmail.com");
            clientPage.SetEmail2ID(clientName + "2@gmail.com");

            clientPage.ClickSaveButton();

            Assert.AreEqual(email1, clientPage.GetEmail1ID(),"Validate the email id1 is set correctly");
            Assert.AreEqual(email2, clientPage.GetEmail2ID(),"Validate the email id2 is set correctly");
            
            string clientID = clientPage.GetClientID();
            clientPage.ClickSaveCloseButton();

            clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.SetClientSearchText(clientID);

            Table clientTable = new Table(clientsSearchPage.GetSearchResultTable());
            clientTable.ClickCellValue("RTA Client Id", clientID, "Full Name");
            
            //Clear the following fields and save record:- Email 1- Email 2
            clientPage = new ClientPage(driver);
            clientPage.ClearEmail1Id();
            clientPage.ClearEmail2Id();
            string date = DateTime.Today.ToString("d/MM/yyyy");
            string time = DateTime.Now.ToString("h:mm tt");

            clientPage.ClickSaveCloseButton();
            clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.SetClientSearchText(clientID);

            clientTable = new Table(clientsSearchPage.GetSearchResultTable());
            clientTable.ClickCellValue("RTA Client Id", clientID, "Full Name");

            clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();

            //Inspect resultant entry in Activity sub grid.
            Table headerTable = new Table(clientPage.GetActivitiesHeaderTable());
            headerTable.ClickTableColumnHeader("Actual End");
            Thread.Sleep(2000);

            Table activitiesTable = new Table(clientPage.GetActivitiesTable());

            //Client Management Activity is created as a note to users that the email address fields were changed (includes current value of each field and a note referring the user to audit history).
            Assert.IsTrue(activitiesTable.MatchingCellFound("Subject", "Client E-mail Address updated"), "Client Management Activity for Email field updation Created!!!");
            activitiesTable.ClickCellValue("Subject", "Client E-mail Address updated", "Subject");       

            ClientManagementActivityPage activityPage = new ClientManagementActivityPage(driver);
            activityPage.ClickPageTitle();
            StringAssert.Contains(activityPage.GetDescription(), "Email Address (primary):   [no value]");
            StringAssert.Contains(activityPage.GetDescription(), "Email Address (secondary):   [no value]");
            StringAssert.Contains(activityPage.GetDescription(), "Note: the current values are displayed above. Refer to Audit History against the Client record for more information.");

            //Select the [Audit History] option from the entity navigation menu in the global ribbon. Inspect audit history relating to deletions made in this test.
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();

            clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.SetClientSearchText(clientID);

            clientTable = new Table(clientsSearchPage.GetSearchResultTable());

            clientTable.ClickCellValue("RTA Client Id", clientID, "Full Name");

            clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();
            
            homePage.HoverClientRibbonTab(clientName);
            homePage.ClickClientXAuditRibbonButton();

            clientPage = new ClientPage(driver);
            Thread.Sleep(2000);
            
            /*Audit history shows:
            - What was deleted
            - Who deleted it
            - When it was deleted
            - What previous value was*/
            clientPage = new ClientPage(driver);
            Table auditTable = new Table(clientPage.GetAuditHistoryTable());
            StringAssert.Contains(auditTable.GetCellValue("Event", "Update", "Changed Field"), "Email Address (primary)");
            StringAssert.Contains(auditTable.GetCellValue("Event", "Update", "Changed Field"), "Email Address (secondary)");

            Assert.AreEqual(email1+System.Environment.NewLine+email2, auditTable.GetCellValue("Event", "Update", "Old Value"));
            Assert.AreEqual(String.Empty,auditTable.GetCellValue("Event", "Update", "New Value").Trim());
            Assert.AreEqual(user.Id.ToLower() + " user", auditTable.GetCellValue("Event", "Update", "Changed By").ToLower());
            Assert.AreEqual( date + " " +time, auditTable.GetCellValue("Event", "Update", "Changed Date"));
        }

        [TestMethod]
        [TestProperty("TestcaseID", "3343")]
        public void ATC3343_CRMAlterCloseReopenFrontCounterActivity()
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
            string FrontCounterActivity1 = "TC 3343 Front Counter Activity 1" + UICommon.GetRandomString(3);
            string FrontCounterActivity2 = "TC 3343 Front Counter Activity 2" + UICommon.GetRandomString(3);

            User user = this.environment.GetUser(SecurityRole.Investigations);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            String HomeWindow = driver.CurrentWindowHandle;
            homePage.ClickCreateIMG();
            homePage.ClickFrontCounterContactActivityRibbonButton();

            FrontCounterContactPage frontCounterContactPage = new FrontCounterContactPage(driver);
            frontCounterContactPage.ClickPageTitle();

            // Assign a Client and fill in all possisble fields.
            frontCounterContactPage.SetSelectSubjectValue("Bond existence");
            frontCounterContactPage.SetSubjectValue(FrontCounterActivity1);
            frontCounterContactPage.SetClientName(ClientName);
            frontCounterContactPage.ClickSaveCloseButton();

            driver = driver.SwitchTo().Window(HomeWindow);
            homePage.ClickCreateIMG();
            homePage.ClickFrontCounterContactActivityRibbonButton();

            frontCounterContactPage = new FrontCounterContactPage(driver);
            frontCounterContactPage.ClickPageTitle();

            // Assign a Client and fill in all possisble fields.
            frontCounterContactPage.SetSelectSubjectValue("Bond existence");
            frontCounterContactPage.SetSubjectValue(FrontCounterActivity2);
            frontCounterContactPage.SetClientName(ClientName);
            frontCounterContactPage.ClickSaveCloseButton();

            driver.Quit();
            driver = null;

            this.TestSetup();
            user = this.environment.GetUser(SecurityRole.SystemAdministrator);
            new LoginDialog().Login(user.Id, user.Password);

            homePage = new HomePage(driver);
            HomeWindow = driver.CurrentWindowHandle;
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();

            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.ClickNewClientButton();
            ClientPage clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();
            clientPage.PopulateNewClient("Test");
            clientPage.ClickSaveButton();
            string ClientID = clientPage.GetClientID();
            clientPage.ClickSaveCloseButton();

            driver = driver.SwitchTo().Window(HomeWindow);
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientActivitiesRibbonButton();

            // Open an existing Front Counter Contact activity (not created by your test user) from the list.
            ActivitiesSearchPage activitiesSearchPage = new ActivitiesSearchPage(driver);
            string BaseWindow = driver.CurrentWindowHandle;
            activitiesSearchPage.SetTenancyRequestSearchText(FrontCounterActivity1);

            Table table = new Table(activitiesSearchPage.GetSearchResultTable());
            table.ClickCellValue("Activity Type", "Front Counter Contact", "Subject");

            frontCounterContactPage = new FrontCounterContactPage(driver);
            // Modify an aspect of the activity - add a client and change the description and then save the Activity.
            frontCounterContactPage.SetClientName("TEST");
            frontCounterContactPage.SetDescription("Modify Front Counter Activity");
            frontCounterContactPage.ClickSaveButton();
            frontCounterContactPage.ClickPageTitle();

            // Activity saves without error.
            StringAssert.Contains(frontCounterContactPage.GetClientName(), "TEST");
            StringAssert.Contains(frontCounterContactPage.GetDescription(), "Modify Front Counter Activity");

            frontCounterContactPage.ClickSaveCloseButton();

            // Open another Front Counter Contact which has not been created by your test user.
            driver = driver.SwitchTo().Window(BaseWindow);
            activitiesSearchPage = new ActivitiesSearchPage(driver);
            activitiesSearchPage.SetTenancyRequestSearchText(FrontCounterActivity2);

            table = new Table(activitiesSearchPage.GetSearchResultTable());
            table.ClickCellValue("Activity Type", "Front Counter Contact", "Subject");

            frontCounterContactPage = new FrontCounterContactPage(driver);

            // Close the Front Counter Contact.Activity Closes without any error
            frontCounterContactPage.ClickPageTitle();
            Thread.Sleep(500);
            frontCounterContactPage.ClickCloseFrontCounterContactButton();
            Thread.Sleep(500);
            frontCounterContactPage.ClickDialogAddButton();
            Thread.Sleep(1000);

            // Ensure that the Activity does not allow editing.
            frontCounterContactPage.ClickPageTitle();
            Assert.IsTrue(UICommon.VerifyElementLocked("subject", driver), "Subject Field is NOT locked!!!");
            Assert.IsTrue(UICommon.VerifyElementLocked("customers", driver), "customers Field is NOT locked!!!");
            Assert.IsTrue(UICommon.VerifyElementLocked("actualend", driver), "actualend Field is NOT locked!!!");
            Assert.IsTrue(UICommon.VerifyElementLocked("actualdurationminutes", driver), "Actual Duration Field is NOT locked!!!");
            Assert.IsTrue(UICommon.VerifyElementLocked("regardingobjectid", driver), "Regarding Field is NOT locked!!!");
            Assert.IsTrue(UICommon.VerifyElementLocked("rta_assistive_service_usedid", driver), "Assistive Service Field is NOT locked!!!");
            Assert.IsTrue(UICommon.VerifyElementLocked("description", driver), "Description Field is NOT locked!!!");

            // Reopen the Front Counter Contact through the Dialog.
            frontCounterContactPage.ClickStartDialogButton();
            table = new Table(frontCounterContactPage.GetProcessSearchResultTable());
            table.ClickCell("Process Name", "Re-open Front Counter Contact activity", "Created On");
            BaseWindow = driver.CurrentWindowHandle;
            frontCounterContactPage.ClickDialogAddButton();

            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "Re-open Front Counter Contact activity");
            ReOpenCall reOpenActivity = new ReOpenCall(driver);
            reOpenActivity.ClickNextButton();
            reOpenActivity.ClickFinishButton();

            driver = driver.SwitchTo().Window(HomeWindow);
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientActivitiesRibbonButton();

            // Open an existing Front Counter Contact activity (not created by your test user) from the list.
            activitiesSearchPage = new ActivitiesSearchPage(driver);
            activitiesSearchPage.SetTenancyRequestSearchText(FrontCounterActivity2);

            table = new Table(activitiesSearchPage.GetSearchResultTable());
            table.ClickCellValue("Activity Type", "Front Counter Contact", "Subject");

            frontCounterContactPage = new FrontCounterContactPage(driver);
            frontCounterContactPage.ClickPageTitle();

            // Ensure that the Activity allows editing.
            Assert.IsFalse(UICommon.VerifyElementLocked("subject", driver), "Subject Field is locked!!!");
            Assert.IsFalse(UICommon.VerifyElementLocked("customers", driver), "customers Field is locked!!!");
            Assert.IsFalse(UICommon.VerifyElementLocked("actualend", driver), "actualend Field is locked!!!");
            Assert.IsFalse(UICommon.VerifyElementLocked("actualdurationminutes", driver), "Actual Duration Field is locked!!!");
            Assert.IsFalse(UICommon.VerifyElementLocked("regardingobjectid", driver), "Regarding Field is locked!!!");
            Assert.IsFalse(UICommon.VerifyElementLocked("rta_assistive_service_usedid", driver), "Assistive Service Field is locked!!!");
            Assert.IsFalse(UICommon.VerifyElementLocked("description", driver), "Description Field is locked!!!");

            frontCounterContactPage.SetDescription("Description updated after reopening activity");
            frontCounterContactPage.ClickPageTitle();

            StringAssert.Contains(frontCounterContactPage.GetDescription(), "Description updated after reopening activity");
        }

        [TestMethod]
        [TestProperty("TestcaseID", "4440")]
        public void ATC4440_CRMClientPhoneDelete()
        {
            //Login as CRM default user.
            User user = this.environment.GetUser(SecurityRole.GeneralStaff);
            new LoginDialog().Login(user.Id, user.Password);
            /*
            Person Client is needed with following fields populated:
            - Personal Mobile
            - Personal Preferred Contact Number
            - Work Preferred Contact Number
            - Fax*/
            string clientName = "CLIENT PHONEDEL" + UICommon.GetRandomString(3);

            HomePage homePage = new HomePage(driver);
            string HomeWindow = driver.CurrentWindowHandle;
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();

            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.ClickNewClientButton();
            ClientPage clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();
            Thread.Sleep(1000);

            clientPage.SetFamilyName(clientName);
            clientPage.ClickSaveButton();

            //Navigate to client phone numbers
            driver = driver.SwitchTo().Window(HomeWindow);
            homePage.HoverClientXRibbonTab(clientName);
            homePage.ClickClientXPhoneNumbersRibbonButton();

            //Add new phone numbers
            clientPage = new ClientPage(driver);
            string BaseWindow = driver.CurrentWindowHandle;
            clientPage.ClickAddNewClientPhoneImage();

            //Enter payment reference details
            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "Client Phone Number");

            //Assert availability list
            ClientPhoneNumberPage clientPhoneNumberPage = new ClientPhoneNumberPage(driver);

            //Create Fixed Line Number
            clientPhoneNumberPage.SetTypeListValue("Fixed Line");
            clientPhoneNumberPage.SetAreaCodeValue("07");
            clientPhoneNumberPage.SetPhoneNumberValue("11111111");
            clientPhoneNumberPage.ClickSaveButton();
            //Assert new phone number has saved
            string FixedLineNumber = clientPhoneNumberPage.GetPhoneNumber();

            clientPhoneNumberPage.ClickNewButton();
            clientPhoneNumberPage.ClickPageTitle();
            clientPhoneNumberPage.SetClientNameList(clientName);
            clientPhoneNumberPage.SetTypeListValue("Mobile");
            clientPhoneNumberPage.SetPhoneNumberValue("0422222222");
            clientPhoneNumberPage.ClickSaveButton();
            //Assert new phone number has saved
            string MobileNumber = clientPhoneNumberPage.GetPhoneNumber();

            clientPhoneNumberPage.ClickNewButton();
            clientPhoneNumberPage.ClickPageTitle();
            clientPhoneNumberPage.SetClientNameList(clientName);
            clientPhoneNumberPage.SetTypeListValue("Fax");
            clientPhoneNumberPage.SetAreaCodeValue("07");
            clientPhoneNumberPage.SetPhoneNumberValue("33333333");
            clientPhoneNumberPage.ClickSaveButton();
            //Assert new phone number has saved
            string FaxNumber = clientPhoneNumberPage.GetPhoneNumber();

            clientPhoneNumberPage.ClickNewButton();
            clientPhoneNumberPage.ClickPageTitle();
            clientPhoneNumberPage.SetClientNameList(clientName);
            clientPhoneNumberPage.SetTypeListValue("Pager");
            clientPhoneNumberPage.SetAreaCodeValue("07");
            clientPhoneNumberPage.SetPhoneNumberValue("44444444");
            clientPhoneNumberPage.ClickSaveButton();
            //Assert new phone number has saved
            string PagerNumber = clientPhoneNumberPage.GetPhoneNumber();
                        
            clientPhoneNumberPage.ClickSaveCloseButton();
            clientPhoneNumberPage.ClickSaveCloseButton();
            clientPhoneNumberPage.ClickSaveCloseButton();
            clientPhoneNumberPage.ClickSaveCloseButton();

            // Verify that Activities table is not showing any activity
            driver = driver.SwitchTo().Window(HomeWindow);
            homePage.ClickClientXRibbonTab(clientName);
            clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();

            Table table = new Table(UICommon.GetSearchResultTable(driver));

            Assert.IsTrue(table.GetNoRecordsInTable(), "Client Phone Number Updation Activity Getting displayed !!!!!");

            // Update the Personal Mobile, Personal Preferred Contact Number, Work Preferred Contact Number and Fax and save the record
            clientPage.SetMobileNumber(MobileNumber);
            clientPage.SetPersonalPreferredMobileNumber(FixedLineNumber);
            clientPage.SetWorkPreferredNumber(PagerNumber);
            clientPage.SetFaxNumber(FaxNumber);

            clientPage.ClickSaveButton();
            string clientId = clientPage.GetClientID();

            clientPage.ClickSaveCloseButton();
            
            //Clients Services > Clients Double click on the prepared record
            clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.SetClientSearchText(clientId);
            table = new Table(clientsSearchPage.GetSearchResultTable());
            table.ClickCellValue("RTA Client Id", clientId, "Full Name");

            clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();
            Thread.Sleep(500);
            string date = DateTime.Today.ToString("d/MM/yyyy");
            string time = DateTime.Now.ToString("h:mm tt");
            /*Clear the following fields and save record:
            - Personal Mobile
            - Personal Preferred Contact Number
            - Work Preferred Contact Number
            - Fax*/

            clientPage.ClearMobileNumber();
            clientPage.ClickPageTitle();
            clientPage.ClearPersonalPreferredMobileNumber();
            clientPage.ClickPageTitle();
            clientPage.ClearWorkPreferredNumber();
            clientPage.ClickPageTitle();
            clientPage.ClearFaxNumber();

            clientPage.ClickSaveCloseButton();

            clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.SetClientSearchText(clientId);
            table = new Table(clientsSearchPage.GetSearchResultTable());
            table.ClickCellValue("RTA Client Id", clientId, "Full Name");
            clientPage = new ClientPage(driver);
            clientPage.ClickPageTitle();

            //Inspect resultant entry in Activity sub grid.
            Table headerTable = new Table(clientPage.GetActivitiesHeaderTable());
            headerTable.ClickTableColumnHeader("Actual End");
            Thread.Sleep(2000);

            Table activitiesTable = new Table(clientPage.GetActivitiesTable());
            Assert.IsTrue(activitiesTable.MatchingCellFound("Subject", "Client Profile Phone Number updated"), "Client Phone Number Updation Activity NOT displayed !!!!!");

            activitiesTable.ClickCellValue("Subject", "Client Profile Phone Number updated", "Subject");

            //Client Management Activity is created as a note to users that the phone numbers were changed (includes current value of each field and a note referring the user to audit history).
            ClientManagementActivityPage activityPage = new ClientManagementActivityPage(driver);
            activityPage.ClickPageTitle();
            StringAssert.Contains(activityPage.GetDescription(), "Personal Mobile:   [no value]");
            StringAssert.Contains(activityPage.GetDescription(), "Personal Preferred Contact Number:   [no value]");
            StringAssert.Contains(activityPage.GetDescription(),"Work Preferred Contact Number:   [no value]");
            StringAssert.Contains(activityPage.GetDescription(), "Fax:   [no value]");
            StringAssert.Contains(activityPage.GetDescription(), "Note: the current values are displayed above. Refer to Audit History against the Client record for more information.");

            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();                      

            clientsSearchPage = new ClientsSearchPage(driver);
            clientsSearchPage.SetPageFilterList("Active Clients");
        
            clientsSearchPage.SetClientSearchText(clientId);

            Table clientTable = new Table(clientsSearchPage.GetSearchResultTable());
            IWebElement clientCell = clientTable.GetCellElementContainsValue("RTA Client Id", clientId, "Full Name");
            UICommon.DoubleClickElement(clientCell, driver);

            Thread.Sleep(5000);
            clientPage = new ClientPage(driver, "Clients Quick");

            /*Select the [Audit History] option from the entity navigation menu in the global ribbon.
             *Inspect audit history relating to deletions made in this test*/

            homePage.HoverClientRibbonTab(clientName);
            homePage.ClickClientXAuditRibbonButton();

            clientPage = new ClientPage(driver, "Clients Quick");
            Thread.Sleep(2000);

            /*Audit history shows:
            - What was deleted
            - Who deleted it
            - When it was deleted
            - What previous value was*/
            Table auditTable = new Table(clientPage.GetAuditHistoryTable());

            StringAssert.Contains(auditTable.GetCellValue("Event", "Update", "Changed Field"), "Fax");
            StringAssert.Contains(auditTable.GetCellValue("Event", "Update", "Changed Field"), "Mobile");
            StringAssert.Contains(auditTable.GetCellValue("Event", "Update", "Changed Field"), "Personal Preferred Contact Number");
            StringAssert.Contains(auditTable.GetCellValue("Event", "Update", "Changed Field"), "Work Preferred Contact Number");
            
            StringAssert.Contains(auditTable.GetCellValue("Event", "Update", "Old Value"), "+61 7 3333 3333");
            StringAssert.Contains(auditTable.GetCellValue("Event", "Update", "Old Value"), "+61 422 222 222");
            StringAssert.Contains(auditTable.GetCellValue("Event", "Update", "Old Value"), "+61 7 3333 3333");
            StringAssert.Contains(auditTable.GetCellValue("Event", "Update", "Old Value"), "+61 7 1111 1111");
            StringAssert.Contains(auditTable.GetCellValue("Event", "Update", "Old Value"), "+61 7 4444 4444");

            Assert.AreEqual(String.Empty, auditTable.GetCellValue("Event", "Update", "New Value").Trim());
            Assert.AreEqual(user.Id.ToLower() + " user", auditTable.GetCellValue("Event", "Update", "Changed By").ToLower());
            Assert.AreEqual(date + " " + time, auditTable.GetCellValue("Event", "Update", "Changed Date"));

        }

        [TestMethod]
        [TestProperty("TestcaseID", "WildcardSearch")]
        public void ATC_CRMEntitySearchClientWildcard()
        {

            #region Start Up Excel
            MyBook = MyApp.Workbooks.Open(DatasourceDir + @"\Clients.xlsx", 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
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

            string clientName = (MyRange.Cells[MyRow, ClientsSchema.GetColumnIndex("CLIENT_NAME")].Value.ToString());
            
            //Login in as role
            User user = this.environment.GetUser(SecurityRole.Investigations);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickClientServicesRibbonButton();
            homePage.HoverClientServicesRibbonTab();
            homePage.ClickClientsRibbonButton();



            //Search for already existing client
            ClientsSearchPage clientsSearchPage = new ClientsSearchPage(driver);
            Trace.Listeners.Add(new TextWriterTraceListener("TextWriterOutput.log", "myListener"));
            Trace.TraceInformation("Search for client start:" + DateTime.Now.ToString("ddMMyyyyhhmmssffff"));
            Trace.Flush();
            clientsSearchPage.SetClientSearchText("*"+clientName);

            Table table = new Table(clientsSearchPage.GetSearchResultTable());
            StringAssert.Equals(table.GetCellValue("Full Name", clientName, "Full Name"), clientName);

            Trace.Listeners.Add(new TextWriterTraceListener("TextWriterOutput.log", "myListener"));
            Trace.TraceInformation("Search for client end:" + DateTime.Now.ToString("ddMMyyyyhhmmssffff"));
            Trace.Flush();

            table.ClickCellValue("Full Name", clientName, "Full Name");

            ClientPage clientPage = new ClientPage(driver);
            String BaseWindow = driver.CurrentWindowHandle;

            driver = driver.SwitchTo().Window(BaseWindow);
            clientPage.ClickSaveButton();

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }
    }
}
