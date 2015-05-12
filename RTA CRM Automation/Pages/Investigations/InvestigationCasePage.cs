using ActionWordsLib.Attributes;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using RTA.Automation.CRM.Utils;
using RTA.Automation.CRM.UI;
using RTA.Automation.CRM.Pages.Investigations;


namespace RTA.Automation.CRM.Pages
{
    [ActionClass]
    public class InvestigationCasePage : IFramePage
    {
        private static string frameId = "contentIFrame0";
        private static string dialogFRAME = "InlineDialog_Iframe";
        private static string frameID2 = "contentIFrame1";
        private static string contentFrameId = "areaActivitiesFrame";
        private static string pageTitle = "Investigation Case:";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;

        public InvestigationCasePage(IWebDriver driver)
            : base(driver, InvestigationCasePage.frameId)
        {
            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

            //Switch to main frame when it is visible
            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);

        }

        [ScenarioMethod]
        public void PopulateNewInvestigationCase()
        {
            this.ClickSaveButton();
        }


        /*
        * NewIMG
        * ************************************************************************
        */

        [ActionMethod]
        public void ClickNewInvestigationCaseButton()
        {
            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickNewButton(driver);
           
        }

        /*
        * SaveIMG
        * ************************************************************************
        */

        [ActionMethod]
        public void ClickSaveButton()
        {

            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickSaveButton(driver);
            this.driver.SwitchTo().Frame(frameId);
            Thread.Sleep(2000);
       }


        [ActionMethod]
        public void ClickSaveCloseButton()
        {

            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickSaveCloseButton(driver);
            

        }

        [ActionMethod]
        public void ClickCRMToolbar()
        {
            this.driver.SwitchTo().DefaultContent();

            try
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
                IWebElement Parent = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("dxtools_QuickView_Area")));
                UICommon.ClickRibbonTab("#Tab1", driver);
                this.driver.SwitchTo().Frame(frameId);
            }
            catch
            {
                this.driver.SwitchTo().Frame(frameId);
            }



        }

        [ActionMethod]
        public void ClickStartDialogButton()
        {

            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickStartDialogButton(driver);
            this.driver.SwitchTo().Frame(frameId);
        }

        [ActionMethod]
        public void ClickDialogAddButton()
        {
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(dialogFRAME);
            UICommon.ClickDialogAddButton(driver);
            Thread.Sleep(3000);
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(frameId);
        }


        /*
        * InvestigationCase Number
        * ************************************************************************
        */

        [ActionMethod]
        public string GetInvestigationCaseNumber()
        {

            return UICommon.GetNewReferenceNumber(driver);
        }

        /*
        * Investigator Field
        * ************************************************************************
        */


        [ActionMethod]
        public void ClickInvestigatorSearchButton()
        {
            
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//div/div/img[contains(@src,'transparent_spacer.gif')]"))).Click();


            //Wait for list to be displayed
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Dialog_header_rta_investigatorid_i_IMenu"))).GetAttribute("id");

            IWebElement elem = driver.FindElement(By.Id("header_rta_investigatorid_i"));
            elem.SendKeys(Keys.Backspace);

            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_dwelling_typeid_ledit"))).GetAttribute("id");



        }

        [ActionMethod]
        public void ClickInvestigatorSearchElement(string investigator)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//li[a[contains(@title,'" + investigator + "')] and parent::*[@id='header_rta_investigatorid_i_IMenu']]"))).Click();
        }

        [ActionMethod]
        public void ClickInvestigatorSearchText()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#header_rta_investigatorid>div")));
            Thread.Sleep(2000);
            Actions action = new Actions(driver);
            action.MoveToElement(elem, 1, 1).ClickAndHold().Build().Perform();
            Thread.Sleep(2000);
            action.MoveToElement(elem).Release().Build().Perform();
        
            //Switch to main frame when it is visible
            driver.SwitchTo().DefaultContent();
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement f1 = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(frameId)));
            driver.SwitchTo().Frame(f1);

            wait.Until(ExpectedConditions.ElementExists(By.Id("header_rta_investigatorid_ledit")));

        }

        [ActionMethod]
        public void SetInvestigatorSearchElement(string investigator)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementExists(By.Id("header_rta_investigatorid_ledit")));
            elem.SendKeys(investigator);
            elem.SendKeys(Keys.Enter);
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Dialog_header_rta_investigatorid_i_IMenu"))).GetAttribute("id");


        }
        [ActionMethod]
        public bool GetInvestigatorSearchElementText(string investigator)
        {
            ClickInvestigatorSearchText();
            //ClickInvestigatorSearchButton();
            SetInvestigatorSearchElement(investigator);
            ClickInvestigatorSearchElement(investigator);
            return true;

        }

        [ActionMethod]
        public void ClickActivitiesAddButton()
        {
            UICommon.ClickAddButton(driver, "Activities_addImageButtonImage");
            
        }

        [ActionMethod]
        public void ClickAddTaskButton(string activity)
        {

            UICommon.ClickAddActivity(activity, driver); 
           

        }


        [ActionMethod]
        public void ClickAddActivity(string Activity)
        {
            UICommon.ClickAddActivity(Activity, driver);
        }
        
        [ActionMethod]
        public void ClickSeeRecordsAssociatedWithThisViewButton(string association)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(association + "_openAssociatedGridViewImageButtonImage")));
            Actions action = new Actions(driver);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).Release().Build().Perform();
            Thread.Sleep(2000);

//            UICommon.ClickSeeRecordsAssociatedWithThisViewButton(association, driver);
        }

        [ActionMethod]
        public void AddNewTaskActivityRecord()
        {
            RefreshPageFrame.RefreshPage(driver, frameId); WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#>a")));
            elem.Click();
        }


        /*
        * Alleged Offences Tab
        * ************************************************************************
        */

        [ActionMethod]
        public void ClickAllegedOffencesTab()
        {
            RefreshPageFrame.RefreshPage(driver, frameId);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//div/a/h2[contains(text(),'Alleged Offences')]")));
            string text = elem.GetAttribute("text");
            Actions action = new Actions(driver);
            Thread.Sleep(2000);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(2000);
            action.MoveToElement(elem).Release().Build().Perform(); 
        }
        

        [ActionMethod]
        public void ClickAllegedOffencesAddButton()
        {
            RefreshPageFrame.RefreshPage(driver, frameId); WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("AllegedOff_addImageButtonImage")));
            //**************************************Use this for highlighting elements**************************************
            var jsDriver = (IJavaScriptExecutor)driver;
            var element = elem;
            string highlightJavascript = @"$(arguments[0]).css({ ""border-width"" : ""2px"", ""border-style"" : ""solid"", ""border-color"" : ""red"" });";
            jsDriver.ExecuteScript(highlightJavascript, new object[] { element });

            Actions action = new Actions(driver);
            Thread.Sleep(2000);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(2000);
            action.MoveToElement(elem).Release().Build().Perform(); 

        }

        [ActionMethod]
        public bool CheckAddNewActivityContents(String activity)
        {
            IWebElement newActivity = driver.FindElement(By.Id("moreActivitiesList"));
            IList<IWebElement> activityList = newActivity.FindElements(By.TagName("li"));
            foreach (IWebElement item in activityList)
            {
                string title = item.GetAttribute("Title");
                title = title.Replace("Add ", "");
                if (activity.Equals(title))
                {
                    return true;
                }
            }
            return false;

        }

        /*
      * search table 
      * ************************************************************************
      */

        [ActionMethod]
        public IWebElement GetActivitiesSearchResultTable()
        {
           
            WaitForPageToLoad.WaitToLoad(driver);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement Parent = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Activities_divDataArea")));
            IWebElement webElementBody = Parent.FindElement(By.Id("gridBodyTable"));
            return webElementBody;
        }

        [ActionMethod]
        public IWebElement GetActivitiesHeaderTable()
        {
            WaitForPageToLoad.WaitToLoad(driver);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));

            IWebElement Parent = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Activities_gridBar")));
            // IWebElement webElementBody = Parent.FindElement(By.Id("Activities_gridBar"));
            return Parent;
        }

        [ActionMethod]
        public IWebElement GetProcessSearchResultTable()
        {

            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(dialogFRAME);
            return UICommon.GetSearchResultTable(driver);
        }


        internal IWebDriver SwitchNewBrowser(IWebDriver driver, string BaseWindow)
        {
            return UICommon.SwitchToNewBrowser(driver, BaseWindow);
        }
        internal IWebDriver SwitchNewBrowserWithTitle(IWebDriver driver, string BaseWindow, string NewBrowserTitle)
        {
            return UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, NewBrowserTitle);
        }

          /*
          * Status and Sub-status
          * ************************************************************************
          */
        [ActionMethod]

        public string GetStatus()
        {
            return driver.FindElement(By.CssSelector("#rta_inv_statusid > div > span")).GetAttribute("title");
        }

        public string GetSubStatus()
        {
            return driver.FindElement(By.CssSelector("#rta_inv_sub_statusid > div > span")).GetAttribute("title");
        }

        public void SetStatus(string status)
        {
            this.OpenLookUpRecordWindow("rta_inv_statusid");

            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(dialogFRAME);
            Table table = new Table(this.GetStatusTable());
            table.SelectTableRow("Name", status);

            string BaseWindow = driver.CurrentWindowHandle; 

            if (driver.WindowHandles.Count > 1)
            {
                driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "Investigation Case Status");
                
                InvestigationCaseStatusPage investigationStatus = new InvestigationCaseStatusPage(driver);
                investigationStatus.CloseWindow();

                driver = driver.SwitchTo().Window(BaseWindow);
            }
        }

        public void SetSubStatus(string subStatus)
        {
            this.OpenLookUpRecordWindow("rta_inv_sub_statusid");

            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(dialogFRAME);
            Table table = new Table(this.GetStatusTable());
            table.SelectTableRow("Name", subStatus);

            string BaseWindow = driver.CurrentWindowHandle; 

            if (driver.WindowHandles.Count > 1)
            {
                driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "Investigation Case Status");
                
                InvestigationCaseStatusPage investigationStatus = new InvestigationCaseStatusPage(driver);
                investigationStatus.CloseWindow();             

                driver = driver.SwitchTo().Window(BaseWindow);
            }

        }

        public int GetSubStatusCount()
        {

            List<IWebElement> subStatusList = driver.FindElements(By.CssSelector("#rta_inv_sub_statusid_i_IMenu>li")).ToList();

            return subStatusList.Count - 1; // 1 item is "Lookup for more records"
        }

        [ActionMethod]
        public void ClickSubStatusSearchButton()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#rta_inv_sub_statusid>div"))).Click();

            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_inv_sub_statusid_i"))).Click();

        }

        [ActionMethod]
        public void ClickSubStatusText()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#rta_inv_sub_statusid>div"))).Click();

        }

        public bool FindSubStatusFromDropdown(string subStatus)
        {
            List<IWebElement> subStatusList = driver.FindElements(By.CssSelector("#rta_inv_sub_statusid_i_IMenu>li")).ToList();

            foreach (IWebElement status in subStatusList)
            {
                if (status.Text.Contains(subStatus))
                {
                    return true;
                }
            }
            throw new Exception("No sub status found");
        }

        public string GetSubStatusErrorText()
        {
            return UICommon.GetTextFromElement("#rta_inv_sub_statusid_err", driver);
        }

        public void OpenLookUpRecordWindow(string CssSelector)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#" + CssSelector + ">div"))).Click();

            wait.Until(ExpectedConditions.ElementIsVisible(By.Id(CssSelector + "_i"))).Click();

            wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("ms-crm-IL-MenuItem-Title-Lookupmore"))).Click();
        }

        [ActionMethod]
        public IWebElement GetStatusTable()
        {
            return UICommon.GetSearchResultTable(driver);
        }

        [ActionMethod]
        public void ClickPageTitle()
        {
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(frameId);
            UICommon.ClickPageTitle(driver);
        }

        [ActionMethod]
        public void SetPageFilterList(string value)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#Activities_SavedNewQuerySelector>span"))).Click();
            IWebElement parent = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Dialog_0")));
            parent.FindElement(By.XPath("//li[a[contains(@title,'" + value + "')]]")).Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#Activities_SavedNewQuerySelector>span")));
        }
          /*
          *     Owner
          * ************************************************************************
          */
        [ActionMethod]
        public void SetOwnerValue(string value)
        {
            UICommon.SetSearchableListValue("header_ownerid", value, driver);

        }

        [ActionMethod]
        public string GetOwnerNameValue()
        {
            return UICommon.GetTextFromElement("#header_ownerid>div>span", driver);

        }
        [ActionMethod]
        public string GetOwnerValidationMessageString(string cssSelector)
        {
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(dialogFRAME);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(cssSelector)));
            string HeaderMessage = elem.Text;
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(frameId);

            return HeaderMessage;
        }
        /*
         *     Received Date
         * ************************************************************************
         */
        [ActionMethod]
        public void SetReceivedDateValue(string receivedDateValue)
        {
            this.ClickPageTitle();
            UICommon.SetDateValue("rta_received_date", receivedDateValue, driver);
        }

        /*
        *     Duration (Days)
        * ************************************************************************
        */
        [ActionMethod]
        public int GetDurationDaysValue()
        {
            this.ClickPageTitle();
            string durationValue = UICommon.GetTextFromElement("#rta_duration>div>span", driver);
            if (durationValue == "--")
            {
                return -1;
            }
            else
            {
                return Convert.ToInt32(durationValue);
            }
        }
        /*
        *     Verify Action Date and Follow Up date are removed
        * ************************************************************************
        */
        [ActionMethod]
        public bool VerifyElementExists(string CssSelector)
        {
            return UICommon.CheckElementExists(CssSelector, driver);
        }


        /*
        *     Queue
        * ************************************************************************
        */

        [ActionMethod]
        public void ClickAddToQueueButton()
        {
            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickAddToQueueButton(driver);
            this.driver.SwitchTo().Frame(frameId);
        }
        [ActionMethod]
        public void SetQueue(string QueueName)
        {
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(dialogFRAME);
            UICommon.SetQueue(QueueName, driver);
        }


        /*
        *     Run Workflow
        * ************************************************************************
        */
        [ActionMethod]
        public void ClickRunWorkflowButton()
        {

            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickImageButton(driver, "...");
            UICommon.ClickImageButton(driver, "Run Workflow");
            this.driver.SwitchTo().Frame(frameId);
        }
        [ActionMethod]
        public void ClickConfirmApplicationOfWindow(string BaseWindow)
        {
            driver = UICommon.SwitchToNewBrowser(driver, BaseWindow);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("butBegin")));
            elem.Click();            
        }

        /*
        *     Activities Associated View Table
        * ************************************************************************
        */
        [ActionMethod]
        public IWebElement GetActivitiesAssociatedViewTable()
        {
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(frameID2);
            driver.SwitchTo().Frame(contentFrameId);
            return UICommon.GetSearchResultTable(driver);
        }

        [ActionMethod]
        public void SetActivitiesSearchText(string SearchText)
        {
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(frameID2);
            driver.SwitchTo().Frame(contentFrameId);
            UICommon.SetSearchText("crmGrid_findCriteria", SearchText, driver);
        }


        /*
        *     Case Party Table
        * ************************************************************************
        */

        [ActionMethod]
        public void ClickAddCasePartyRecordButton()
        {
            UICommon.ClickAddButton(driver, "CaseParties_addImageButtonImage");
        }

        [ActionMethod]
        public IWebElement GetCasePartyTable()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement webElementBody = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("table[ologicalname='rta_case_party']")));
            return webElementBody;
        }

        /*
        *     Case Party Associated View Table
        * ************************************************************************
        */

        [ActionMethod]
        public IWebElement GetCasePartyAssociatedTable()
        {
            driver.SwitchTo().Frame("rta_inv_case_rta_case_partyFrame");
            return UICommon.GetSearchResultTable(driver);
        }

        [ActionMethod]
        public void ClickNewActivityButton()
        {
            //driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(contentFrameId);
            UICommon.ClickElementWithId("activitypointer|NoRelationship|SubGridAssociated|Mscrm.SubGrid.activitypointer.NewRecord", driver);
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(frameId);
        }



        public void ClickSaveFooter()
        {
            UICommon.ClickSaveFooter(driver);
        }

        public void CheckForErrors()
        {
            try
            {
                driver.SwitchTo().Frame(dialogFRAME);
                UICommon.ClickElementWithId("butBegin", driver);
                driver.SwitchTo().DefaultContent();
                driver.SwitchTo().Frame(frameId);
            }
            catch
            { }
            
        }
    }
}
