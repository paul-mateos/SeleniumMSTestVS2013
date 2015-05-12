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


namespace RTA.Automation.CRM.Pages
{
    [ActionClass]
    public class InvestigationMasterCasePage : IFramePage
    {
        private static string frameId = "contentIFrame0";
        private static string pageTitle = "Investigation Master Case:";
        public static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;
        

        public InvestigationMasterCasePage(IWebDriver driver)
            : base(driver, InvestigationMasterCasePage.frameId)
        {
            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

            ////Switch to main frame when it is visible
            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);

            //click on title to move focus
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("FormTitle"))).Click();

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
        public void ClickSaveCloseButton()
        {

            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickSaveCloseButton(driver);
            this.driver.SwitchTo().Frame(frameId);


        }

        [ActionMethod]
        public IWebElement GetInvestigationCasesSearchResultTable()
        {

            WaitForPageToLoad.WaitToLoad(driver);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement Parent = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Subgrid_InvCases")));
            IWebElement webElementBody = Parent.FindElement(By.Id("gridBodyTable"));
            return webElementBody;
        }


        /*
        * Investigation Master Case Number
        * ************************************************************************
        */

        [ActionMethod]
        public string GetInvestigationMasterCaseNumber()
        {
            return UICommon.GetNewReferenceNumber(driver);
        }

        /*
        * Activities  Tab
        * ************************************************************************
        */

        [ActionMethod]
        public void ClickActivitiesAddButton()
        {
           
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("FormTitle"))).Click();
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Subgrid_Activities_addImageButtonImage")));
            Actions action = new Actions(driver);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).Release().Build().Perform();
           
        }

        [ActionMethod]
        public void ClickInvestigationCaseAddButton()
        {

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("FormTitle"))).Click();
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Subgrid_InvCases_addImageButtonImage")));
            elem.Click();
        }

        [ActionMethod]
        public void ClickInvestigationCaseAssociatedView()
        {

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("FormTitle"))).Click();
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Subgrid_InvCases_openAssociatedGridViewImageButtonImage")));
            elem.Click();
        }

        [ActionMethod]
        public String GetCurrentView()
        {
            RefreshPageFrame.RefreshPage(driver, frameId);
            driver.SwitchTo().Frame("rta_inv_master_case_rta_inv_case_investigation_master_caseidFrame");
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement element = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#crmGrid_rta_inv_master_case_rta_inv_case_investigation_master_caseid_SavedNewQuerySelector>span")));
            return element.Text;
        }

        [ActionMethod]
        public void ClickPageTitle()
        {
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(frameId);
            UICommon.ClickPageTitle(driver);
        }
              
        [ActionMethod]
        public void SetInvestigationCaseNumberToAssociateMaster(string listValue)
        {

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("lookup_Subgrid_InvCases_ledit")));
            Actions actions = new Actions(driver);
            actions.MoveToElement(elem).Build().Perform();
            Thread.Sleep(500);
            elem.Click();
            elem.SendKeys(Keys.Backspace);
            elem.SendKeys(Keys.Clear);
            elem.SendKeys(listValue);
            elem.SendKeys(Keys.Enter);
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//li[a[contains(@title,'" + listValue + "')]]"))).Click();
        }


         [ActionMethod]
        public void ClickAddTaskButton(string activity)
        {
            UICommon.ClickAddActivity(activity, driver);

            //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            //IWebElement parent = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("moreActivitiesList")));
            //IWebElement elem = parent.FindElement(By.CssSelector("#AddtaskButton>span>a>img[alt='Task']"));
            //Actions action = new Actions(driver);

            //action.MoveToElement(elem).Build().Perform();
            //Thread.Sleep(1000);
            //action.MoveToElement(elem).ClickAndHold().Build().Perform();
            //Thread.Sleep(3000);
            //action.MoveToElement(elem).Release().Build().Perform();
            //Thread.Sleep(3000);
        }

        [ActionMethod]
        public void SetUnknowInvestigatorValue(string value)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));

            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#header_rta_investigatorid>div"))).Click();
            IWebElement parent = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("header_rta_investigatorid_lookupTable")));
            IWebElement elem = parent.FindElement(By.Id("header_rta_investigatorid_ledit"));
            elem.Clear();
            elem.SendKeys(value);
            elem.SendKeys(Keys.Enter);
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//li[a[contains(@title,'No records found. Create a new record.')] and parent::*[@id='header_rta_investigatorid_i_IMenu']]"))).Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("FormTitle"))).Click();
   
        }

        [ActionMethod]
        public void SetInvestigatorValue(string value)
        {
            UICommon.SetSearchableListValue("header_rta_investigatorid", value, driver);

        }

        [ActionMethod]
        public void SetClientValue(string value)
        {
            UICommon.SetSearchableListValue("rta_clientid", value.ToUpper(), driver);

        }
        
        [ActionMethod]
        public void SetOwnerValue(string value)
        {
            UICommon.SetSearchableListValue("header_ownerid_ledit", value, driver);

        }



        internal IWebDriver SwitchNewBrowserWithTitle(IWebDriver driver, string BaseWindow, string title)
        {
            return UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, title);
        }

        public void SwitchToMasterCasePageFrame()
        {
            RefreshPageFrame.RefreshPage(driver, frameId);
            driver.SwitchTo().Frame("rta_inv_master_case_rta_inv_case_investigation_master_caseidFrame");
        }

        [ActionMethod]
        public IWebElement GetInvestigationCasesAssociatedViewTable()
        {

            WaitForPageToLoad.WaitToLoad(driver);
            return UICommon.GetSearchResultTable(driver);
        }
    }
}
