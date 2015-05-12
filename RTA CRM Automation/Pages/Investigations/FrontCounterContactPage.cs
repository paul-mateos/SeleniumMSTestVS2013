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
    public class FrontCounterContactPage : IFramePage
    {
        private static string frameId = "contentIFrame0";        
        private static string pageTitle = "Front Counter Contact:";
        private static string dialogFRAME = "InlineDialog_Iframe";
        private static string connectionsFRAME = "areaConnectionsFrame";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;

        public FrontCounterContactPage(IWebDriver driver)
            : base(driver, FrontCounterContactPage.frameId)
        {
            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

            //Switch to main frame when it is visible
            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);
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
        public void ClickSaveCloseButton()
        {
            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickSaveCloseButton(driver);
        }

        [ActionMethod]
        public void ClickCloseFrontCounterContactButton()
        {
            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickImageButton(driver, "Close Front Counter Contact");
        }

        [ActionMethod]
        public void ClickStartDialogButton()
        {
            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickStartDialogButton(driver);
        }

        [ActionMethod]
        public void SetSubjectValue(string subject)
        {
            //Switch to main frame when it is visible
            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);
            UICommon.SetTextBoxValue("subject", subject, driver);
        }
       
        [ActionMethod]
        public void ClickPageTitle()
        {
            UICommon.ClickPageTitle(driver);            
        }

        [ActionMethod]
        public void SetClientName(string ClientName)
        {
            UICommon.SetSearchableMultiListValue("customers", ClientName, driver); 
        }

        [ActionMethod]
        public string GetClientName()
        {
            return UICommon.GetTextFromElement("#customers>div>span", driver);
        }

        [ActionMethod]
        public void SetActualEndDate(string Date)
        {
            UICommon.SetDateValue("actualend", Date, driver);
        }

        [ActionMethod]
        public void SetActualDuration(string Duration)
        {
             // UICommon.SetSelectListValue("actualdurationminutes", Duration, driver);
             WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
             wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#actualdurationminutes>div"))).Click();
             wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#actualdurationminutes_iSelectInput"))).SendKeys(Duration);
             wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#actualdurationminutes_iSelectInput"))).SendKeys(Keys.Enter);
        }

        [ActionMethod]
        public void SetAssistiveService(string Service)
        {
            UICommon.SetSearchableListValue("rta_assistive_service_usedid", Service, driver);
        }

        [ActionMethod]
        public void SetRegardingClientValue(string ClientName)
        {
            UICommon.SetSearchableListValue("regardingobjectid", ClientName, driver);
        }

        [ActionMethod]
        public void SetSelectSubjectValue(string SelectSubjectValue)
        {
            UICommon.SetSearchableListValue("rta_activity_subjectid", SelectSubjectValue, driver);
        }

        [ActionMethod]
        public void SetDescription(string description)
        {
            UICommon.SetTextBoxValue("description", description, driver);
        }

        [ActionMethod]
        public string GetDescription()
        {
            return UICommon.GetTextFromElement("#description>div>span", driver);
        }

        /*
        *     Queue
        * ************************************************************************
        */

        [ActionMethod]
        public void ClickAddToQueueButton()
        {
            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickImageButton(driver,"...");
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

        [ActionMethod]
        public void ClickDialogAddButton()
        {
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(dialogFRAME);
            UICommon.ClickDialogAddButton(driver);
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(frameId);
        }

        [ActionMethod]
        public void SetConnectList(string connectType)
        {

            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(frameId);
            driver.SwitchTo().Frame(connectionsFRAME);

            UICommon.SetConnectList(connectType, driver);
        }

        public IWebDriver SwitchNewBrowser(IWebDriver driver, string BaseWindow, string newWindow)
        {
            return UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, newWindow);
        }

        /*
        * search table 
        * ************************************************************************
        */

        [ActionMethod]
        public IWebElement GetConnectionsTable()
        {
            this.SwitchFrame();
            return UICommon.GetSearchResultTable(driver);
        }

        [ActionMethod]
        public void SwitchFrame()
        {
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(frameId);
            driver.SwitchTo().Frame(connectionsFRAME);        
        }

        [ActionMethod]
        public IWebElement GetProcessSearchResultTable()
        {

            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(dialogFRAME);
            return UICommon.GetSearchResultTable(driver);
        }

    }
}

