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
    public class PhoneCallPage : IFramePage
    {
        private static string frameId = "contentIFrame0";
        private static string frameId2 = "contentIFrame1";
        private static string dialogFRAME = "InlineDialog_Iframe";
        private static string pageTitle = "Phone Call";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;

        public PhoneCallPage(IWebDriver driver)
            : base(driver, PhoneCallPage.frameId)
        {
            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

            //Switch to main frame when it is visible
            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);

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
        }

        [ActionMethod]
        public void ClickSaveCloseButton()
        {

            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickSaveCloseButton(driver);

        }

        [ActionMethod]
        public void ClickStartDialogButton()
        {

            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickStartDialogButton(driver);
            this.driver.SwitchTo().Frame(frameId);
        }

        [ActionMethod]
        public void ClickPageTitle()
        {
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(frameId);
            UICommon.ClickPageTitle(driver);
        }

        [ActionMethod]
        public string GetPageTitle()
        {
            return UICommon.GetNewReferenceNumber(driver);
        }

        [ActionMethod]
        public void ClickClosePhoneCallButton()
        {

            this.driver.SwitchTo().DefaultContent();
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[alt='Close Phone Call']")));
            Actions action = new Actions(driver);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            action.MoveToElement(elem).Release().Build().Perform();
            Thread.Sleep(3000);
            this.driver.SwitchTo().Frame(frameId);
        }

        [ActionMethod]
        public void ConfirmDeactivation(String listValue)
        {
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(dialogFRAME);
            String elementId = "statusCode";
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#" + elementId))).Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//select[@id='" + elementId + "']/optgroup/option[text()='" + listValue + "']"))).Click();
            Thread.Sleep(1000);
        }

        [ActionMethod]
        public void ClickConfirmDeactivationCloseButton()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("butBegin"))).Click();
        }

        [ActionMethod]
        public IWebElement GetProcessSearchResultTable()
        {

            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(dialogFRAME);
            return UICommon.GetSearchResultTable(driver);
        }

        [ActionMethod]
        public void ClickDialogAddButton()
        {

            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(dialogFRAME);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("butBegin"))).Click();
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(frameId);
        }

        [ActionMethod]
        public void ClickMarkCompleteButton()
        {

            this.driver.SwitchTo().DefaultContent();

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[alt='Mark Complete']")));
            Actions action = new Actions(driver);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).Release().Build().Perform();
            Thread.Sleep(3000);
        }

        /*
        * Task Number
        * ************************************************************************
        */

            
        [ActionMethod]
        public void SetSelectSubjectValue(string subject)
        {

            //Switch to main frame when it is visible
            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);
            UICommon.SetSearchableListValue("rta_activity_subjectid", subject, driver);

        }

        [ActionMethod]
        public void SetSubject(string subject)
        {

            //Switch to main frame when it is visible
            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);
            UICommon.SetTextBoxValue("subject", subject, driver);

        }

        [ActionMethod]
        public void SetRecipient(string receipient)
        {

            //Switch to main frame when it is visible
            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);
            UICommon.SetSearchableMultiListValue("to", receipient, driver);

        }

        public string GetSubjectValue()
        {
            return UICommon.GetTextFromElement("#subject>div>span", driver);
        }

        public string GetSenderValue()
        {
            return UICommon.GetTextFromElement("#from>div>span", driver);
        }
        public string GetSelectSubjectValue()
        {
            return UICommon.GetTextFromElement("#rta_activity_subjectid>div>span", driver);
        }
        public string GetRecipientValue()
        {
            return UICommon.GetTextFromElement("#to>div>span", driver);
        }
        public void SwitchFrame()
        {
            this.driver.SwitchTo().DefaultContent();
            this.driver.SwitchTo().Frame(frameId2);        
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
