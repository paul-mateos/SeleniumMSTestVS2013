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

namespace RTA.Automation.CRM.Pages.Investigations
{
    [ActionClass]
    public class InvestigationGeneralCasePage : IFramePage
    {
        private static string frameId = "contentIFrame0";
        //private static string dialogFRAME = "InlineDialog_Iframe";
        private static string pageTitle = "General Case:";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;

        public InvestigationGeneralCasePage(IWebDriver driver)
            : base(driver, InvestigationGeneralCasePage.frameId)
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
            Thread.Sleep(2000);
        }

        [ActionMethod]
        public void ClickSaveCloseButton()
        {
            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickSaveCloseButton(driver);
        }

        /*
        * General Case Number
        * ************************************************************************
        */

        [ActionMethod]
        public string GetGeneralCaseNumber()
        {
            return UICommon.GetNewReferenceNumber(driver);
        }


        [ActionMethod]
        public string GetReceivedDateErrorMessage()
        {
            return UICommon.GetTextFromElement("#rta_received_date_err", driver);
        }
        

        /*
        * General Case Title
        * ************************************************************************
        */
        [ActionMethod]
        public void SetTitle(string caseTitle)
        {
            UICommon.SetTextBoxValue("title", caseTitle, driver);
        }

        /*
        * Client Name
        * ************************************************************************
        */
        [ActionMethod]
        public void SetClientName(string clientName)
        {
            UICommon.SetSearchableListValue("customerid", clientName, driver);
        }

        /*
        * General Case Type
        * ************************************************************************
        */
        [ActionMethod]
        public void SetType(string caseType)
        {
            UICommon.SetSearchableListValue("rta_case_typeid", caseType, driver);
        }
        
        /*
        * Owner Field
        * ************************************************************************
        */
        [ActionMethod]
        public void SetInvestigatorSearchElementText(string investigator)
        {
            UICommon.SetSearchableListValue("header_ownerid", investigator, driver);
        }

        [ActionMethod]
        public void ClickPageTitle()
        {
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(frameId);
            UICommon.ClickPageTitle(driver);
        }

        /*
         * Activities Table
         * ************************************************************************
         */


        [ActionMethod]
        public void ClickActivitiesAddButton()
        {
            UICommon.ClickAddButton(driver, "Activities_addImageButtonImage");
        }

        [ActionMethod]
        public void ClickAddActivity(string ActivitycssSelectorId)
        {
            UICommon.ClickAddActivity(ActivitycssSelectorId, driver);
        }

        [ActionMethod]
        public IWebElement GetActivitiesTable()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement webElementBody = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("table[ologicalname='activitypointer']")));
            return webElementBody;
        }

        internal IWebDriver SwitchNewBrowser(IWebDriver driver, string BaseWindow)
        {
            return UICommon.SwitchToNewBrowser(driver, BaseWindow);
        }

        internal void SetReceivedDate(string tomorrowDate)
        {
            UICommon.SetDateValue("rta_received_date", tomorrowDate, driver);
        }

        internal void ClickSaveFooter()
        {
            UICommon.ClickSaveFooter(driver);
        }

        public IWebDriver SwitchNewBrowserWithTitle(IWebDriver driver, string BaseWindow, string title)
        {
            return UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, title);
        }
    }
}
