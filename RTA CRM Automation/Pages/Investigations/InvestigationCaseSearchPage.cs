using ActionWordsLib.Attributes;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RTA.Automation.CRM.Utils;
using System.Threading;
using RTA.Automation.CRM.UI;

namespace RTA.Automation.CRM.Pages
{
    [ActionClass]
    public class InvestigationCaseSearchPage : IFramePage
    {
        private static string frameId = "contentIFrame0";
        private static string pageTitle = "Investigation Case";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;

        public InvestigationCaseSearchPage(IWebDriver driver)
            : base(driver, InvestigationCaseSearchPage.frameId)
        {
            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

            //Switch to main frame when it is visible
            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);


        }

        [ActionMethod]
        public void SetPageFilterList(string value)
        {
            UICommon.SetPageFilterList(value, driver);
            //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            //wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#crmGrid_SavedNewQuerySelector>span"))).Click();
            //IWebElement parent = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Dialog_0")));
            //parent.FindElement(By.XPath("//li[a[contains(@title,'"+value+"')]]")).Click();// and parent::*[@id='Dialog_0']]"))).Click();
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

        [ActionMethod]
        public bool VerifyNewInvestigationCaseButtonPresent()
        {
            this.driver.SwitchTo().DefaultContent();
            return UICommon.CheckElementExists("img[alt='New']", driver);
        }

        /*
        * search criteria
        * ************************************************************************
        */

        [ActionMethod]
        public void SetSearchRecord(string searchValue)
        {
            UICommon.SetSearchText("crmGrid_findCriteria", searchValue, driver);

        }

        /*
        * search criteria
        * ************************************************************************
        */

        [ActionMethod]
        public void SetInvestigationSearchText(string searchValue)
        {
            UICommon.SetSearchText("crmGrid_findCriteria", searchValue, driver);
        }


        /*
       * search table 
       * ************************************************************************
       */

        [ActionMethod]
        public IWebElement GetSearchResultTable()
        {
            return UICommon.GetSearchResultTable(driver);
        }

        [ActionMethod]
        public IWebElement GetHeaderSearchResultTable()
        {
            return UICommon.GetHeaderSearchResultTable(driver);
        }

        [ActionMethod]
        public void ClickPageTitle()
        {
             UICommon.ClickPageTitle(driver);
        }
    }

}
