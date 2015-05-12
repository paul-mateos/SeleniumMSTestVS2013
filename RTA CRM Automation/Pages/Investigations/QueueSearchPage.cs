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
    public class QueueSearchPage : IFramePage
    {
        private static string frameId = "contentIFrame0";
        private static string pageTitle = "Queue Items";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;

        public QueueSearchPage(IWebDriver driver)
            : base(driver, QueueSearchPage.frameId)
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
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#crmGrid_SavedNewQuerySelector>span"))).Click();
            IWebElement parent = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Dialog_0")));
            parent.FindElement(By.XPath("//li[a[contains(@title,'"+value+"')]]")).Click();
        
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

        /*
        * Select Queue
        * ************************************************************************
        */

        [ActionMethod]
        public void SetQueue(string queueValue)
        {
            driver.FindElement(By.Id("crmQueueSelector")).SendKeys(queueValue);
        }
    }
}
