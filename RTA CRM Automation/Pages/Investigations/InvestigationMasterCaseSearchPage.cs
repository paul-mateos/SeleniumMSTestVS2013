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
    public class InvestigationMasterCaseSearchPage : IFramePage
    {
        private static string frameId = "contentIFrame0";
        private static string pageTitle = "Investigation Master Cases";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;

        public InvestigationMasterCaseSearchPage(IWebDriver driver)
            : base(driver, InvestigationMasterCaseSearchPage.frameId)
        {
            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

            //Switch to main frame when it is visible
            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);


        }

        /*
       * New Button
       * ************************************************************************
       */

        [ActionMethod]
        public void ClickNewButton()
        {
            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickNewButton(driver);
            
        }
        /*
        * Create Button
        * ************************************************************************
        */

        [ActionMethod]
        public void ClickCreateButton()
        {
            this.driver.SwitchTo().DefaultContent();
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("navTabGlobalCreateImage")));
            Actions action = new Actions(driver);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();          
            Thread.Sleep(2000); //Temporary
            action.MoveToElement(elem).Release().Build().Perform();
            Thread.Sleep(1000);

        }

        [ActionMethod]
        public void ClickTaskButton()
        {

           
            //this.driver.SwitchTo().DefaultContent();
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement parent = wait.Until(ExpectedConditions.ElementExists(By.Id("actionGroupControl")));
            IWebElement elem = parent.FindElement(By.Id("4212"));
            Actions action = new Actions(driver);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(2000); //Temporary
            action.MoveToElement(elem).Release().Build().Perform();

        }

        /*
        * search criteria
        * ************************************************************************
        */

        [ActionMethod]
        public void SearchRecord(string searchValue)
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
    }

}
