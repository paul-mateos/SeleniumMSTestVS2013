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
    public class UserPage : IFramePage
    {
        private static string frameId = "contentIFrame0";
        private static string frameId2 = "contentIFrame1";
        private static string investigationFRAME = "rta_systemuser_rta_inv_case_investigatoridFrame";
        private static string connectionFRAME = "areaConnectionsFrame";
        private static string pageTitle = "User:";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;

        public UserPage(IWebDriver driver)
            : base(driver, UserPage.frameId)
        {
            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

            //Switch to main frame when it is visible
            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);

        }

        [ActionMethod]
        public void ClickPageTitle()
        {
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(frameId);
            UICommon.ClickPageTitle(driver);
        }

        [ActionMethod]
        public string GetFullName()
        {
            return UICommon.GetTextFromElement("#fullname>div>span",driver);
        }

        /*
 * search table 
 * ************************************************************************
 */

        [ActionMethod]
        public IWebElement GetSearchResultTable()
        {
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame("contentIFrame1");
            driver.SwitchTo().Frame(investigationFRAME);
            WaitForPageToLoad.WaitToLoad(driver);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("gridBodyTable")));
            IWebElement webElementBody = driver.FindElement(By.Id("gridBodyTable"));
            return webElementBody;
        }

        /*
        * search criteria
        * ************************************************************************
        */

        [ActionMethod]
        public void SetSearchRecord(string searchValue)
        {
            WaitForPageToLoad.WaitToLoad(driver);
            driver.SwitchTo().Frame(investigationFRAME);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement element = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("crmGrid_rta_systemuser_rta_inv_case_investigatorid_findCriteria")));
            driver.FindElement(By.Id("crmGrid_rta_systemuser_rta_inv_case_investigatorid_findCriteria"));
            element.Clear();
            element.SendKeys(searchValue.ToString());
            element.SendKeys(Keys.Enter);
        }

        /*
        * Search Filter List
        * ************************************************************************
        */
        [ActionMethod]
        public string GetPageFilterList()
        {
            driver.SwitchTo().Frame(investigationFRAME);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement element = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#crmGrid_rta_systemuser_rta_inv_case_investigatorid_SavedNewQuerySelector>span")));
            return element.Text;
        }

        public void SwitchConnectionFrame()
        {
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(frameId2);
            driver.SwitchTo().Frame(connectionFRAME);        
        }

        [ActionMethod]
        public IWebElement GetConnectionsTable()
        {
            this.SwitchConnectionFrame();
            return UICommon.GetSearchResultTable(driver);
        }

        [ActionMethod]
        public void SetConnectionsSearchRecord(string searchValue)
        {
            this.SwitchConnectionFrame();
            UICommon.SetSearchText("crmGrid_systemuser_connections1_findCriteria", searchValue, driver);
        }
    }
}
