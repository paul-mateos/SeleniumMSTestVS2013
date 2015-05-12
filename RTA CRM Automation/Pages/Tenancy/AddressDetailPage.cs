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
    class AddressDetailPage :IFramePage
    {
        private static string frameId = "contentIFrame0";
        private static string FRAMETenancyRequest = "rta_address_detail_rta_tenancy_requestFrame";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;
        private static string pageTitle = "Address Detail:";

        public AddressDetailPage(IWebDriver driver)
            : base(driver, AddressDetailPage.frameId)
        {
            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

            //Switch to main frame when it is visible
            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);

        }

        [ActionMethod]
        public void ClickAddNewTenancyRequestButton()
        {
            RefreshPageFrame.RefreshPage(driver, frameId);
            this.driver.SwitchTo().Frame(FRAMETenancyRequest);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[alt='Add New Tenancy Request']"))).Click();
            
        }
        
        public IWebDriver SwitchNewBrowser(IWebDriver driver, string BaseWindow)
        {
            return UICommon.SwitchToNewBrowser(driver, BaseWindow);
        }

        [ActionMethod]
        public IWebElement GetTenancyRequestSearchResultTable()
        {
            RefreshPageFrame.RefreshPage(driver,frameId);
            driver.SwitchTo().Frame(FRAMETenancyRequest);
            return UICommon.GetSearchResultTable(driver);
        }
           
        [ActionMethod]
        public void SetTenancyRequestSearchText(string searchValue)
        {
            driver.SwitchTo().Frame(FRAMETenancyRequest);

            UICommon.SetSearchText("crmGrid_rta_address_detail_rta_tenancy_request_findCriteria", searchValue, driver);

            //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            //IWebElement element = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("crmGrid_rta_address_detail_rta_tenancy_request_findCriteria")));
            //driver.FindElement(By.Id("crmGrid_rta_address_detail_rta_tenancy_request_findCriteria"));
            //element.Clear();
            //element.SendKeys(searchValue.ToString());
            //element.SendKeys(Keys.Enter);   

        }

    }
}
