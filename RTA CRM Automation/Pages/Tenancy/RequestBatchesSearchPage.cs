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
    public class RequestBatchesSearchPage : IFramePage
    {
        private static string frameId = "contentIFrame0";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;
        private static string pageTitle = "Request Batches";

        public RequestBatchesSearchPage(IWebDriver driver)
            : base(driver, RequestBatchesSearchPage.frameId)
        {
            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
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
        public void ClickNewRequestBatchButton()
        {
            driver.SwitchTo().DefaultContent();
            UICommon.ClickNewButton(driver);
           
        }

        /*
        * search criteria
        * ************************************************************************
        */

        [ActionMethod]
        public void SetRequestBatchSearchText(string searchValue)
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

        public bool GetPaymentRefernceRefreshTable(string tenancyrequest)
        {

           
            int i = 1;
            while (i <= 60) //waits for 60sec
            {
                Table table = new Table(this.GetSearchResultTable());
                if (table.GetCellValue("Name", tenancyrequest, "Status Reason") == "Pending Financials")
                {
                    return true;

                }else
                {   
                    Thread.Sleep(1000);
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
                    wait.Until(ExpectedConditions.ElementIsVisible(By.Id("grid_refresh"))).Click();
                    i++;
                }

            } throw new Exception(String.Format("Status Reason has not changed to Pending Financials"));
        }
      
    }
}
