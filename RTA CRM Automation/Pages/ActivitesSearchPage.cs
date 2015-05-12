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
    public class ActivitiesSearchPage : IFramePage
    {
        private static string frameId = "contentIFrame0";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;
        private static string pageTitle = "Activities";

        public ActivitiesSearchPage(IWebDriver driver)
            : base(driver, ActivitiesSearchPage.frameId)
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
        public void ClickNewTenancyRequestButton()
        {
            driver.SwitchTo().DefaultContent();
            UICommon.ClickNewButton(driver);
            
           
        }

        /*
        * Grid Refresh
        * ************************************************************************
        */

        [ActionMethod]
        public void ClickTableRefresh()
        {
            UICommon.ClickTableRefreshButton(driver);
            
           
        }
        

        /*
        * search criteria
        * ************************************************************************
        */

        [ActionMethod]
        public void SetTenancyRequestSearchText(string searchValue)
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
        public void SetPageFilterList(string filterValue)
        {
            UICommon.SetPageFilterList(filterValue, driver);
            //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            //wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#crmGrid_SavedNewQuerySelector>span"))).Click();
            //IWebElement parent = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Dialog_0")));
            //parent.FindElement(By.CssSelector("li a[title='" + filterValue + "']")).Click();
        }

        [ActionMethod]
        public IWebElement GetSearchResultRow(int rowIndex = 0)
        {
            IWebElement element = this.GetSearchResultTable();
            IReadOnlyCollection<IWebElement> tableRows = element.FindElements(By.CssSelector("tr.ms-crm-List-Row"));

            if(tableRows.Count > 0)
            {
                return tableRows.ElementAt(rowIndex);
            }

            throw new Exception("There are no results in the Page Filter to select");
        }
      
    }
}
