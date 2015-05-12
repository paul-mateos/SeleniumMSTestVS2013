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
    public class TenancySearchPage : IFramePage
    {
        private static string FRAME = "contentIFrame0";
        private static string pageTitle = "Tenancies Active Tenancies";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;

        public TenancySearchPage(IWebDriver driver)
            : base(driver, TenancySearchPage.FRAME)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

            //Switch to main frame when it is visible
            driver.SwitchTo().DefaultContent();
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement f1 = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(FRAME)));
            driver.SwitchTo().Frame(f1);
        }
    
        /*
        * NewIMG
        * ************************************************************************
        */

        [ActionMethod]
        public void ClickNewTenancyButton()
        {
            driver.SwitchTo().DefaultContent();
            UICommon.ClickNewButton(driver);
           
        }

        /*
        * search criteria
        * ************************************************************************
        */

        [ActionMethod]
        public void SetTenancySearchText(string searchValue)
        {
            
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement element = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("crmGrid_quickFindContainer")));
            element = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("crmGrid_findCriteria")));          
            element.Clear();
            element.SendKeys(searchValue.ToString());
            element.SendKeys(Keys.Enter);
            
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
