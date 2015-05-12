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
    public class InvestigationPenaltyINoticeSearchPage : IFramePage
    {
        private static string frameId = "contentIFrame0";
        private static string pageTitle = "Penalty Infringement Notices";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;

        public InvestigationPenaltyINoticeSearchPage(IWebDriver driver)
            : base(driver, InvestigationPenaltyINoticeSearchPage.frameId)
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
        public void ClickNewPenaltyNoticeButton()
        {
            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickNewButton(driver);

        }

        /*
        * search criteria
        * ************************************************************************
        */


        [ActionMethod]
        public void SetSearchText(string searchValue)
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
    }
}

