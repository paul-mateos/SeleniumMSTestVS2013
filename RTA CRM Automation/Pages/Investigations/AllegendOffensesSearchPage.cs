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
    public class AllegendOffensesSearchPage : IFramePage
    {
        private static string frameId = "contentIFrame0";
        private string pageTitle = "Alleged Offences";
        private int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;

        public AllegendOffensesSearchPage(IWebDriver driver)
            : base(driver, AllegendOffensesSearchPage.frameId)
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
        public void ClickNewAllegedOffenceButton()
        {
            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickNewButton(driver);

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
    }

}
