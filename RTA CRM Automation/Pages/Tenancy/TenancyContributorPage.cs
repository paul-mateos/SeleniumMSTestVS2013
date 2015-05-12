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


namespace RTA.Automation.CRM.Pages
{
    [ActionClass]
    public class TenancyContributorPage : IFramePage
    {
        public static string frameId = "contentIFrame0";
        private static string pageTitle = "Tenancy Contributor";
        public static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;

        public TenancyContributorPage(IWebDriver driver)
            : base(driver, TenancyContributorPage.frameId)
        {
            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);
        }

   

        /*
        * Start Date
        * ************************************************************************
        */
      

        [ActionMethod]
        public void SetStartDateValue(string startDate)
        {
            
            UICommon.SetDateValue("rta_start_date", startDate, driver);
        }

        /*
        * End Date
        * ************************************************************************
        */
       

        [ActionMethod]
        public void SetEndDateValue(string startDate)
        {

            UICommon.SetDateValue("rta_end_date", startDate, driver);
        }

      

        /*
        * SaveIMG
        * ************************************************************************
        */

        [ActionMethod]
        public void ClickSaveButton()
        {

            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickSaveButton(driver);
            this.driver.SwitchTo().Frame(frameId);
           
        }

        /*
       * SaveCloseIMG
       * ************************************************************************
       */

        [ActionMethod]
        public void ClickSaveCloseButton()
        {
            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickSaveCloseButton(driver);
        }

    
        /*
        * End Date ERROR
        * ************************************************************************
        */
        [ActionMethod]
        public String GetStartDateErrorText()
        {
            
            return UICommon.GetTextFromElement("#rta_start_date_err", driver);
        }



        internal void ClickPageTitle()
        {
            
            UICommon.ClickPageTitle(driver);
        }
    }
}
