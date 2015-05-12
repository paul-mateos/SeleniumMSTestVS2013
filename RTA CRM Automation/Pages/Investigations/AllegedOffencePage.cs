using ActionWordsLib.Attributes;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Interactions.Internal;
using System.Collections.ObjectModel;
using System.Threading;
using RTA.Automation.CRM.Utils;
using RTA.Automation.CRM.UI;




namespace RTA.Automation.CRM.Pages
{
    [ActionClass]
    public class AllegedOffencePage : IFramePage
    {
        //public static string WINDOW = "Payment Reference: New Payment Reference - Microsoft Dynamics CRM";
        private static string frameId = "contentIFrame0";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;
        private static string pageTitle = "Alleged Offence:";

        public AllegedOffencePage(IWebDriver driver)
            : base(driver, AllegedOffencePage.frameId)
        {
            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

            //Switch to main frame when it is visible
            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);

            //setfocus to title
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("FormTitle"))).Click();

        }

            

        /*
      * Reference Number
      * ************************************************************************
      */
        [ActionMethod]
        public string GetReferenceNumber()
        {

            return UICommon.GetNewReferenceNumber(driver);
            
        }

        /*
       * Status Reason
       * ************************************************************************
       */
      

        /*
       * Save
       * ************************************************************************
       */

        [ActionMethod]
        public void ClickSaveButton()
        {

            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickSaveButton(driver);
            this.driver.SwitchTo().Frame(frameId);

        }

        [ActionMethod]
        public void ClickSaveCloseButton()
        {

            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickSaveCloseButton(driver);

        }

      

        /*
        * Provision
        * ************************************************************************
        */
        [ActionMethod]
        public void SetProvisionValue(string provision)
        {
            UICommon.SetSearchableListValue("rta_provisionid", provision, driver);
        }
        
        /*
        * Outcome
        * ************************************************************************
        */
        [ActionMethod]
        public void SetOutcomeValue(string outcome)
        {
            UICommon.SetSearchableListValue("rta_outcomeid", outcome, driver);
        }

        /*
       * Offences Date
       * ************************************************************************
       */
        [ActionMethod]
        public void SetOffenceDateValue(string value)
        {

            UICommon.SetDateValue("rta_offence_date", value, driver);
            
        }

        [ActionMethod]
        public string GetOffenceDateValue()
        {
            return UICommon.GetTextFromElement("#rta_offence_date>div", driver);
            
        }

        [ActionMethod]
        public string GetStatutoryLimitationValue()
        {

            return UICommon.GetTextFromElement("#rta_statutory_limitation>div", driver);

        }

        /*
      * Belief Formed Date
      * ************************************************************************
      */
        [ActionMethod]
        public void SetBeliefFormedDateValue(string value)
        {

            UICommon.SetDateValue("rta_belief_formed_on", value, driver);

        }

        [ActionMethod]
        public string GetBefliefFormedDateValue()
        {

            return UICommon.GetTextFromElement("#rta_belief_formed_on>div", driver);


        }

       /*
      * Investigation Case 
      * ************************************************************************
      */
        [ActionMethod]
        public void SetInvestigationCaseValue(string value)
        {

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("FormTitle"))).Click();
            
            UICommon.SetSearchableListValue("rta_investigation_caseid", value, driver);

        }


        public string GetStatusReason()
        {
            return UICommon.GetTextFromElement("#statuscode>div>span", driver);
        }

        internal void ClickPageTitle()
        {
            UICommon.ClickPageTitle(driver);
        }
    }
}
