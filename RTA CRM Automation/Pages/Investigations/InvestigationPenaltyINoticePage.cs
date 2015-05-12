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
    public class InvestigationPenaltyINoticePage : IFramePage
    {
        private static string frameId = "contentIFrame0";
        //private static string dialogFRAME = "InlineDialog_Iframe";
        private static string pageTitle = "Penalty Infringement Notice:";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;

        public InvestigationPenaltyINoticePage(IWebDriver driver)
            : base(driver, InvestigationPenaltyINoticePage.frameId)
        {
            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

            //Switch to main frame when it is visible
            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);

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
            Thread.Sleep(2000);
        }

        [ActionMethod]
        public void ClickSaveCloseButton()
        {
            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickSaveCloseButton(driver);            
        }

        /*
        * PenaltyInfringementNotice Number
        * ************************************************************************
        */

        [ActionMethod]
        public string GetPenaltyInfringementNoticeNumber()
        {

            return UICommon.GetNewReferenceNumber(driver);
        }

        /*
        * Client Field
        * ************************************************************************
        */

        [ActionMethod]
        public void SetClientName(string ClientName)
        {
            UICommon.SetSearchableListValue("rta_clientid", ClientName, driver);
        }

        [ActionMethod]
        public void SetIssuedAgainstField(string IssuedAgainst)
        {
            UICommon.SetSelectListValue("rta_issued_against_client_type", IssuedAgainst, driver);
        }

          /*
          * Penalty Unit
          * ************************************************************************
          */
        [ActionMethod]

        public string GetPenaltyUnitsFieldText()
        {
            return UICommon.GetTextFromElement("#rta_penalty_units > div > span", driver);
            //return driver.FindElement(By.CssSelector("#rta_penalty_units > div > span")).GetAttribute("title");
        }

        public string GetPerUnitAmountFieldText()
        {
            return UICommon.GetTextFromElement("#rta_amount_per_penalty_unit > div > span", driver);
        }

        public string GetPenaltyAmountFieldText()
        {
            return UICommon.GetTextFromElement("#rta_penalty_amount > div > span", driver);
        }

        public void SetPenaltyUnits(string PenaltyUnits)
        {
            UICommon.SetTextBoxValue("rta_penalty_units", PenaltyUnits, driver);
        }

        public void SetPerUnitAmount(string PerUnitAmount)
        {
            UICommon.SetTextBoxValue("rta_amount_per_penalty_unit", PerUnitAmount, driver);
        }

        public void SetPenaltyAmountFieldText(string PenaltyAmount)
        {
            UICommon.SetTextBoxValue("rta_penalty_amount", PenaltyAmount, driver);
        }

        public bool CheckPenaltyUnitsLocked()
        {
            return UICommon.VerifyElementLocked("rta_penalty_units", driver);
        }

        public bool CheckPerUnitAmountLocked()
        {
            return UICommon.VerifyElementLocked("rta_amount_per_penalty_unit", driver);
        }

        public bool CheckPenaltyAmountFieldLocked()
        {
            return UICommon.VerifyElementLocked("rta_penalty_amount", driver);
        }

        [ActionMethod]
        public void ClickPageTitle()
        {
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(frameId);
            UICommon.ClickPageTitle(driver);
        }
    }
}
