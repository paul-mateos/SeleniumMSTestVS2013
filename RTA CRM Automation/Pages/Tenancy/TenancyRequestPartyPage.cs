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
    public class TenancyRequestPartyPage : IFramePage
    {
        //public static string WINDOW = "Payment Reference: New Payment Reference - Microsoft Dynamics CRM";
        private static string frameId = "contentIFrame0";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;
        private static string pageTitle = "Tenancy Request Party:";

        public TenancyRequestPartyPage(IWebDriver driver)
            : base(driver, TenancyRequestPartyPage.frameId)
        {
            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

            //Switch to main frame when it is visible
            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);

        }


  
        /*
        * Search For Client
        * ************************************************************************
        */
        [ActionMethod]
        public void SetClientNameValue(string clientName)
        {
            
            UICommon.SetSearchableListValue("rta_clientid", clientName, driver);

        }

  
        /*
       * Amount
       * ************************************************************************
       */
        [ActionMethod]
        public void SetAmountValue(string amountValue)
        {

            UICommon.SetTextBoxValue("rta_contribution_amount", amountValue, driver);

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
         * SaveIMG
         * ************************************************************************
         */

        [ActionMethod]
        public void ClickSaveButton()
        {

            driver.SwitchTo().DefaultContent();
            UICommon.ClickSaveButton(driver);
            driver.SwitchTo().Frame(frameId);


        }

        /*
       * SaveCloseIMG
       * ************************************************************************
       */

        [ActionMethod]
        public void ClickSaveCloseButton()
        {
            driver.SwitchTo().DefaultContent();
            UICommon.ClickSaveCloseButton(driver);
        }

        /*
      * Payment Type
      * ************************************************************************
      */
        [ActionMethod]
        public void SetPaymentTypeValue(string paymentType)
        {

            UICommon.SetSelectListValue("rta_payment_type", paymentType, driver);

        }
        /*
      * Click Page Title
      * ************************************************************************
      */
        [ActionMethod]
        public void ClickPageTitle()
        {
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(frameId);
            UICommon.ClickPageTitle(driver);

        }

        /*
     * Deactivate
     * ************************************************************************
     */

        [ActionMethod]
        public void ClickDeactivateButton()
        {

            this.driver.SwitchTo().DefaultContent();

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[alt='Deactivate']")));
            Actions action = new Actions(driver);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).Release().Build().Perform();
         }
    }
}
