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
    public class PaymentReferncePage : IFramePage
    {
        //public static string WINDOW = "Payment Reference: New Payment Reference - Microsoft Dynamics CRM";
        private static string frameId = "contentIFrame0";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;
        private static string pageTitle = "Payment";

        public PaymentReferncePage(IWebDriver driver)
            : base(driver, PaymentReferncePage.frameId)
        {
            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

            //Switch to main frame when it is visible
            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);

        }

     

        /*
        * Client Name
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

            UICommon.SetTextBoxValue("rta_amount", amountValue, driver);

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
      * Payment Type
      * ************************************************************************
      */
        [ActionMethod]
        public void SetPaymentTypeValue(string paymentType)
        {

            UICommon.SetSelectListValue("rta_payment_type", paymentType, driver);

        }

        [ActionMethod]
        public void ClickDeactivateButton()
        {

            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickDeactivateButton(driver);
            this.driver.SwitchTo().Frame(frameId);
        }

        [ActionMethod]
        public string GetInactiveStatusFooter()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#footer_statecode>div>span")));
            return elem.Text;
        }

        internal void ClickPageTitle()
        {

            UICommon.ClickPageTitle(driver);
        }
        
    }
}
