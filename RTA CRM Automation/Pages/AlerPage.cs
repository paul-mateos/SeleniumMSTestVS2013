using OpenQA.Selenium;
using System.Collections.ObjectModel;
using System.Linq;
using System;
using ActionWordsLib.Attributes;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System.Threading.Tasks;
using System.Threading;
using RTA.Automation.CRM.Utils;
using RTA.Automation.CRM.UI;
using System.Collections.Generic;




namespace RTA.Automation.CRM.Pages
{
    [ActionClass]
    public class AlertPage : IFramePage
    {
        private static string frameId = "contentIFrame0";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;
        private static string pageTitle = "Alert";

        public AlertPage(IWebDriver driver)
            : base(driver, AlertPage.frameId)
        {
            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);

        }



        /*
        * Other Alert
        * ************************************************************************
        */
        [ActionMethod]
        public void SetOtherAlertText(string otherAlert)
        {

            UICommon.SetTextBoxValue("rta_alert", otherAlert, driver);

        }

        [ActionMethod]
        public string GetOtherAlertControlMode()
        {
            return UICommon.GetElementProperty("#rta_alert", "data-controlmode", driver);
        }

        /*
        * Description
        * ************************************************************************
        */
        [ActionMethod]
        public void SetDesctiptionText(string otherAlert)
        {

            UICommon.SetTextBoxValue("rta_description", otherAlert, driver);

        }


        /*
        * Alert Type
        * ************************************************************************
        */

        [ActionMethod]
        public bool GetAlertTypeText(string alertType)
        {
            SetAlertTypeElement(alertType);
           
            return true;

        }

       
      

        [ActionMethod]
        public void SetAlertTypeElement(string alertType)
        {
            UICommon.SetSearchableListValue("rta_alert_typeid", alertType, driver);

        }
        

      

        /*
        * Save
        * ************************************************************************
        */

        [ActionMethod]
        public void ClickSaveIMG()
        {

            driver.SwitchTo().DefaultContent();
            UICommon.ClickSaveButton(driver);
            driver.SwitchTo().Frame(frameId); 


        }

        [ActionMethod]
        public void ClickSaveCloseIMG()
        {

            driver.SwitchTo().DefaultContent();
            UICommon.ClickSaveCloseButton(driver);
         }

        /*
        * Deactivate
        * ************************************************************************
        */

        [ActionMethod]
        public void ClickDeactivateIMG()
        {
            
            RefreshPageFrame.RefreshPage(driver, frameId);
            driver.SwitchTo().DefaultContent();
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[alt='Deactivate']")));
            Thread.Sleep(2000);
            Actions action = new Actions(driver);
            action.MoveToElement(driver.FindElement(By.CssSelector("img[alt='Deactivate']"))).ClickAndHold().Build().Perform();
            Thread.Sleep(2000);
            action.MoveToElement(driver.FindElement(By.CssSelector("img[alt='Deactivate']"))).Release().Build().Perform();
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[alt='Activate']")));
            driver.SwitchTo().Frame(frameId);
            

        }

        /*
        * Alert State
        * ************************************************************************
        */

        [ActionMethod]
        public string GetAlertState()
        {

            return UICommon.GetTextFromElement("#footer_statecode>div>span",driver);
        }

        

        /*
       * Message from Webpage
       * ************************************************************************
       */
        [ActionMethod]
        public void AcceptRTAValidationMessage(string msgValidation)
        {
            UICommon.GetAlertMessage(driver);

        }

        /*
        * Alert Number
        * ************************************************************************
        */

        [ActionMethod]
        public string GetAlertNumber()
        {

            return UICommon.GetNewReferenceNumber(driver);
        }


        /*
       * Alert Start Date
       * ************************************************************************
       */
        [ActionMethod]
        public void SetStartDateValue(string dateValue)
        {
            UICommon.SetDateValue("rta_start_date", dateValue, driver);

        }

        [ActionMethod]
        public void SetEndDateValue(string dateValue)
        {
            RefreshPageFrame.RefreshPage(driver, frameId); 
            UICommon.SetDateValue("rta_end_date", dateValue, driver);

        }

              
        [ActionMethod]
        public String GetStartDateErrorText()
        {
            return UICommon.GetTextFromElement("#rta_start_date_err", driver);
        }



        internal void ClickPageTitle()
        {
            UICommon.ClickPageTitle(driver);
        }

        public bool VerifyOtherAlertMandatoryFieldIconPresent()
        {
            IList<IWebElement> mandatoryIcon = driver.FindElements(By.CssSelector("#rta_alert_c>span>img[alt='Required']"));
            if (mandatoryIcon.Count > 0)
            {
                return true;
            }
            return false;
        }
    }
}
