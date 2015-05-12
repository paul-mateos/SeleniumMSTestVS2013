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
using System.Collections.ObjectModel;
using RTA.Automation.CRM.UI;

namespace RTA.Automation.CRM.Pages
{
    [ActionClass]
    public class TenancyRequestQueueReasonStatusPage : IFramePage
    {
        private static string frameId = "contentIFrame1";
        
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;
        private static string pageTitle = "Request Queue Reason";

        public TenancyRequestQueueReasonStatusPage(IWebDriver driver)
            : base(driver, TenancyRequestQueueReasonStatusPage.frameId)
        {

           
            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

            //Switch to main frame when it is visible
            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);           
        }

        /*
       * General - Reason
       * ************************************************************************
       */
        [ActionMethod]
        public string GetReasonValue()
        {
            return UICommon.GetTextFromElement("div#rta_reasonid span", driver);
        }

        /*
       * Set the status for the reason
       * ************************************************************************
       */
        [ActionMethod]
        public void SetReasonValue(string reason)
        {
            UICommon.SetSelectListValue("statuscode", reason, driver);
        }

        /*
       * Get the status for the reason
       * ************************************************************************
       */
        [ActionMethod]
        public string GetReasonValue(string reason)
        {
            return UICommon.GetTextFromElement("div#statuscode span",driver);
        }

        /*
       * Click the Save and Close Button
       * ************************************************************************
       */
        [ActionMethod]
        public void ClickSaveCloseButton()
        {
            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickSaveCloseButton(driver);
        }
       
        /*
       * General - Click Override
       * Ticks the box if tickCheckBox is true, unticks it if it's false. 
       * ************************************************************************
       */
        [ActionMethod]
        public void SetOverrideCheckBox(bool tickCheckBox)
        {
            if (tickCheckBox)
            {
                if (!driver.FindElement(By.Id("rta_override_i")).Selected)
                {
                    driver.FindElement(By.Id("rta_override_i")).Click();
                }
            }
            else
            {
                if (driver.FindElement(By.Id("rta_override_i")).Selected)
                {
                    driver.FindElement(By.Id("rta_override_i")).Click();
                }
            }
        }

        /*
        * General - Get Override checkbox value
        * Returns true or false 
        * ************************************************************************
        */
        [ActionMethod]
        public bool GetOverrideCheckBoxValue()
        {
            return driver.FindElement(By.Id("rta_override_i")).Selected;
        }
    }
}
