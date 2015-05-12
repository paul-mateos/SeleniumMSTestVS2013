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
    public class TenancyRequestQueueReasonPage : IFramePage
    {
        private static string frameId = "contentIFrame1";
        
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;
        private static string pageTitle = "Queue Reason";

        public TenancyRequestQueueReasonPage(IWebDriver driver)
            : base(driver, TenancyRequestQueueReasonPage.frameId)
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
        * General - Click Override
        * Ticks the box if tickCheckBox is true, unticks it if it's false. 
        * ************************************************************************
        */
        [ActionMethod]
        public void SetOverrideCheckBox(bool tickCheckBox)
        {
            if(tickCheckBox)
            {
                if(!driver.FindElement(By.Id("rta_override_i")).Selected)
                {
                    driver.FindElement(By.Id("rta_override_i")).Click();
                }
            }
            else
            {
                if(driver.FindElement(By.Id("rta_override_i")).Selected)
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

            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickSaveCloseButton(driver);

        }

        /*
      * Save Unsaved Changes
      * ************************************************************************
      */

        [ActionMethod]
        public void ClickUnsavedChangesButton()
        {

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("savefooter_statuscontrol")));

            Actions action = new Actions(driver);
            Thread.Sleep(2000);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(2000);
            action.MoveToElement(elem).Release().Build().Perform();
            Thread.Sleep(2000);
            WaitForPageToLoad.WaitToLoad(driver);

        }



        internal void SetStatusReasonValue(string p, IWebDriver d)
        {
            UICommon.SetSelectListValue("statuscode", p, d);
        }
    }
}
