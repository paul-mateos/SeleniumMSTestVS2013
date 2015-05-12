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

namespace RTA.Automation.CRM.Pages.Clients
{
    [ActionClass]
    class CopyAddressPage : BasePage
    {
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;
        private static string pageTitle = "Copy an address";

        public CopyAddressPage(IWebDriver driver)
            : base(driver)
        {
            //Wait for title to be displayed
            string title = driver.Title;
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

        }

        /*
       * Next
       * ************************************************************************
       */

        [ActionMethod]
        public void ClickNextButton()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("butNext"))).Click();
        }

        /*
        * Finish
        * ************************************************************************
        */

        [ActionMethod]
        public void ClickFinishButton()
        {


            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("butFinish"))).Click();


        }

        [ActionMethod]
        public void ClickPreviousButton()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("butPrev"))).Click();
        }

        [ActionMethod]
        public void SetCopyPostalAddressYes()
        {
            UICommon.SetRadioButton(driver,"rad_InteractionStep41");
        }

        [ActionMethod]
        public void SetCopyPostalAddressNo()
        {
            UICommon.SetRadioButton(driver,"rad_InteractionStep42");
        }

        [ActionMethod]
        public void SetCopyPhysicalAddressYes()
        {
            UICommon.SetRadioButton(driver,"rad_InteractionStep31");
        }

        [ActionMethod]
        public void SetCopyPhysicalAddressNo()
        {
            UICommon.SetRadioButton(driver,"rad_InteractionStep32");
        }

        [ActionMethod]
        public string GetErrorMessage()
        {
            return UICommon.GetTextFromElement(".PromptText", driver);
        }

        [ActionMethod]
        public string GetFinishMessage()
        {
            return UICommon.GetTextFromElement(".FinishText", driver);
        }


    }
}
