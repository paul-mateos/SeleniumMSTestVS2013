
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
    public class AppointmentPage : IFramePage
    {
        private static string frameId = "contentIFrame0";        
        private static string pageTitle = "Appointment:";
        private static string dialogFRAME = "InlineDialog_Iframe";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;

        public AppointmentPage(IWebDriver driver)
            : base(driver, AppointmentPage.frameId)
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
        }

        [ActionMethod]
        public void ClickSaveCloseButton()
        {
            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickSaveCloseButton(driver);
            try
            {
                IReadOnlyCollection<IWebElement> ErrorFrame = driver.FindElements(By.Id(dialogFRAME));
                if (ErrorFrame.Count > 0)
                {
                    this.ClickIgnoreAndSaveButton();
                }
            }
            catch
            { 
            
            }
        }

        private void ClickIgnoreAndSaveButton()
        {
            // driver.SwitchTo().DefaultContent();
            //driver.SwitchTo().Frame(dialogFRAME);
            //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            //wait.Until(ExpectedConditions.ElementIsVisible(By.Id("butSave"))).Click();
            //driver.SwitchTo().DefaultContent();
            //driver.SwitchTo().Frame(frameId);

            RefreshPageFrame.RefreshPage(driver, dialogFRAME);

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
            IWebElement parent = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("tdDialogFooter")));
            IWebElement elem = parent.FindElement(By.Id("btnSave"));
            Actions action = new Actions(driver);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(2000);
            action.MoveToElement(elem).Release().Build().Perform();
   
        }

        [ActionMethod]
        public void SetSubjectValue(string subject)
        {
            //Switch to main frame when it is visible
            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);
            UICommon.SetTextBoxValue("subject", subject, driver);
        }
       
        [ActionMethod]
        public void ClickPageTitle()
        {
            UICommon.ClickPageTitle(driver);            
        }

        [ActionMethod]
        public void ClickRecurrenceButton()
        {
            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickImageButton(driver, "Recurrence");
            this.driver.SwitchTo().Frame(frameId);
        }

        [ActionMethod]
        public void SetStartRange(string date)
        { 
            ////driver.SwitchTo().DefaultContent();
            ////driver.SwitchTo().Frame(dialogFRAME);
            //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            //IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("scheduledstart")));
            //driver.FindElement(By.CssSelector("#scheduledstart > div > span")).Click();
            //driver.FindElement(By.Id("DateInput")).Clear();
            //driver.FindElement(By.Id("DateInput")).SendKeys(date);
            //// UICommon.SetDateValue("", DateTime.Today.ToString("dd/mm/yyyy"),driver);

            UICommon.SetDateValue("scheduledstart", date, driver);
        }

        [ActionMethod]
        public void ClickSetButton()
        {
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(dialogFRAME);
            driver.FindElement(By.Id("button_ok")).Click();
        }
    }
}
