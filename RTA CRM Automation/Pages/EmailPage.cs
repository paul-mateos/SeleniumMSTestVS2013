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




namespace RTA.Automation.CRM.Pages
{
    [ActionClass]
    public class EmailPage : IFramePage
    {
        private static string frameId = "contentIFrame0";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;
        private static string pageTitle = "Email";

        public EmailPage(IWebDriver driver)
            : base(driver, EmailPage.frameId)
        {

            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

            //Switch to main frame when it is visible
            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);

        }

        /*
        * Email Body
        * ************************************************************************
        */
        [ActionMethod]
        public void SetEmailBODYText(string emailBody)
        {
            this.driver.SwitchTo().Frame("descriptionIFrame");
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("html>body")));
            Actions action = new Actions(driver);
            action.MoveToElement(this.driver.FindElement(By.CssSelector("html>body"))).DoubleClick().Build().Perform();
            IWebElement elem = this.driver.FindElement(By.CssSelector("html>body"));
            elem.SendKeys(emailBody);
            this.driver.SwitchTo().DefaultContent();

        }


        /*
        * Set TO Value
        * ************************************************************************
        */
        [ActionMethod]
        public void SetToValueText(string toValue)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#to>div")));
            Actions action = new Actions(driver);
            action.MoveToElement(this.driver.FindElement(By.CssSelector("#to>div"))).DoubleClick().Build().Perform();
            IWebElement elem = this.driver.FindElement(By.Id("to_ledit_multi"));
            elem.SendKeys(toValue);
            
            
        }

        /*
        * Set CC Value
        * ************************************************************************
        */
        [ActionMethod]
        public void SetCCValueText(string ccValue)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#cc>div")));
            Actions action = new Actions(driver);
            action.MoveToElement(this.driver.FindElement(By.CssSelector("#cc>div"))).DoubleClick().Build().Perform();
            IWebElement elem = this.driver.FindElement(By.Id("cc_ledit_multi"));
            elem.SendKeys(ccValue);
            elem.SendKeys(Keys.Tab);
        }

        /*
        * Set BCC Value
        * ************************************************************************
        */
        [ActionMethod]
        public void SetBCCValueText(string bccValue)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#bcc>div")));
            Actions action = new Actions(driver);
            action.MoveToElement(this.driver.FindElement(By.CssSelector("#bcc>div"))).DoubleClick().Build().Perform();
            IWebElement elem = this.driver.FindElement(By.Id("bcc_ledit_multi"));
            elem.SendKeys(bccValue);
            elem.SendKeys(Keys.Tab);
        }

        /*
        * Set Regarding Value
        * ************************************************************************
        */
        [ActionMethod]
        public void SetRegardingValueText(string regardingValue)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//div[@id='regardingobjectid']/div")));
            elem.Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("regardingobjectid_i"))).Click();

            elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("regardingobjectid_ledit")));
            elem.Clear();
            elem.SendKeys(regardingValue);
            elem.SendKeys(Keys.Enter);
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//li[a[contains(@title,'" + regardingValue + "')] and parent::*[@id='regardingobjectid_i_IMenu']]"))).Click();
       }
        
        /*
        * Set Subject Value
        * ************************************************************************
        */
        [ActionMethod]
        public void SetSubjectValueText(string subjectValue)
        {
            
            UICommon.SetTextBoxValue("subject", subjectValue, driver);
        }

        /*
        * Send Email
        * ************************************************************************
        */
        [ActionMethod]
        public void ClickSendEmailIMG()
        {
 
            this.driver.SwitchTo().DefaultContent();
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[alt='Send']"))).Click();



        }

        [ActionMethod]
        public void ClickSaveIMG()
        {

            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickSaveButton(driver);
            this.driver.SwitchTo().Frame(frameId);


        }

        [ActionMethod]
        public void ClickSaveCloseIMG()
        {
 
            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickSaveCloseButton(driver);


        }

        /*
       * Message from Webpage
       * ************************************************************************
       */
        [ActionMethod]
        public void AcceptRTAValidationMessage(string msgValidation)
        {
            string alertMessage = driver.SwitchTo().Alert().Text;
            if (alertMessage.Contains(msgValidation)) 
            {
            driver.SwitchTo().Alert().Accept();
            }
            else
            {
                IAlert alert = driver.SwitchTo().Alert();
                string alertText = alert.Text;
                alert.Accept();
                throw new Exception("Failuer in test case: " + alertText);   
            }

        }


        internal void ClickPageTitle()
        {
            
            UICommon.ClickPageTitle(driver);
        }
    }
}
