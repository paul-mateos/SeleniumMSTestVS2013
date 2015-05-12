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
    public class ConnectionPage : IFramePage
    {
        private static string frameId = "contentIFrame";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;
        private static string pageTitle = "Connection";

        public ConnectionPage(IWebDriver driver)
            : base(driver, ConnectionPage.frameId)
        {
            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);


        }



       
      
        /*
        * Description
        * ************************************************************************
        */
        [ActionMethod]
        public void SetDesctiptionText(string description)
        {

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement parent = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("description_container")));
            parent.Click();
            IWebElement elem = parent.FindElement(By.Id("description"));
            elem.Clear();
            elem.SendKeys(description);
            elem.SendKeys(Keys.Enter);


        }


       

        [ActionMethod]
        public void SetNameText(string name)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement parent = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#record2id_lookupTable>tbody>tr>td>div")));
            parent.Click();
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("record2id_ledit")));
            elem.Clear();
            elem.SendKeys(name);
            elem.SendKeys(Keys.Enter);
            //wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//li[a[contains(@title,'" + name + "')] and parent::*[@id='record2id_IMenu']]"))).Click();
            //wait.Until(ExpectedConditions.ElementIsVisible(By.Id("form_title_div"))).Click();


        }

        [ActionMethod]
        public void SetAsThisRoleText(string role)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement parent = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#record2roleid_lookupTable>tbody>tr>td>img")));
            parent.Click();
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("record2roleid_ledit")));
            elem.Clear();
            elem.SendKeys(role);
            elem.SendKeys(Keys.Enter);
            //wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//li[a[contains(@title,'" + role + "')] and parent::*[@id='record2roleid_IMenu']]"))).Click();
            //wait.Until(ExpectedConditions.ElementIsVisible(By.Id("form_title_div"))).Click();
            
            
        }

        

       

        /*
        * Save
        * ************************************************************************
        */

        [ActionMethod]
        public void ClickSaveIMG()
        {

            driver.SwitchTo().DefaultContent();
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("connection|NoRelationship|Form|Mscrm.Form.connection.Save-Large")));

            Actions action = new Actions(driver);
            action.MoveToElement(driver.FindElement(By.Id("connection|NoRelationship|Form|Mscrm.Form.connection.Save-Large"))).ClickAndHold().Build().Perform();
            Thread.Sleep(2000);
            action.MoveToElement(driver.FindElement(By.Id("connection|NoRelationship|Form|Mscrm.Form.connection.Save-Large"))).Release().Build().Perform();
            driver.SwitchTo().Frame(frameId); 


        }

        [ActionMethod]
        public void ClickSaveCloseIMG()
        {

            driver.SwitchTo().DefaultContent();
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("connection|NoRelationship|Form|Mscrm.Form.connection.SaveAndClose-Large")));

            Actions action = new Actions(driver);
            action.MoveToElement(driver.FindElement(By.Id("connection|NoRelationship|Form|Mscrm.Form.connection.SaveAndClose-Large"))).ClickAndHold().Build().Perform();
            Thread.Sleep(2000);
            action.MoveToElement(driver.FindElement(By.Id("connection|NoRelationship|Form|Mscrm.Form.connection.SaveAndClose-Large"))).Release().Build().Perform();
            driver.SwitchTo().Frame(frameId); 
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

        }

       


       
        [ActionMethod]
        public void ClickStartDate()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#rta_start_date>div>span"))).Click();
            
        }

        [ActionMethod]
        public void SetStartDate(string dateValue)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("effectivestart")));
            elem.Click();

            IWebElement element = elem.FindElement(By.Id("DateInput"));
            element.Clear();
            element.SendKeys(dateValue.ToString());
            element.SendKeys(Keys.Enter);
            WaitForPageToLoad.WaitToLoad(driver);

        }

       
        [ActionMethod]
        public void ClickEndDate()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#rta_end_date>div>span"))).Click();

        }

        [ActionMethod]
        public void SetEndDate(string dateValue)
        {
            RefreshPageFrame.RefreshPage(driver, frameId); WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("effectiveend")));
            elem.Click();
            
            IWebElement element = elem.FindElement(By.Id("DateInput"));
            element.Clear();
            element.SendKeys(dateValue.ToString());
            element.SendKeys(Keys.Enter);
            WaitForPageToLoad.WaitToLoad(driver);

        }

     

       
        [ActionMethod]
        public String GetStartDateErrorText()
        {

            return UICommon.GetTextFromElement("#rta_start_date_err", driver);
        }

        [ActionMethod]
        public void ClickPageTitle()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#form_title_div>div>h1")));
            Actions action = new Actions(driver);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).Release().Build().Perform();

        }

    }
}
