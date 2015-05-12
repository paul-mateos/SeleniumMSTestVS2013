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
    public class ClientPhoneNumberPage : IFramePage
    {
        //public static string WINDOW = "Payment Reference: New Payment Reference - Microsoft Dynamics CRM";
        private static string frameId = "contentIFrame0";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;
        private static string pageTitle = "Client Phone Number";

        //protected IWebDriver driver = null;

        public ClientPhoneNumberPage(IWebDriver driver)
            : base(driver, ClientPhoneNumberPage.frameId)
        {

            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);

        }

        /*
        * Availability
        * ************************************************************************
        */

        [ActionMethod]
        public void ClickAvailabilityList()
        {
       
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#FormTitle>h1")));
            string title = elem.GetAttribute("title");
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#rta_availability>div"))).Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_availability_i"))).GetAttribute("id");

        }

        [ActionMethod]
        public bool GetAvailabilityListValue(string listValue)
        {
            new SelectElement(driver.FindElement(By.Id("rta_availability_i"))).SelectByText(listValue);
            return true;

        }

        [ActionMethod]
        public bool GetAvailabilityListItem(string listValue)
        {
            ClickAvailabilityList();
            GetAvailabilityListValue(listValue);
            return true;

        }

        /*
        * Type
        * ************************************************************************
        */

        //[ActionMethod]
        //public void ClickTypeList()
        //{
        //    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
        //    wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#rta_type>div"))).Click();
        //    wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_type_i"))).GetAttribute("id");

        //}

        //[ActionMethod]
        //public bool GetTypeListValue(string listValue)
        //{
        //    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
        //    wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_type_i"))).GetAttribute("id");
        //    new SelectElement(driver.FindElement(By.Id("rta_type_i"))).SelectByText(listValue);
        //    return true;

        //}

        //[ActionMethod]
        //public bool GetTypeListItem(string listValue)
        //{
        //    ClickTypeList();
        //    GetTypeListValue(listValue);
        //    return true;

        //}

        [ActionMethod]
        public void SetTypeListValue(string listValue)
        {
            //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            //wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_type_i"))).GetAttribute("id");
            //new SelectElement(driver.FindElement(By.Id("rta_type_i"))).SelectByText(listValue);

            UICommon.SetSelectListValue("rta_type", listValue, driver);
 
        }

        /*
        * Area Code
        * ************************************************************************
        */

        //[ActionMethod]
        //public void ClickAreaCodeElement()
        //{
        //    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
        //    wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#rta_area_code>div"))).Click();
        //    wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_area_code_i"))).GetAttribute("id");

        //}

        
        [ActionMethod]
        public void SetAreaCodeValue(string areaCode)
        {
            //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            //IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_area_code_i")));
            //elem.Clear();
            //elem.SendKeys(areaCode);
            //elem.SendKeys(Keys.Enter);

            UICommon.SetTextBoxValue("rta_area_code", areaCode, driver);
        }

        /*
        * Phone Number
        * ************************************************************************
        */

        //[ActionMethod]
        //public void ClickPhoneNumberElement()
        //{
        //    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
        //    wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#rta_phone_number>div"))).Click();
        //    wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_phone_number_i"))).GetAttribute("id");

        //}


        [ActionMethod]
        public void SetPhoneNumberValue(string phoneNumber)
        {
            //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            //IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_phone_number_i")));
            //elem.Clear();
            //elem.SendKeys(phoneNumber);
            //elem.SendKeys(Keys.Enter);
            UICommon.SetTextBoxValue("rta_phone_number", phoneNumber, driver);



        }

        [ActionMethod]
        public string GetPhoneNumber()
        {
            //************Replace with Custome expected coditions class***********************************
            //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            //wait.Until(d => !d.FindElement(By.Id("FormTitle")).Text.Contains("New"));
            //IWebElement element = this.driver.FindElement(By.CssSelector("#FormTitle>h1"));
            //string value = element.Text;
            //return value;
            return UICommon.GetTextFromElement("#FormTitle>h1", driver);
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
      * New
      * ************************************************************************
      */

        [ActionMethod]
        public void ClickNewButton()
        {

            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickNewButton(driver);
            this.driver.SwitchTo().Frame(frameId);

        }
        /*
       * Client Name
       * ************************************************************************
       */


        [ActionMethod]
        public void SetClientNameList(string clientname)
        {

            //frameId = UICommon.FindVisibleIFrame(driver);
            //RefreshPageFrame.RefreshPage(driver, frameId); 

            //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            //IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_clientid")));
            //elem.Click();
            //elem.FindElement(By.Id("rta_clientid_i")).Click();

            //elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_clientid_ledit")));
            //elem.Clear();
            //elem.SendKeys(clientname);
            //elem.SendKeys(Keys.Enter);

            //wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//li[a[contains(@title,'" + clientname + "')] and parent::*[@id='rta_clientid_i_IMenu']]"))).Click();

            UICommon.SetSearchableListValue("rta_clientid", clientname, driver);
 
        }

        //[ActionMethod]
        //public IList<IWebElement> GetRtaClientNameListValues()
        //{
        //    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
        //    wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_clientid_i_IMenu")));

        //    IWebElement elem = driver.FindElement(By.Id("rta_clientid_i_IMenu"));
        //    SelectElement selectList = new SelectElement(elem);
        //    IList<IWebElement> options = selectList.Options;
        //    return options;

        //}


        [ActionMethod]
        public String GetClientNameListItem()
        {
            return new SelectElement(this.driver.FindElement(By.Id("rta_clientid_i"))).SelectedOption.Text;
        }

        [ActionMethod]
        public void SetClientNameListValue(string clientName)
        {
            new SelectElement(this.driver.FindElement(By.Id("rta_clientid_i"))).SelectByText(clientName);
        }

        [ActionMethod]
        public bool GetClientNameListValue(string clientName)
        {
            SetClientNameList(clientName);
            //SetClientNameListValue(clientName);
            return true;
        }

        [ActionMethod]
        public void ClickPageTitle()
        {
            UICommon.ClickPageTitle(driver);

        }
        
    }
}
