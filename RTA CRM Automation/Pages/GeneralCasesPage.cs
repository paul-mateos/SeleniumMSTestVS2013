using ActionWordsLib.Attributes;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using RTA.Automation.CRM.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;


namespace RTA.Automation.CRM.Pages
{
    [ActionClass]
    public class GeneralCasesPage : IFramePage
    {
        private static string FRAME = "contentIFrame1";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;

        public GeneralCasesPage(IWebDriver driver)
            : base(driver, GeneralCasesPage.FRAME)
        {
            
        }

        [ScenarioMethod]
        public void PopulateNewClient(string familyName)
        {
            //this.ClickFamilyNameTextBoxArea();
            //this.SetFamilyName(familyName);
        }


        /*
        * NewIMG
        * ************************************************************************
        */

        [ActionMethod]
        public void ClickNewClientButton()
        {
            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickNewButton(driver);
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
            this.driver.SwitchTo().Frame(this.frame);
        }

       /*
       * ClientID
       * ************************************************************************
       */

        [ActionMethod]
        public string GetClientID()
        {
            return UICommon.GetNewReferenceNumber(driver);
        }

        /*
       * Unknown Client
       * ************************************************************************
       */

        [ActionMethod]
        public string GetUnknownClientListValues()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            string textValue = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_unknownclient_i"))).Text;
            return textValue;         
        }
    }
}
