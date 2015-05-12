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

namespace RTA.Automation.CRM.Pages
{
    [ActionClass]
    public class NewActivityPage : BasePage
    {
       
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;
        private static string pageTitle = "Process:";

        public NewActivityPage(IWebDriver driver)
            : base(driver)
        {
            //Wait for title to be displayed
            string title = driver.Title;
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

          
        }




        [ActionMethod]
        public void ClickActivityListItem(string activity)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement parent = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("tblItems")));
            IWebElement elem = parent.FindElement(By.XPath("//*[contains(text(),'" + activity + "')]"));
            Actions action = new Actions(driver);
            action.MoveToElement(elem).Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(2000);
            action.MoveToElement(elem).Release().Build().Perform();
            
        }


        [ActionMethod]
        public void ClickOKButton()
        {
            UICommon.ClickElementWithId("butBegin", driver);
            
        }
    }
}
