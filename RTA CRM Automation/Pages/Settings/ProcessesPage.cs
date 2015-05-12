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
    public class ProcessesPage : BasePage
    {
       
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;
        private static string pageTitle = "Process:";

        public ProcessesPage(IWebDriver driver)
            : base(driver)
        {
            //Wait for title to be displayed
            string title = driver.Title;
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

          
        }
    
       
       
         /*
       * Activate
       * ************************************************************************
       */

        [ActionMethod]
        public void ClickActivateButton()
        {

 
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[alt='Activate']"))).Click();
 

        }

        /*
     * Deactivate
     * ************************************************************************
     */

        [ActionMethod]
        public void ClickDeactivateButton()
        {


            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[alt='Deactivate']"))).Click();


        }

      /*
     * Close
     * ************************************************************************
     */

        [ActionMethod]
        public void ClickCloseButton()
        {


            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[alt='Close']"))).Click();


        }
        

    }
}
