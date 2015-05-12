using ActionWordsLib.Attributes;
using AutoItX3Lib;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RTA.Automation.CRM.Pages;
using RTA.Automation.CRM.UI;

namespace RTAAutomation.Utils
{
    [ActionClass]
    public class NavigateToURLWithAuthCU
    {
        private IWebDriver driver;
        public NavigateToURLWithAuthCU(IWebDriver driver)
        {
            this.driver = driver;
        }

        [ActionMethod]
        public void Login(String URL, string usernameParam, string passwordParam)
        {
           
            //Initialize AutoIT
           // var AutoIT = new AutoItX3();

            //Set Selenium page load timeout to 2 seconds so it doesn't wait forever
            //driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(2));

            
            driver.Navigate().GoToUrl(URL);

            LoginDialog loginDialog = new LoginDialog();
            loginDialog.Login(usernameParam, passwordParam);
       
        
        }
       
    }
}
