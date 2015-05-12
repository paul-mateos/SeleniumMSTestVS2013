using ActionWordsLib.Attributes;
using AutoItX3Lib;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RTAAutomation.Utils
{
    [ActionClass]
    public class NavigateToURLWithAuth
    {
        private IWebDriver driver;
        public NavigateToURLWithAuth(IWebDriver driver)
        {
            this.driver = driver;
        }

        [ActionMethod]
        public void navigateToURLWithAuth(String URL, String username, String password)
        {
           
            //Initialize AutoIT
            var AutoIT = new AutoItX3();

            //Set Selenium page load timeout to 2 seconds so it doesn't wait forever
            //driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(2));

            
            driver.Navigate().GoToUrl(URL);

            //Wait 10 seconds for the authentication window to appear 
            string windowName = "Windows Security";
            int found = AutoIT.WinWait(windowName,"", 10);

            //If the window appears then send username and password
            if (found != 0)
            {
                AutoIT.WinActivate(windowName);
                AutoIT.Send(username);
                AutoIT.Send("{TAB}");
                AutoIT.Send(password);
                AutoIT.Send("{ENTER}");

            }
            driver.Manage().Window.Maximize();
            //return driver;
        
        }
    }
}
