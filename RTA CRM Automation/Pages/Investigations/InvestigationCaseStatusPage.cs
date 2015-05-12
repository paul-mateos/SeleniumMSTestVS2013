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

namespace RTA.Automation.CRM.Pages.Investigations
{

    [ActionClass]
    public class InvestigationCaseStatusPage : BasePage
    {

        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;
        private static string pageTitle = "Investigation Case Status";

        public InvestigationCaseStatusPage(IWebDriver driver)
            : base(driver)
        {
            //Wait for title to be displayed
            string title = driver.Title;
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });


        }

        /*
        * Close
        * ************************************************************************
        */
        [ActionMethod]
        public void CloseWindow()
        {
            // driver.SwitchTo().Window(driver.WindowHandles.ToList().Last()); // switches to the new window
            driver.Close();
        }
    }   
}
