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

namespace RTA.Automation.CRM.Utils
{
    class RefreshPageFrame
    {
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;

        public static void RefreshPage(IWebDriver driver, string frame, string childFrame = null)
        {
            //Switch to main frame when it is visible
            driver.SwitchTo().DefaultContent();
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement f1 = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(frame)));
            if (childFrame != null)
            {
                driver.SwitchTo().Frame(f1);
                driver.SwitchTo().Frame(childFrame);
            }
            else
            {
                driver.SwitchTo().Frame(f1);
            }           
        }
    }
}
