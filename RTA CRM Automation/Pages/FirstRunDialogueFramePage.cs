using ActionWordsLib.Attributes;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using RTA.Automation.CRM.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RTA.Automation.CRM.Pages
{
    
    public class FirstRunDialogueFramePage : IFramePage
    {

        private static string FRAME = "InlineDialog_Iframe";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;


        public FirstRunDialogueFramePage(IWebDriver driver)
            : base(driver)
        {
            

            if (this.driver.FindElements(By.Id(FirstRunDialogueFramePage.FRAME)).Count > 0)
            {
                this.frame = FirstRunDialogueFramePage.FRAME;
                this.driver.SwitchTo().Frame(this.frame);
            }
        }

        /*
        * These methods are for the messages that appear on top of the home page
        * They might not be present  all the time.
        * The find element methods are surrounded by try catch so that the script doesn't fail when the elements are not found
        */

        
        public void ClickButtoncancel()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement parent = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("navTourPage1")));
            parent.FindElement(By.Id("buttonCancel")).Click();
        }

      
    }
}
