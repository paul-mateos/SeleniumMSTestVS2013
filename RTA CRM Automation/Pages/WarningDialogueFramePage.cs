using ActionWordsLib.Attributes;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using RTA.Automation.CRM.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace RTA.Automation.CRM.Pages
{
    
    public class WarningDialogueFramePage : IFramePage
    {
        private static string frameId = "InlineDialog1_Iframe";
        private static string frameId2 = "InlineDialog_Iframe";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;

        public WarningDialogueFramePage(IWebDriver driver)
            : base(driver)
        {
            new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec)).Until(ExpectedConditionsExtensions.ElementIsInvisible(By.Id("DialogLoadingDiv")));

            if (this.driver.FindElements(By.Id(WarningDialogueFramePage.frameId)).Count > 0)
            {
                this.driver.SwitchTo().Frame(frameId);
            }
        }

        /*
       * These methods are for the messages that appear on top of the home page
       * They might not be present  all the time.
       * The find element methods are surrounded by try catch so that the script doesn't fail when the elements are not found
       */

        /*
       * EmailWarningButton This button is always shown on top of FirstRunDialogueFramePageModel Buttoncancel and should be clicked first
       * ************************************************************************
       */


        
        public void ClickBeginButton()
        {
            try
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                IWebElement parent = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("crmDialogFooter")));
                IWebElement elem = parent.FindElement(By.Id("butBegin"));
                Actions action = new Actions(driver);
                action.MoveToElement(elem).ClickAndHold().Build().Perform();
                Thread.Sleep(2000);
                action.MoveToElement(elem).Release().Build().Perform();
            }
            catch
            {

            }

        }

        public void ClickProcessBeginButton()
        {
            try
            {
                RefreshPageFrame.RefreshPage(driver, frameId2);

                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                IWebElement parent = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("tdDialogFooter")));
                IWebElement elem = parent.FindElement(By.Id("butBegin"));
                Actions action = new Actions(driver);
                action.MoveToElement(elem).ClickAndHold().Build().Perform();
                Thread.Sleep(2000);
                action.MoveToElement(elem).Release().Build().Perform();
            }
            catch
            {

            }

        }
    }
}
