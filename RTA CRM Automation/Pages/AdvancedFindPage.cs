using OpenQA.Selenium;
using System.Collections.ObjectModel;
using System.Linq;
using System;
using ActionWordsLib.Attributes;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System.Threading.Tasks;
using System.Threading;
using RTA.Automation.CRM.Utils;
using RTA.Automation.CRM.UI;

namespace RTA.Automation.CRM.Pages
{
    [ActionClass]
    class AdvancedFindPage : IFramePage
    {
       private static string frameId = "contentIFrame0";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;
        private static string pageTitle = "Advanced Find";

        public AdvancedFindPage(IWebDriver driver)
            : base(driver, AdvancedFindPage.frameId)
        {
            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);
        }

        public void SelectLookForListItem(string LookForValue)
        {
            IWebElement select = driver.FindElement(By.Id("slctPrimaryEntity"));
            ReadOnlyCollection<IWebElement> options = select.FindElements(By.TagName("option"));
            foreach (IWebElement option in options)
            {
                if (LookForValue.Equals(option.Text))
                {
                    option.Click();
                }
            }
            throw new Exception("Look for Item not found!!!");
        }

        public bool VerifyLookForListItemPresent(string LookForValue)
        { 
            IWebElement select = driver.FindElement(By.Id("slctPrimaryEntity"));
            ReadOnlyCollection<IWebElement> options = select.FindElements(By.TagName("option"));
            foreach (IWebElement option in options) 
            {
                if(LookForValue.Equals(option.Text))
                {
                    return true;
                }
            }
            return false;
        }

        public void CloseWindow()
        {
            driver.Close();
        }
    }
}
