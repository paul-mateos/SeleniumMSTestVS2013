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


namespace RTA.Automation.CRM.Pages
{
    [ActionClass]
    public class CasePartyPage : IFramePage
    {
        private static string frameId = "contentIFrame0";    
        private static string dialogFRAME = "InlineDialog_Iframe";
        private static string pageTitle = "Case Party:";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;

        public CasePartyPage(IWebDriver driver)
            : base(driver, CasePartyPage.frameId)
        {
            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

            //Switch to main frame when it is visible
            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);
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
            this.driver.SwitchTo().Frame(frameId);
        }

        [ActionMethod]
        public void ClickSaveCloseButton()
        {
            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickSaveCloseButton(driver);
        }

        [ActionMethod]
        public void SetClientName(string clientName)
        {
            ////Switch to main frame when it is visible
            //frameId = UICommon.FindVisibleIFrame(driver);
            //RefreshPageFrame.RefreshPage(driver, frameId);
            UICommon.SetSearchableListValue("rta_clientid", clientName, driver);
        }
        [ActionMethod]
        public void ClickPageTitle()
        {
            UICommon.ClickPageTitle(driver);
        }

        [ActionMethod]
        public void ClickStartDialog()
        {
            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickImageButton(driver, "Start Dialog");
            this.driver.SwitchTo().Frame(frameId);
        }

        [ActionMethod]
        public IWebElement GetProcessSearchResultTable()
        {
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(dialogFRAME);
            return UICommon.GetSearchResultTable(driver);
        }

        [ActionMethod]
        public void ClickDialogAddButton()
        {
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(dialogFRAME);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("butBegin"))).Click();
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(frameId);
        }

        [ActionMethod]
        public bool VerifyCasePartyTypeExists(string PartyType)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_party_type_i")));

            IReadOnlyCollection<IWebElement> selectors =  elem.FindElements(By.CssSelector("Option"));
            foreach (IWebElement type in selectors)
            { 
                if (type.Text.Equals(PartyType))
                {
                    return true;
                }
            }
            return false;
        }

        [ActionMethod]
        public void ClickPartyType()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#rta_party_type>div")));
            elem.Click();
        }
    }
}
