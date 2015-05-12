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
using RTA.Automation.CRM.Pages.Investigations;

namespace RTA.Automation.CRM.Pages.Investigations
{

  [ActionClass]
    public class InvestigationGeneralCaseSearchPage : IFramePage
    {
        private static string frameId = "contentIFrame0";
        private static string pageTitle = "General Cases";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;

        public InvestigationGeneralCaseSearchPage(IWebDriver driver)
            : base(driver, InvestigationGeneralCaseSearchPage.frameId)
        {
            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

            //Switch to main frame when it is visible
            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);
        }

        /*
        * NewIMG
        * ************************************************************************
        */

        [ActionMethod]
        public void ClickNewGeneralCaseButton()
        {
            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickImageButton(driver, "New Case");          
        }
     }
}
