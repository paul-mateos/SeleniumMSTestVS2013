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
    public class TaskPage : IFramePage
    {
        private static string frameId = "contentIFrame0";
        
        private static string pageTitle = "Task";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;

        public TaskPage(IWebDriver driver)
            : base(driver, TaskPage.frameId)
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
        public void ClickNewInvestigationCaseButton()
        {
            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickNewButton(driver);
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
        public void ClickMarkCompleteButton()
        {

            this.driver.SwitchTo().DefaultContent();

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[alt='Mark Complete']")));
            Actions action = new Actions(driver);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).Release().Build().Perform();
            Thread.Sleep(3000);
        }

        /*
        * Task Number
        * ************************************************************************
        */

        [ActionMethod]
        public string GetTaskNumber()
        {
            return UICommon.GetNewReferenceNumber(driver);
        }

      
      

        [ActionMethod]
        public void SetSelectSubjectValue(string subject)
        {

            //Switch to main frame when it is visible
            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);
            UICommon.SetSearchableListValue("rta_activity_subjectid", subject, driver);

        }
       
        public string GetSubjectValue()
        {
            return UICommon.GetTextFromElement("#subject>div>span", driver);
        }

        [ActionMethod]
        public void ClickPageTitle()
        {
            UICommon.ClickPageTitle(driver);
            
        }
    }
      



    
}
