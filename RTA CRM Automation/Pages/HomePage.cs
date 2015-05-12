using System.Collections.ObjectModel;
using System.Linq;
using System;
using ActionWordsLib.Attributes;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using System.Collections.Generic;
using RTA.Automation.CRM.UI;




namespace RTA.Automation.CRM.Pages
{
    [ActionClass]
    public class HomePage : RibbonPage
    {
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;

        public HomePage(IWebDriver driver)
            : base(driver)
        {
            WarningDialogueFramePage warningDialogueFramePage = new WarningDialogueFramePage(this.driver);
            warningDialogueFramePage.ClickBeginButton();
            warningDialogueFramePage = new WarningDialogueFramePage(this.driver);
            warningDialogueFramePage.ClickBeginButton();
            FirstRunDialogueFramePage firstRunDialogueFramePage = new FirstRunDialogueFramePage(this.driver);
            firstRunDialogueFramePage.ClickButtoncancel();
            this.driver.SwitchTo().DefaultContent();
        }

       

        /*
        * AdvancedfindIMG
        * ************************************************************************
        */
       

        [ActionMethod]
        public void ClickAdvancedfindIMG()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementExists(By.CssSelector("img[alt='Advanced Find']"))).Click();
         
        }

        

        

        

        /*
        * ButtonClients
        * ************************************************************************
        */

        public IWebElement GetButtonClientsElement()
        {

            this.driver.SwitchTo().Frame("contentIFrame0");
            ReadOnlyCollection<IWebElement> elements = this.driver.FindElements(By.Id("nav_conts"));
            IWebElement element = elements.First();
            return element;
        }

        [ActionMethod]
        public void ClickButtonClients()
        {
            IWebElement element = this.GetButtonClientsElement();
            element.Click();
            this.driver.SwitchTo().DefaultContent();
        }

        /*
        * New Activity
        * ************************************************************************
        */
        [ActionMethod]
        public void ClickNewActivityIMG()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[alt='New Activity']"))).Click();

        }

        [ActionMethod]
        public void ClickEmailIMG()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementExists(By.CssSelector("img[alt='Email']"))).Click();

        }

        [ActionMethod]
        public void ClickFrontCounterContactIMG()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementExists(By.CssSelector("img[alt='Front Counter Contact']"))).Click();

        }

        [ActionMethod]
        public bool CheckNewActivityContents(String activity)
        {
            driver.FindElement(By.CssSelector("img[alt='New Activity']")).Click();
            IWebElement newActivity = driver.FindElement(By.ClassName("ui-flyout-dialog-moreCommands"));
            IList<IWebElement> activityList = newActivity.FindElements(By.TagName("li"));
            foreach (IWebElement item in activityList)
            {
                string title=item.GetAttribute("Title");
                string[] value = title.Split(' ');
                if (activity.Equals(value[0]))
                {
                    return true;
                }
            }
            return false;

        }

        /*
       * Investigations / Client Tile
       * ************************************************************************
       */
        [ActionMethod]
        public void ClickInvestigationsClientTile()
        {
            UICommon.ClickHomePageTile("/_imgs/NavBar/ActionImgs/Contact_32.png", driver);

            //this.driver.SwitchTo().Frame("contentIFrame0");
            
            //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            //wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//ul[@id='uxMenu']/li[3]")));
            //IWebElement elem = driver.FindElement(By.XPath("//ul[@id='uxMenu']/li[3]"));
            //driver.FindElement(By.XPath("*//img[contains(@src,'/_imgs/NavBar/ActionImgs/Contact_32.png')]")).Click();
            
            //this.driver.SwitchTo().DefaultContent();
        }

        [ActionMethod]
        public void ClickInvestigationsCaseTile()
        {
            UICommon.ClickHomePageTile("/_imgs/NavBar/ActionImgs/Cases_32.png", driver);

            //this.driver.SwitchTo().Frame("contentIFrame0");
            
            //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            //wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//ul[@id='uxMenu']/li[3]")));
            //IWebElement elem = driver.FindElement(By.XPath("//ul[@id='uxMenu']/li[3]"));
            //driver.FindElement(By.XPath("*//img[contains(@src,'/_imgs/NavBar/ActionImgs/Cases_32.png')]")).Click();
            
            //this.driver.SwitchTo().DefaultContent();
        }
    }
}
