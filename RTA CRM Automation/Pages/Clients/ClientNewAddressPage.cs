using ActionWordsLib.Attributes;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using RTA.Automation.CRM.UI;
using RTA.Automation.CRM.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using AutoItX3Lib;

namespace RTA.Automation.CRM.Pages.Clients
{
    [ActionClass]
    class ClientNewAddressPage : IFramePage
    {
        //public static string WINDOW = "Client Address: New Client Address - Microsoft Dynamics CRM";
        private static string frameId = "contentIFrame0";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;
        private static string pageTitle = "Client Address";
        private static string FRAME = "InlineDialog_Iframe";

        public ClientNewAddressPage(IWebDriver driver)
            : base(driver, ClientNewAddressPage.frameId)
        {
            if (this.driver.FindElements(By.Id(FRAME)).Count > 0)
            {
                FirstRunDialogueFramePage firstRunDialogueFramePage = new FirstRunDialogueFramePage(this.driver);
                firstRunDialogueFramePage.ClickButtoncancel();
                this.driver.SwitchTo().DefaultContent();
            }
            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(50));
            wait.Until((d) => { return driver.Title.Contains(pageTitle); });

            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);

        }

        /*
        * Address Detail
        * ************************************************************************
        */
        [ActionMethod]
        public void SetAddressDetails(string AddressType, int RoadNumber, string RoadName)
        {
            
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#rta_address_detailid>div"))).Click();
            this.driver.FindElement(By.Id("rta_address_detailid_i")).Click();

            string BaseWindow = driver.CurrentWindowHandle;

            this.driver.FindElement(By.ClassName("ms-crm-InlineLookup-FooterSection-AddAnchor")).Click();

            Thread.Sleep(1000);
            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "Address Detail");
            
            ClientNewAddressDetailsPage newAddressDetailPage = new ClientNewAddressDetailsPage(driver);
            // click on title
            
            newAddressDetailPage.SetAddressType(AddressType);
            newAddressDetailPage.SetRoadNumber(Convert.ToString(RoadNumber));
            newAddressDetailPage.SetRoadName(RoadName);
            newAddressDetailPage.ClickPageTitle();
            newAddressDetailPage.ClickSaveAndClose();
            driver = driver.SwitchTo().Window(BaseWindow);
            this.ClickSaveAndClose();
        }

        [ActionMethod]
        public void ClickSaveAndClose()
        {
            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickSaveCloseButton(driver);
        }

        [ActionMethod]
        public void ClickSaveButton()
        {
            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickSaveButton(driver);
            this.driver.SwitchTo().Frame(frameId);
        }
        
        [ActionMethod]
        public void ClickPageTitle()
        {
            UICommon.ClickPageTitle(driver);
        }

        [ActionMethod]
        public void SetAddressDetails(string address)
        {
            UICommon.SetSearchableListValue("rta_address_detailid", address, driver);

        }


        [ActionMethod]
        public string GetCleintAddress()
        {
            return UICommon.GetNewReferenceNumber(driver);
        }
    }
}
