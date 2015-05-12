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
    class ClientNewAddressDetailsPage : IFramePage
    {
        //public static string WINDOW = "Address Detail: New Address Detail - Microsoft Dynamics CRM";
        private static string frameId = "contentIFrame0";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;
        private static string pageTitle = "Address Detail";

        //protected IWebDriver driver = null;

        public ClientNewAddressDetailsPage(IWebDriver driver)
            : base(driver , ClientNewAddressDetailsPage.frameId)
        {

            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(50));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);

        }


        /* 
        * Address Type
        * ************************************************************************
        */

        private IWebElement GetAddressTypeElement()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            return this.driver.FindElement(By.Id("rta_type_i"));
            
        }

        [ActionMethod]
        public void SetAddressType(string AddressType)
        {
            SelectElement type = new SelectElement(this.GetAddressTypeElement());
            type.SelectByText(AddressType);
            //UICommon.SetSelectListValue("rta_type", AddressType, driver);
        }

        [ActionMethod]
        public String GetAddressType()
        {
            return this.GetAddressTypeElement().Text;
        }
        /* 
        * Click on Title
        * ************************************************************************
        */
        [ActionMethod]
        public void ClickPageTitle()
        {
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(frameId);
            UICommon.ClickPageTitle(driver);
        }
        
        /* 
        * Road number
        * ************************************************************************
        */

        //private void ClickRoadNumberTextBoxArea()
        //{
        //    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
        //    wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#rta_road_number>div>span"))).Click();
        //}

        [ActionMethod]
        public void SetRoadNumber(string RoadNumber)
        {
            //this.ClickRoadNumberTextBoxArea();
            //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            //wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_road_number_i"))).SendKeys(RoadNumber);

            UICommon.SetTextBoxValue("rta_road_number", RoadNumber, driver);
        }


        //[ActionMethod]
        //public String GetRoadNumber()
        //{
        //    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
        //    return wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_road_number_i"))).Text;
        //}
        
        /* 
        * Road name
        * ************************************************************************
        */

        //private void ClickRoadNameTextBoxArea()
        //{
        //    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
        //    wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#rta_road_name>div>span"))).Click();
        //}

        [ActionMethod]
        public void SetRoadName(string RoadName)
        {
            //this.ClickRoadNameTextBoxArea();
            //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            //wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_road_name_i"))).SendKeys(RoadName);

            UICommon.SetTextBoxValue("rta_road_name", RoadName, driver);
        }

        //[ActionMethod]
        //public String GetRoadName()
        //{
        //    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
        //    return wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_road_name_i"))).Text;
        //}
        
        [ActionMethod]
        public void SetComplexUnitNumber(string ComplexUnitNumber)
        {
            //this.ClickComplexUnitNumberTextBoxArea();
            //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            //wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_complex_unit_number_i"))).SendKeys(ComplexUnitNumber);

            UICommon.SetTextBoxValue("rta_complex_unit_number", ComplexUnitNumber, driver);
        }

        [ActionMethod]
        public void SetRoomNumber(string ComplexUnitNumber)
        {
            //this.ClickRoomNumberTextBoxArea();
            //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            //wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_room_site_number_i"))).SendKeys(ComplexUnitNumber);

            UICommon.SetTextBoxValue("rta_room_site_number", ComplexUnitNumber, driver);
        }

        //[ActionMethod]
        //public void ClickRoomNumberTextBoxArea()
        //{
        //    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
        //    wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#rta_room_site_number>div>span"))).Click();
        //}

        //[ActionMethod]
        //public void ClickComplexUnitNumberTextBoxArea()
        //{
        //    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
        //    wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#rta_complex_unit_number>div>span"))).Click();
        //}

        public void ClickSaveAndClose()
        {
            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickSaveCloseButton(driver);
        }

        public void SetLocality(string Locality)
        {
            UICommon.SetSearchableListValue("rta_localityid", Locality, driver);
        }

        public void SetRoomType(string roomtype)
        {
            if(!roomtype.Equals(""))
            {
                UICommon.SetSearchableListValue("rta_room_site_typeid", roomtype, driver);
            }       
        }
    }
}
