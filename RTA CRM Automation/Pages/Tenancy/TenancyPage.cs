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
using RTA.Automation.CRM.Pages.Clients;

namespace RTA.Automation.CRM.Pages
{
    [ActionClass]
    public class TenancyPage : IFramePage
    {
        private static string frameId = "contentIFrame0";
        private static string connectionsFRAME = "areaConnectionsFrame";
        private static string tenancyrequestFRAME = "rta_tenancy_rta_tenancy_requestFrame";
        private static string pageTitle = "Tenancy:";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;

        public TenancyPage(IWebDriver driver)
            : base(driver, TenancyPage.frameId)
        {
            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

            //Switch to main frame when it is visible
            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);

            //Ensure that this table has loaded first
            //GetContributorsTable();
            
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

        /*
       * SaveCloseIMG
       * ************************************************************************
       */

        [ActionMethod]
        public void ClickSaveCloseButton()
        {

            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickSaveCloseButton(driver);
            
            

        }

        /*
        * Connections
        * ************************************************************************
        *
        */

        [ActionMethod]
        public void ClickConnectionsElement()
        {

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//div/a/h2[contains(.,'Connections')]")));

            Actions action = new Actions(driver);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(2000);
            action.MoveToElement(elem).Release().Build().Perform();

        }

        public void ClickConnectionsAssociationsIMG()
        {

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Connections_openAssociatedGridViewImageButtonImage")));

            Actions action = new Actions(driver);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).Release().Build().Perform();

        }

        public void SetConnectList(string connectType)
        {

            Thread.Sleep(5000);
            
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            RefreshPageFrame.RefreshPage(driver, frameId);

            driver.SwitchTo().Frame(connectionsFRAME);


            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("associatedGridRibbon")));
            IWebElement elemList = elem.FindElement(By.Id("connection|OneToMany|SubGridAssociated|Mscrm.SubGrid.connection.AddConnection"));

            RefreshPageFrame.RefreshPage(driver, frameId);

            driver.SwitchTo().Frame(connectionsFRAME);
            
            IWebElement arrow = elemList.FindElement(By.XPath("//li/span/a/img[contains(@src,'CommandBarMenuDown.png')]"));//  //span/a/img[contains(@src,'/CommandBarMenuDown.png']"));
            Actions action = new Actions(driver);
            Thread.Sleep(2000);
            action.MoveToElement(arrow).ClickAndHold().Build().Perform();
            Thread.Sleep(2000);
            action.MoveToElement(arrow).Release().Build().Perform();
            
            IWebElement elemMenu = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("connection_OneToMany_SubGridAssociated_Mscrm_SubGrid_connection_AddConnectionMenu")));
            elemMenu.FindElement(By.XPath("//li/span/a[contains(.,'" + connectType + "')]")).Click();

            
        }

        
        /*
       * All Contributers
       * ************************************************************************
       *
       */

        [ActionMethod]
        public void ClickAllContributionsElement()
        {
          
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec)); 
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//div/a/h2[contains(.,'All Contributors')]")));
            
            Actions action = new Actions(driver);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).Release().Build().Perform();
            Thread.Sleep(2000);
            
        }

        /*
       * search table 
       * ************************************************************************
         * 
       */

        [ActionMethod]
        public IWebElement GetAllContributorsTable()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
    
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("AllContributors_gridBodyContainer")));
            elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("AllContributors_divDataArea")));
            IWebElement webElementBody = elem.FindElement(By.Id("gridBodyTable"));
         
            return webElementBody;
        }

        [ActionMethod]
        public IWebElement GetContributorsTable()
        {

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));

            //Wait for table to load
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//table[@name=\"Contributors\"]")));
            IWebElement webElementBody = elem.FindElement(By.Id("gridBodyTable"));
            
            return webElementBody;

        }

        /*
       * Validation Alert Message
       * ************************************************************************
       */
        [ActionMethod]
        public string GetRTAValidationMessage()
        {
            return UICommon.GetAlertMessage(driver);

        }

        /*
       * search criteria
       * ************************************************************************
       */

        [ActionMethod]
        public void SetTenancyConnectionSearchText(string searchValue)
        {
            
            RefreshPageFrame.RefreshPage(driver, frameId);
            driver.SwitchTo().Frame(connectionsFRAME);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("crmGrid_rta_tenancy_connections1_findCriteria")));
            IWebElement element = this.driver.FindElement(By.Id("crmGrid_rta_tenancy_connections1_findCriteria"));
            element.Clear();
            element.SendKeys(searchValue.ToString());
            element.SendKeys(Keys.Enter);
            
        }

        /*
        * search table 
        * ************************************************************************
        */

        [ActionMethod]
        public IWebElement GetSearchResultTable()
        {
            return UICommon.GetSearchResultTable(driver);
        }

        
        

        /*
        * tenancy Bond properties 
        * ************************************************************************
        */
        [ActionMethod]
        public void HoverBondPropertyRibbonTab()
        {
            driver.SwitchTo().DefaultContent();
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("TabNode_tab0Tab")));

            Actions action = new Actions(driver);
            action.MoveToElement(elem).Build().Perform();
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Node_nav_rta_tenancy_rta_tenancy_request")));



        }

        [ActionMethod]
        public void ClickBondTenancyRequestRibbonButton()
        {
           
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Node_nav_rta_tenancy_rta_tenancy_request")));
            Actions action = new Actions(driver);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).Release().Build().Perform();

            RefreshPageFrame.RefreshPage(driver, frameId);
       }

        /*
        * Bond Tenancy Request View
        * ************************************************************************
        */

        [ActionMethod]
        public void ClickSelectViewButton()
        {

            RefreshPageFrame.RefreshPage(driver, frameId);
            driver.SwitchTo().Frame(tenancyrequestFRAME);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[alt='Select a view']")));

            Actions action = new Actions(driver);
            action.MoveToElement(driver.FindElement(By.CssSelector("img[alt='Select a view']"))).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(driver.FindElement(By.CssSelector("img[alt='Select a view']"))).Release().Build().Perform();
            
        }

        public void SetViewList(string listview)
        {
           
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elemMenu = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//div[@id='Dialog_0']/div/ul")));
            elemMenu.FindElement(By.XPath("//li/a[contains(.,'" + listview + "')]")).Click();
           
        }


        internal IWebDriver SwitchNewBrowser(IWebDriver driver, string BaseWindow, string pageTitle)
        {
            return UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, pageTitle);
        }

        [ActionMethod]
        public void ClickPageTitle()
        {
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(frameId);
            UICommon.ClickPageTitle(driver);

        }

        [ActionMethod]
        public void SetResidentialTenancyTypeList(string type)
        {
            UICommon.SetSelectListValue("rta_tenancy_type", type, driver);

        }

        [ActionMethod]
        public void SetManagingPartyListValue(string party)
        {
            UICommon.SetSearchableListValue("rta_lessor_clientid", party, driver);
        }

        [ActionMethod]
        public void SetTenancyManagementTypeListValue(string RTATenancyManagementType)
        {
            UICommon.SetSelectListValue("rta_tenancy_management_type", RTATenancyManagementType, driver);
        }

        [ActionMethod]
        public void SetDwellingTypeListValue(string dwelling)
        {

            UICommon.SetSearchableListValue("rta_dwelling_typeid", dwelling, driver);
        }

        [ScenarioMethod]
        public void CreateNewAddress(string roadno, string roadname, string locality, string roomtype = "", string roomno = "", string complexunitno = "")
        {
            string BaseWindow = driver.CurrentWindowHandle;
            this.ClickNewRentalPremises(BaseWindow);

            ClientNewAddressDetailsPage addressDetailPage = new ClientNewAddressDetailsPage(driver);
            addressDetailPage.SetAddressType("Australian Physical");
            addressDetailPage.SetComplexUnitNumber(complexunitno);
            addressDetailPage.SetRoomType(roomtype);
            addressDetailPage.SetRoomNumber(roomno);
            addressDetailPage.SetRoadNumber(roadno);
            addressDetailPage.SetRoadName(roadname);
            addressDetailPage.SetLocality(locality);
            addressDetailPage.ClickSaveAndClose();
            driver.SwitchTo().Window(BaseWindow);
        }

        public void ClickNewRentalPremises(string BaseWindow)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_rental_premise_addressid")));
            Actions actions = new Actions(driver);
            actions.MoveToElement(elem).Build().Perform();
            Thread.Sleep(500);

            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#rta_rental_premise_addressid>div>div>img"))).Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_rental_premise_addressid_i"))).Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//div[@id='Dialog_rta_rental_premise_addressid_i_IMenu']//a[contains(@title,'Create a new Address Detail.')]"))).Click();
            Thread.Sleep(5000);

            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "Address Detail");
        }

        /*
        * Validation Warning Message
        * ************************************************************************
        */
        [ActionMethod]
        public String GetWarningMessage()
        {
            try
            {
                return UICommon.GetTextFromElement("#crmNotifications", driver);
            }
            catch (Exception)
            {
                return "";
            }
        }

        public string GetRentalPremises()
        {
            return UICommon.GetTextFromElement("#rta_rental_premise_addressid>div>span", driver);
        }

        public void ClearRentalPremisesValue()
        {
            string elementId = "rta_rental_premise_addressid";
            UICommon.ClearListBoxValue(elementId, driver);
        }

        public bool VerifyWarningMessagePresent(string cssSelector)
        {
            return UICommon.VerifyWarningMessagePresent(cssSelector, driver);
        }
    }
}
