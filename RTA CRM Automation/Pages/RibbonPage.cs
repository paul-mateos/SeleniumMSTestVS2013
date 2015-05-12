using ActionWordsLib.Attributes;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RTA.Automation.CRM.Utils;
using System.Threading;
using RTA.Automation.CRM.UI;

namespace RTA.Automation.CRM.Pages
{
    public abstract class RibbonPage : BasePage
    {

        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;

        public RibbonPage(IWebDriver driver)
            : base(driver)
        {
            
        }

       
        [ActionMethod]
        public void HoverCRMRibbonTab()
        {
            
            UICommon.HoverRibbonTab("Tab1", driver); 
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("actionGroupControl_scrollableContainer")));

        }

        [ActionMethod]
        public void ClickCRMRibbonTab()
        {
            driver.SwitchTo().DefaultContent();
            UICommon.ClickRibbonTab("#Tab1", driver);
           
           
        }

       

        [ActionMethod]
        public void HoverClientServicesRibbonTab()
        {
            
            UICommon.HoverRibbonTab("TabSFA", driver);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("actionGroupControl_scrollableContainer")));
        }

        [ActionMethod]
        public void HoverSettingsRibbonTab()
        {

            
            UICommon.HoverRibbonTab("TabSettings", driver);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("actionGroupControl_scrollableContainer")));
        }

       
        [ActionMethod]
        public void HoverRBSRibbonTab()
        {
             
            UICommon.HoverRibbonTab("TabRBS", driver);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("actionGroupControl_scrollableContainer")));
           

        }
      
        [ActionMethod]
        public void HoverInvestigationsRibbonTab()
        {

            Thread.Sleep(3000);
            UICommon.HoverRibbonTab("TabCS", driver);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("actionGroupControl_scrollableContainer")));

        }

        /*
       * Client X Tab
       * ************************************************************************
       */
       
        [ActionMethod]
        public void HoverClientXRibbonTab(string clientName)
        {
            
            UICommon.HoverXRibbonTab(clientName, driver);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("actionGroupControl_scrollableContainer")));


        }

        [ActionMethod]
        public void HoverClientRibbonTab(string clientName)
        {
           
            UICommon.HoverXRibbonTab(clientName, driver);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("actionGroupControl_scrollableContainer")));
        }

        [ActionMethod]
        public void ClickClientXRibbonTab(string clientName)
        {
            driver.SwitchTo().DefaultContent();
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//span[@id='TabNode_tab0Tab-main' and contains(@title, '" + clientName + "')]")));  
            Actions action = new Actions(driver);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).Release().Build().Perform();
            Thread.Sleep(2000);
        }

        /*
        * Cleint X Ribbon Buttons
        * ************************************************************************
        */
        [ActionMethod]
        public void ClickClientXPhoneNumbersRibbonButton()
        {
            UICommon.ClickRibbonButton("Node_nav_rta_contact_client_rta_client_ph_no_clientid", driver);

             
        }

        [ActionMethod]
        public void ClickClientXClientNamesRibbonButton()
        {
           UICommon.ClickRibbonButton("Node_nav_rta_contact_rta_client_name_currently_known_asid", driver);

        }

        [ActionMethod]
        public void ClickClientXClientIdArtefactRibbonButton()
        {
            UICommon.ClickRibbonButton("Node_nav_rta_contact_rta_client_id_artefact_clientid", driver);

        }


        [ActionMethod]
        public void ClickClientXConnectionsRibbonButton()
        {
            
            UICommon.ClickRibbonButton("Node_navConnections", driver);
             
        }

        [ActionMethod]
        public void ClickClientXActivitiesRibbonButton()
        {

            UICommon.ClickRibbonButton("Node_navActivities", driver);

        }

        [ActionMethod]
        public void ClickClientXAuditRibbonButton()
        {

            UICommon.ClickRibbonButton("Node_navAudit", driver);

        }

        [ActionMethod]
        public void ClickClientXAddressesRibbonButton()
        {

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Node_nav_rta_contact_rta_client_address_clientid")));
            Actions action = new Actions(driver);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).Release().Build().Perform();
            Thread.Sleep(2000);

        }       

        

        /*
        * Rental Bond Services Button
        * ************************************************************************
        */
        [ActionMethod]
        public void ClickRBSRibbonButton()
        {
           UICommon.ClickRibbonButton("RBS", driver);
          
        }

        /*
        * Investigations Button
        * ************************************************************************
        */
        [ActionMethod]
        public void ClickInvestigationsRibbonButton()
        {
            UICommon.ClickRibbonButton("CS", driver);
        }

        [ActionMethod]
        public bool VerifyInvestigationsRibbonButtonPresent()
        {
            IReadOnlyCollection<IWebElement> ribbon = this.driver.FindElements(By.Id("CS"));
            if (ribbon.Count >= 1)
            {
                return true;
            }
            return false;
        }

        /*
        * Investigations Button
        * ************************************************************************
        */

        [ActionMethod]
        public void ClickInvestigationsCasesRibbonButton()
        {
            UICommon.ClickRibbonButton("rta_inv_case", driver);
          
        }

        [ActionMethod]
        public void ClickInvestigationsMasterCasesRibbonButton()
        {
            UICommon.ClickRibbonButton("rta_inv_master_case", driver);
        }

        [ActionMethod]
        public void ClickInvestigationsClientRibbonButton()
        {
            UICommon.ClickRibbonButton("nav_conts", driver);
        }

        [ActionMethod]
        public void ClickInvestigationsQueuesRibbonButton()
        {
            UICommon.ClickRibbonButton("nav_queues", driver);
        }

        [ActionMethod]
        public void ClickInvestigationsGeneralCasesRibbonButton()
        {
            UICommon.ClickRibbonButton("nav_cases", driver);

        }

        [ActionMethod]
        public void ClickInvestigatorXCasesRibbonButton()
        {

            UICommon.ClickRibbonButton("Node_nav_rta_systemuser_rta_inv_case_investigatorid", driver);

        }

        public void ClickInvestigationsPenaltyInfringementNoticesRibbonButton()
        {
            UICommon.ClickRibbonButton("rta_pin", driver);
        }


        [ActionMethod]
        public void HoverInvestigationXRibbonTab(string InvestigationID)
        {
            UICommon.HoverXRibbonTab(InvestigationID, driver);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("actionGroupControl_scrollableContainer")));
        }

        [ActionMethod]
        public void ClickInvestigationXCasePartiesButton()
        {
            UICommon.ClickRibbonButton("Node_nav_rta_inv_case_rta_case_party", driver);
        }
        /*
        * Client Services Button
        * ************************************************************************
        */
        [ActionMethod]
        public void ClickClientServicesRibbonButton()
        {
            UICommon.ClickRibbonButton("SFA", driver);
           
        }

        /*
        * Settings Button
        * ************************************************************************
        */
        [ActionMethod]
        public void ClickSettingsRibbonButton()
        {
            UICommon.ClickRibbonButton("Settings", driver);
        }

        /*
        * Clients Button
        * ************************************************************************
        */
        [ActionMethod]
        public void ClickClientsRibbonButton()
        {
            UICommon.ClickRibbonButton("nav_conts", driver);
        }

        /*
       * Client Activities Button
       * ************************************************************************
       */
        [ActionMethod]
        public void ClickClientActivitiesRibbonButton()
        {
            UICommon.ClickRibbonButton("nav_activities", driver);
        }

        /*
        * Processes Button
        * ************************************************************************
        */
        [ActionMethod]
        public void ClickProcessesRibbonButton()
        {
             UICommon.ClickRibbonButton("nav_workflow", driver);
        }

        /*
       * Alleged Offenses Button
       * ************************************************************************
       */
        [ActionMethod]
        public void ClickAllegedOffencesButton()
        {
            UICommon.ClickRibbonButton("rta_alleged_offence", driver);
        }

        /*
      * Ribbon Right Scroll Button
      * ************************************************************************
      */
        [ActionMethod]
        public void ClickRightScrollRibbonButton()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rightNavLink")));
            elem.Click();
            Thread.Sleep(2000);
            
        }



        /*
        * Rta Tenancy Request Button
        * ************************************************************************
        */
        [ActionMethod]
        public void ClickRequestBatchRibbonButton()
        {

           UICommon.ClickRibbonButton("rta_request_batch", driver);
        }

        /*
       * Rta Tenancy Request Button
       * ************************************************************************
       */
        [ActionMethod]
        public void ClickRtaTenancyRequestRibbonButton()
        {

            UICommon.ClickRibbonButton("rta_tenancy_request", driver);

        }

        /*
        * Rta Tenancy Button
        * ************************************************************************
        */
        [ActionMethod]
        public void ClickRtaTenancyRibbonButton()
        {

            UICommon.ClickRibbonButton("rta_tenancy", driver);
            
        }

        /*
      * Rta Address Detail Button
      * ************************************************************************
      */
        [ActionMethod]
        public void ClickRtaAddressDetailRibbonButton()
        {

            UICommon.ClickRibbonButton("rta_address_detail", driver);

        }

        /*
      * Rta Queues Button
      * ************************************************************************
      */
        [ActionMethod]
        public void ClickRtaQueuesButton()
        {
            UICommon.ClickRibbonButton("nav_queues", driver);
        }

       /*
      * Tenancy Request Tab
      * ************************************************************************
      */

        [ActionMethod]
        public void HoverTRRibbonTab(string tenancyRequest)
        {
            
            UICommon.HoverXRibbonTab(tenancyRequest, driver);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("actionGroupControl_scrollableContainer")));
            
        }



        [ActionMethod]
        public void ClickAuditRibbonButton()
        {
            
            UICommon.ClickRibbonButton("Node_navAudit", driver);

        }

        /***Address Detail Record Tab
        *****************************************************
        */
        [ActionMethod]
        public void HoverAddressDetailRibbonTab(string addressDetail)
        {
            UICommon.HoverXRibbonTab(addressDetail, driver);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("actionGroupControl_scrollableContainer")));
        }

        [ActionMethod]
        public void ClickTRAddressDetailViewRibbonButton()
        {
            UICommon.ClickRibbonButton("Node_nav_rta_address_detail_rta_tenancy_request", driver);
        }

        [ActionMethod]
        public void ClickCreateIMG()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementExists(By.Id("navTabGlobalCreateImage")));
            Actions action = new Actions(driver);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).Release().Build().Perform();
            Thread.Sleep(2000);

        }

        [ActionMethod]
        public void ClickCreatePhoneActivityRibbonButton()
        {
            UICommon.ClickRibbonButton("4210", driver);

        }

        /*
        *    Front Counter Contact Activity
        *****************************************************
        */
        [ActionMethod]
        public void ClickFrontCounterContactActivityRibbonButton()
        {
            // UICommon.ClickRibbonButton("10049", driver);
            UICommon.ClickRibbonTab("a[title='Front Counter Contact']", driver);
        }

        [ActionMethod]
        public void HoverFrontCounterContactXRibbonTab(string ActivityName)
        {
            UICommon.HoverXRibbonTab(ActivityName, driver);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("actionGroupControl_scrollableContainer")));
        }

        [ActionMethod]
        public void ClickFrontCounterContactXConnectionsButton()
        {
            UICommon.ClickRibbonButton("Node_navConnections", driver);
        }
    }
}
