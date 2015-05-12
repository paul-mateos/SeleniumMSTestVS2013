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
using System.Collections.ObjectModel;
using RTA.Automation.CRM.UI;


namespace RTA.Automation.CRM.Pages
{
    [ActionClass]
    public class RequestBatchPage : IFramePage
    {
        private static string frameId = "";
        private static string FRAMEpaymentReference = "rta_request_batch_rta_payment_referenceFrame";
        //private static string FRAMErequestParty = "rta_tenancy_request_rta_tenancy_request_partyFrame";

        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;
        private static string pageTitle = "Request Batch:";

        public RequestBatchPage(IWebDriver driver)
            : base(driver, RequestBatchPage.frameId)
        {

           
            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

            
            //Switch to main frame when it is visible
            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);
            

        }

        


        /*
        * ManagingPartyElement
        * ************************************************************************
        */

        [ActionMethod]
        public void SetManagingPartyText(string party)
        {
            UICommon.SetSearchableListValue("rta_managing_partyid", party, driver);
        }



        /*
        * RequestNumber
        * ************************************************************************
        */

        [ActionMethod]
        public string GetRequestNumber()
        {

            return UICommon.GetNewReferenceNumber(driver);
        }

        /*
        * SaveIMG
        * ************************************************************************
        */

        [ActionMethod]
        public void ClickSaveButton()
        {

            driver.SwitchTo().DefaultContent();
            UICommon.ClickSaveButton(driver);
            driver.SwitchTo().Frame(frameId);

 
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
      * Save Unsaved Changes
      * ************************************************************************
      */

        [ActionMethod]
        public void ClickUnsavedChangesButton()
        {

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("savefooter_statuscontrol")));
            
            Actions action = new Actions(driver);
            Thread.Sleep(2000);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(2000);
            action.MoveToElement(elem).Release().Build().Perform();
            Thread.Sleep(2000);
            WaitForPageToLoad.WaitToLoad(driver);
    
        }
        

        /*
        * NewIMG
        * ************************************************************************
        */

        [ActionMethod]
        public void ClickNewTenancyRequestButton()
        {
            UICommon.ClickNewButton(driver);
          
        }


        /*
        * Validation Alert Message
        * ************************************************************************
        */
        [ActionMethod]
        public string GetAlertMessage()
        {
            return UICommon.GetAlertMessage(driver);
        }

        /*
        * Validation Warning Message
        * ************************************************************************
        */
        [ActionMethod]
        public String GetWarningMessage()
        {

            return UICommon.GetTextFromElement("#crmNotifications", driver);
   

        }
        /*
        * Funded Status
        * ************************************************************************
        */
        [ActionMethod]
        public string GetPropertyDataControlModeRTAFundedStatus()
        {
            return UICommon.GetElementProperty("#rta_funded_status", "data-controlmode", driver);
            
        }



        [ActionMethod]
        public string GetFundedStatus()
        {
            RefreshPageFrame.RefreshPage(driver, frameId);
            return UICommon.GetTextFromElement("#rta_funded_status", driver);

        }

        /*
       * Open Payment Reference Associated
       * ************************************************************************
       */
        [ActionMethod]
        public void ClickPaymentRefOpenAssociated()
        {
            
            RefreshPageFrame.RefreshPage(driver, frameId); 
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Subgrid_PaymentReferences_openAssociatedGridViewImageButtonImage"))).Click();
            
        }

        /*
       * Add associated requests
       * ************************************************************************
       */
        [ActionMethod]
        public void ClickAddAssociatedRequestsButton()
        {
            RefreshPageFrame.RefreshPage(driver, frameId);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Subgrid_TenancyRequests_addImageButtonImage"))).Click();
            
        }
        

        /*
      * Open Request Party Associated
      * ************************************************************************
      */
        [ActionMethod]
        public void ClickRequestPartyAssociated()
        {     
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("SubGrid_TenancyRequestParties_openAssociatedGridViewImageButtonImage"))).Click();
            
        }

        /*
        * Add New Payment Reference
        * ************************************************************************
        */
        [ActionMethod]
        public void ClickAddNewPaymentRefImage()
        {
            RefreshPageFrame.RefreshPage(driver, frameId);
            this.driver.SwitchTo().Frame(FRAMEpaymentReference);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[alt='Add New Payment Reference']"))).Click();
            WaitForPageToLoad.WaitToLoad(driver);
            this.driver.SwitchTo().DefaultContent();

        }

        /*
        * Add New Request Party Reference
        * ************************************************************************
        */
        [ActionMethod]
        public void ClickAddNewRequestPartyImage()
        {
            //this.driver.SwitchTo().Frame(FRAMErequestParty);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("SubGrid_TenancyRequestParties_addImageButtonImage"))).Click();
            WaitForPageToLoad.WaitToLoad(driver);
           

        }

        /*
       * search criteria
       * ************************************************************************
       */

        [ActionMethod]
        public void SetTenancyRequestPaymentSearchText(string searchValue)
        {
            UICommon.SetSearchText("crmGrid_findCriteria", searchValue, driver);

        }


        /*
        * search table 
        * ************************************************************************
        */

        [ActionMethod]
        public IWebElement GetPaymentSearchResultTable()
        {
            return UICommon.GetSearchResultTable(driver);
        }

        /*
        * search table 
        * ************************************************************************
        */

        [ActionMethod]
        public IWebElement GetPaymentSummaryResultTable()
        {

            RefreshPageFrame.RefreshPage(driver, frameId);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement parent = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Subgrid_PaymentReferences_divDataArea")));

            IWebElement webElementBody = parent.FindElement(By.Id("gridBodyTable"));
            return webElementBody;
        }


       
        /*
        * Status Reason
        * ************************************************************************
        */
        
      
        [ActionMethod]
        public string GetValidationStatusReason()
        {
            RefreshPageFrame.RefreshPage(driver, frameId);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = null;

            for (int i = 0; i < 30; i++)
            {
            
                elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#statuscode>div>span")));
                if (elem.Text == "Ready for validation")
                {
                    Thread.Sleep(1000);
                }
                else
                {
                    break;
                }
            }
            WaitForPageToLoad.WaitToLoad(driver);
            return elem.Text;
        }

       
        [ActionMethod]
        public void SetStatusReason(string reason)
        {
            UICommon.SetSelectListValue("statuscode", reason, driver);
        }

        /*
       * Payment Type
       * ************************************************************************
       */

        [ActionMethod]
        public void SetPaymentType(string type)
        {
            UICommon.SetSelectListValue("rta_payment_type", type, driver);

        }

        /*
      * Associated Request
      * ************************************************************************
      */

        [ActionMethod]
        public void SetAssociatedRequest(string reference)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("lookup_Subgrid_TenancyRequests_ledit")));
            elem.Click();
            elem.Clear();
            elem.SendKeys(reference);
            elem.SendKeys(Keys.Enter);
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//li[a[contains(@title,'" + reference + "')] and parent::*[@id='lookup_Subgrid_TenancyRequests_i_IMenu']]"))).Click();
            Thread.Sleep(2000);

        }

        /*
       * Amount Matched
       * ************************************************************************
       */
        [ActionMethod]
        public string GetAmountMatched()
        {
            RefreshPageFrame.RefreshPage(driver, frameId);
            return UICommon.GetTextFromElement("#rta_amount_matched>div>span", driver);
        }
                
        [ActionMethod]
        public IWebElement GetTenancyRequestTable()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement parent = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Subgrid_TenancyRequests_divDataArea")));

            IWebElement webElementBody = parent.FindElement(By.Id("gridBodyTable"));
            return webElementBody;
        }

        public string GetDialogErrorMessageText()
        {
            driver.SwitchTo().DefaultContent();
            RefreshPageFrame.RefreshPage(driver, "InlineDialog_Iframe");
            return UICommon.GetTextFromElement("#ErrorMessage", driver);
            
        }

        public void ClickErrorMessageOkButton()
        {
            driver.SwitchTo().DefaultContent();
            RefreshPageFrame.RefreshPage(driver, "InlineDialog_Iframe");
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("butBegin")));
            elem.Click();
        }

        internal string GetStatusReason()
        {
            return UICommon.GetTextFromElement("#statuscode>div>span", driver);
        }

        internal void ClickPageTitle()
        {

            UICommon.ClickPageTitle(driver);
        }

        public string GetSumBondamountPaid()
        {
            return UICommon.GetTextFromElement("#rta_sum_amount_bond_paid", driver);
        }
        
        [ActionMethod]
        public IWebElement GetAssociatedTenancyRequestTable()
        {
            RefreshPageFrame.RefreshPage(driver, frameId);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement parent = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Subgrid_PaymentReferences_divDataArea")));

            IWebElement webElementBody = parent.FindElement(By.Id("gridBodyTable"));
            return webElementBody;
        }


        [ActionMethod]
        public string GetRecordStatus()
        {
            return UICommon.GetTextFromElement("#titlefooter_statuscontrol", driver);
        }
    }
}
