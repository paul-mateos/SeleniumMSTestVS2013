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
using RTA.Automation.CRM.Pages.Clients;


namespace RTA.Automation.CRM.Pages
{
    [ActionClass]
    public class TenancyRequestPage : IFramePage
    {
        private static string frameId = "contentIFrame0";
        private static string dialogFRAME = "InlineDialog_Iframe";
        private static string FRAMEpaymentReference = "rta_tenancy_request_rta_payment_referenceFrame";
        //private static string FRAMErequestParty = "rta_tenancy_request_rta_tenancy_request_partyFrame";
        
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;
        private static string pageTitle = "Tenancy Request";

        public TenancyRequestPage(IWebDriver driver)
            : base(driver, TenancyRequestPage.frameId)
        {

           
            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

            //Switch to main frame when it is visible
            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);           
            


        }

        [ScenarioMethod]
        public void PopulateTenancyRequestFormResidentialTenancy(string requestType, string rentalPremises, string managingParty, string tenancyType, string managementType, string numBedrooms, string initailRequestParty, string initialControbution, string amountPaid, string lodgementType)
        {
            
            this.SetRequestTypeListValue(requestType);
            this.SetRentalPremisesValue(rentalPremises);
            this.SetManagingPartyListValue(managingParty); 
            this.SetResidentialTenancyTypeList(tenancyType);
            this.SetResidentialManagementTypeList(managementType);
            this.SetNumberOfBedrooms(numBedrooms);
            this.SetInitialRequestPartyWithSearch(initailRequestParty);
            this.SetInitialConrtibution(initialControbution);
            this.SetAmountPaid(amountPaid);
            this.SetLodgementTypeListValue(lodgementType);
        }

        [ActionMethod]
        public void SetResidentialTenancyTypeList(string type)
        {
            UICommon.SetSelectListValue("rta_tenancy_type", type, driver);

        }

        [ActionMethod]
        public void SetResidentialManagementTypeList(string type)
        {
            UICommon.SetSelectListValue("rta_tenancy_management_type", type, driver);

        }

        [ScenarioMethod]
        public void PopulateTenancyRequestFormTopUpResidentialTenancy(string requestType, string rentalPremises, string managingParty, string tenancy, string managementType, string tenancyType, string initailRequestParty, string initialControbution, string amountPaid, string lodgementType, string tenancyStart, string numBedrooms)
        {

            this.SetRequestTypeListValue(requestType);
            this.SetRentalPremisesValue(rentalPremises);
            this.SetManagingPartyListValue(managingParty);
            this.SetTenancyValue(tenancy);
            this.SetResidentialTenancyTypeList(tenancyType);
            this.SetResidentialManagementTypeList(managementType);
            this.SetInitialRequestPartyWithSearch(initailRequestParty);
            this.SetInitialConrtibution(initialControbution);
            this.SetAmountPaid(amountPaid);
            this.SetLodgementTypeListValue(lodgementType);
            this.SetTenancyStartDate(tenancyStart);
            this.SetNumberOfBedrooms(numBedrooms);
            
        }

        [ScenarioMethod]
        public void PopulateTenancyRequestTopUpValidationSuccessful(string requestType, string rentalPremises, string managingParty, string tenancyType, string managementType, string numBedrooms, string initailRequestParty, string weeklyRent, string initialControbution, string amountPaid, string lodgementType, string tenancyStart, string anticipatedEnd, string paymentType, string bondRef)
        {
            this.SetRequestTypeListValue(requestType);
            this.SetRentalPremisesValue(rentalPremises);
            this.SetManagingPartyListValue(managingParty);
            this.SetResidentialTenancyTypeList(tenancyType);
            this.SetResidentialManagementTypeList(managementType);
            this.SetNumberOfBedrooms(numBedrooms);
            this.SetInitialRequestPartyWithSearch(initailRequestParty);
            this.SetWeeklyRent(weeklyRent);
            this.SetInitialConrtibution(initialControbution);
            this.SetAmountPaid(amountPaid);
            this.SetLodgementTypeListValue(lodgementType);
            this.SetTenancyStartDate(tenancyStart);
            this.SetPaymentType(paymentType);
            this.SetAnticipatedEndDate(anticipatedEnd);
            this.SetTenancyValue(bondRef);     
        }

        [ScenarioMethod]
        public void PopulateTenancyRequestFormBondLodgement(string requestType, string rentalPremises, string managingParty, 
            string tenancyType, string managementType, string weeklyRent, string initialRequestParty, string initialContribution, 
            string amountPaid, string lodgementType, string tenancyStart, string tenancyEnd, string paymentType, string subsidy)
        {
            this.SetRequestTypeListValue(requestType);
            this.SetRentalPremisesValue(rentalPremises);
            this.SetManagingPartyListValue(managingParty);
            this.SetTenancyTypeListValue(tenancyType);
            this.SetResidentialManagementTypeList(managementType);
            this.SetWeeklyRent(weeklyRent);
            this.SetInitialRequestPartyWithSearch(initialRequestParty); 
            this.SetInitialConrtibution(initialContribution);
            this.SetAmountPaidWithLodgement(amountPaid);
            this.SetLodgementTypeListValue(lodgementType);
            this.SetTenancyStartDate(tenancyStart);
            this.SetAnticipatedEndDate(tenancyEnd);
            this.SetPaymentType(paymentType);
            this.SetSubsidy(subsidy);
        }

        [ScenarioMethod]
        public void CreateNewClient(string clientName)
        {
            this.ClickNewClientButton();
            Thread.Sleep(5000);

            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);
       
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            UICommon.SetTextBoxValue("lastname", clientName, driver);
                             
            this.driver.SwitchTo().DefaultContent();
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement saveButton = wait.Until(ExpectedConditions.ElementExists(By.CssSelector("#globalquickcreate_save_button_rta_managing_partyid_i_lookup_quickcreate")));
            
            Actions action = new Actions(driver);
            action.MoveToElement(saveButton).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(saveButton).Release().Build().Perform();
            Thread.Sleep(3000);
        }

        [ActionMethod]
        public void ClickNewClientButton()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_managing_partyid")));
            Actions actions = new Actions(driver);
            actions.MoveToElement(elem).Build().Perform();
            Thread.Sleep(500);             

            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#rta_managing_partyid>div>div>img"))).Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_managing_partyid_i"))).Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//div[@id='Dialog_rta_managing_partyid_i_IMenu']//a[contains(@title,'Create a new Client.')]"))).Click();
         }

        [ScenarioMethod]
        public string CreateNewBatchRequest(string managingParty,string paymentType)
        {
            this.ClickNewRequestBatch();

            RequestBatchPage requestBatchPage = new RequestBatchPage(driver);
            requestBatchPage.ClickPageTitle();
                        
            requestBatchPage.SetManagingPartyText(managingParty);
            requestBatchPage.SetPaymentType(paymentType);
            requestBatchPage.ClickSaveButton();
            string requestBatch = requestBatchPage.GetRequestNumber();
            
            requestBatchPage.ClickSaveCloseButton();
            
            WebDriverWait waitHandle = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            waitHandle.Until((d) => { return d.WindowHandles.Count == 1; });

            driver.SwitchTo().Window(driver.WindowHandles[0]);
            return requestBatch;
        }

        [ActionMethod]
        public void ClickNewRequestBatch()
       {
           WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
           IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_request_batchid")));
           Actions actions = new Actions(driver);
           actions.MoveToElement(elem).Build().Perform();
           Thread.Sleep(500);
           int count = driver.WindowHandles.Count;
           
           wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#rta_request_batchid>div>div>img"))).Click();
           wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_request_batchid_i"))).Click();
           wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//div[@id='Dialog_rta_request_batchid_i_IMenu']//a[contains(@title,'Create a new Request Batch.')]"))).Click();


           WebDriverWait waitHandle = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
           waitHandle.Until((d)=> { return d.WindowHandles.Count > count; });
       
           List<string> windowHandles = new List<String>(driver.WindowHandles);
           foreach (string eachHandle in windowHandles)
           {
               if(!driver.CurrentWindowHandle.Equals(eachHandle))
               {
                   driver.SwitchTo().Window(eachHandle);
               }
           }
       }

        [ActionMethod]
        public void SetLodgementTypeListValue(string lodgementType)
        {
               UICommon.SetSelectListValue("rta_lodgement_type", lodgementType, driver);

        }

  
        [ActionMethod]
        public void SetTenancyValue(string tenancy)
        {
            UICommon.SetSearchableListValue("rta_tenancyid", tenancy, driver);
        }


        [ScenarioMethod]
        public void PopulateTRNoTenancyType(string requestType, string rentalPremises, string managingParty, string tenancyType, string managementType, string numBedrooms, string initailRequestParty, string initialControbution, string amountPaid, string lodgementType)
        {
            
            this.SetRequestTypeListValue(requestType);
            this.SetRentalPremisesValue(rentalPremises);
            this.SetManagingPartyListValue(managingParty);
            this.SetResidentialManagementTypeList(managementType);
            this.SetNumberOfBedrooms(numBedrooms);
            this.SetInitialRequestPartyWithSearch(initailRequestParty);
            this.SetInitialConrtibution(initialControbution);
            this.SetAmountPaid(amountPaid);
            this.SetLodgementTypeListValue(lodgementType);
        }

        [ScenarioMethod]
        public void PopulateTenancyRequestFormRoomingAccomodation(string requestType, string rentalPremises, string managingParty, string tenancyType, string managementType, string initailRequestParty, string initialControbution, string amountPaid, string lodgementType)
        {
            this.SetRequestTypeListValue(requestType);
            this.SetRentalPremisesValue("*"+rentalPremises);
            this.SetManagingPartyListValue(managingParty);
            this.SetResidentialTenancyTypeList(tenancyType);
            this.SetResidentialManagementTypeList(managementType);
            this.SetInitialRequestPartyWithSearch(initailRequestParty);
            this.SetInitialConrtibution(initialControbution);
            this.SetAmountPaid(amountPaid);
            this.SetLodgementTypeListValue(lodgementType);
            
        }

        [ScenarioMethod]
        public void PopulateTenancyRequestRoomingAccomodationWithNoRentalPremises(string requestType, string managingParty, string tenancyType, string managementType, string initailRequestParty, string initialControbution, string amountPaid, string lodgementType)
        {
            this.SetRequestTypeListValue(requestType);
            this.SetManagingPartyListValue(managingParty);
            this.SetResidentialTenancyTypeList(tenancyType);
            this.SetResidentialManagementTypeList(managementType);
            this.ClickInitialRequestPartyListValue(initailRequestParty);
            this.SetInitialConrtibution(initialControbution);
            this.SetAmountPaid(amountPaid);
            this.SetLodgementTypeListValue(lodgementType);
        }

        [ScenarioMethod]
        public void PopulateTenancyRequestValidationSuccessful(string requestType, string rentalPremises, string managingParty, string tenancyType, string managementType, string numBedrooms, string initailRequestParty, string weeklyRent, string initialControbution, string amountPaid, string lodgementType, string tenancyStart, string anticipatedEnd, string paymentType)
        {

            this.SetRequestTypeListValue(requestType);
            Thread.Sleep(1000);
            this.ClickPageTitle();
            Thread.Sleep(1000);
            this.SetRentalPremisesValue("*"+rentalPremises);
            Thread.Sleep(1000);
            this.ClickPageTitle();
            Thread.Sleep(1000); 
            this.SetManagingPartyListValue(managingParty);
            Thread.Sleep(1000);
            this.ClickPageTitle();
            Thread.Sleep(1000); 
            this.SetResidentialTenancyTypeList(tenancyType);
            Thread.Sleep(1000);
            this.ClickPageTitle();
            Thread.Sleep(1000); 
            this.SetResidentialManagementTypeList(managementType);
            Thread.Sleep(1000);
            this.ClickPageTitle();
            Thread.Sleep(1000); 
            this.SetNumberOfBedrooms(numBedrooms);
            Thread.Sleep(1000);
            this.ClickPageTitle();
            Thread.Sleep(1000);
            this.SetInitialRequestPartyWithSearch(initailRequestParty);
            Thread.Sleep(1000);
            this.ClickPageTitle();
            Thread.Sleep(1000); 
            this.SetWeeklyRent(weeklyRent);
            this.SetInitialConrtibution(initialControbution);
            this.SetAmountPaid(amountPaid);
            this.SetLodgementTypeListValue(lodgementType);
            this.SetTenancyStartDate(tenancyStart);
            this.SetPaymentType(paymentType);
            this.SetAnticipatedEndDate(anticipatedEnd);

        }
        
        [ScenarioMethod]
        public void PopulateTenancyRequestWithNoInitialAndManagingParty(string requestType, string rentalPremises, string tenancyType, string managementType, string numBedrooms, string weeklyRent, string initialControbution, string amountPaid, string lodgementType, string tenancyStart, string anticipatedEnd, string paymentType)
        {
            this.SetRequestTypeListValue(requestType);
            this.SetRentalPremisesValue(rentalPremises);
            this.SetResidentialTenancyTypeList(tenancyType);
            this.SetResidentialManagementTypeList(managementType);
            this.SetNumberOfBedrooms(numBedrooms);
            this.SetWeeklyRent(weeklyRent);
            this.SetInitialConrtibution(initialControbution);
            this.SetAmountPaid(amountPaid);
            this.SetLodgementTypeListValue(lodgementType);
            this.SetTenancyStartDate(tenancyStart);
            this.SetPaymentType(paymentType);
            this.SetAnticipatedEndDate(anticipatedEnd);
        }
       
        [ScenarioMethod]
        public void PopulateTenancyRequestWithoutPaymentType(string requestType, string rentalPremises, string managingParty, string tenancyType, string managementType, string numBedrooms, string initailRequestParty, string weeklyRent, string initialControbution, string amountPaid, string lodgementType, string tenancyStart, string anticipatedEnd)
        {

            this.SetRequestTypeListValue(requestType);
            this.SetRentalPremisesValue(rentalPremises);
            this.SetManagingPartyListValue(managingParty);
            this.SetResidentialTenancyTypeList(tenancyType);
            this.SetResidentialManagementTypeList(managementType);
            this.SetNumberOfBedrooms(numBedrooms);
            this.SetInitialRequestPartyWithSearch(initailRequestParty);
            this.SetWeeklyRent(weeklyRent);
            this.SetInitialConrtibution(initialControbution);
            this.SetAmountPaid(amountPaid);
            this.SetLodgementTypeListValue(lodgementType);
            this.SetTenancyStartDate(tenancyStart);
            this.SetAnticipatedEndDate(anticipatedEnd);

        }

        [ScenarioMethod]
        public void PopulateMandatoryFieldValues(string requestType, string rentalPremises, string managingParty, string tenancyType, string managementType, string numBedrooms, string initailRequestParty, string weeklyRent, string initialControbution, string amountPaid, string lodgementType, string tenancyStart, string anticipatedEnd, string paymentType)
        {
            this.SetRequestTypeListValue(requestType);
            this.SetRentalPremisesValue(rentalPremises);
            this.SetManagingPartyListValue(managingParty);
            this.SetResidentialTenancyTypeList(tenancyType);
            this.SetTenancyManagementTypeListValue(managementType);
            this.SetNumberOfBedrooms(numBedrooms);
            this.SetInitialRequestPartyWithSearch(initailRequestParty);
            this.SetWeeklyRent(weeklyRent);
            this.SetInitialConrtibution(initialControbution);
            this.SetAmountPaid(amountPaid);
            this.SetLodgementTypeListValue(lodgementType);
            this.SetTenancyStartDate(tenancyStart);
            this.SetPaymentType(paymentType);
            this.SetAnticipatedEndDate(anticipatedEnd);
            
        }

        [ScenarioMethod]
        public void PopulateTenancyRequestChangeOfManagingPartyWithoutRequestParty(string requestType, string rentalPremises, 
            string managingParty, string tenancy, string previousManagingParty, string prevEndDate, string newCommenceDate,
            string acceptPrevDate, string acceptNewDate, string prevMethod, string acceptNewMeth)
        {
            this.SetRequestTypeListValue(requestType);
            this.SetRentalPremisesValue(rentalPremises);
            this.SetManagingPartyListValue(managingParty);
            this.SetTenancyValue(tenancy);
            this.SetPrevManagingPartyListValue(previousManagingParty);
            this.SetPreviousManagingPartyEndDateValue(prevEndDate);
            this.SetManagingPartyCommenceDateValue(newCommenceDate);
            this.SetPrevManagingPartyAcceptDateValue(acceptPrevDate); 
            this.SetManagingPartyAcceptDateValue(acceptNewDate);
            this.SetPrevAcceptMethod(prevMethod);
            this.SetAcceptMethod(acceptNewMeth);
         
        }

        [ScenarioMethod]
        public void PopulateTenancyRequestWithoutRentalPremisesFromAddressDetailAssociatedView(string requestType, string managingParty, string tenancyType, string managementType, string numBedrooms, string initailRequestParty, string weeklyRent, string initialControbution, string amountPaid, string lodgementType, string tenancyStart, string anticipatedEnd, string paymentType)
        {
            this.SetRequestTypeListValue(requestType);
            this.SetManagingPartyListValue(managingParty);
            this.SetResidentialTenancyTypeList(tenancyType);
            this.SetTenancyManagementTypeListValue(managementType);
            this.SetNumberOfBedrooms(numBedrooms);
            this.SetInitialRequestPartyWithSearch(initailRequestParty);
            this.SetWeeklyRent(weeklyRent);
            this.SetInitialConrtibution(initialControbution);
            this.SetAmountPaid(amountPaid);
            this.SetLodgementTypeListValue(lodgementType);
            this.SetTenancyStartDate(tenancyStart);
            this.SetPaymentType(paymentType);
            this.SetAnticipatedEndDate(anticipatedEnd);

        }
        
        /*
        * RentalPremisses
        * ************************************************************************
        *
        */

        public void SetRentalPremisesValue(string premisis)
        {
            
            UICommon.SetSearchableListValue("rta_rental_premisesid", premisis, driver);

        }

        public void ClearRentalPremisesValue()
        {
            string elementId = "rta_rental_premisesid";
            UICommon.ClearListBoxValue(elementId, driver);
        }

        /*
        * DwellingType
        * ************************************************************************
        */
        [ActionMethod]
        public bool GetDwellingTypeText(string dwelling)
        {
            SetDwellingTypeListValue(dwelling);
            GetDwellingType();
            return true;

        }

        [ActionMethod]
        public string GetDwellingType()
        {
            return UICommon.GetElementProperty("#rta_dwelling_typeid>div>span","title", driver);
        }

      
       
       
        [ActionMethod]
        public void SetDwellingTypeListValue(string dwelling)
        {
            
            UICommon.SetSearchableListValue("rta_dwelling_typeid", dwelling, driver);
        }

        

        /*
        * RentalPremisesTextBox
        * ************************************************************************
        */
        [ActionMethod]
        public string GetRentalPremisesTextBox()
        {

            return UICommon.GetTextFromElement("#rta_rental_premisesid>div>span", driver);
        }

    
   
        /*
        * ResidentialtenancyType
        * ************************************************************************
        */

        
        [ActionMethod]
        public void SetRequestTypeListValue(string type)
        {
            //Switch to main frame when it is visible
            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);

            UICommon.SetSelectListValue("rta_type", type, driver);
  
        }


        /*
        * ResidentialtenancyType
        * ************************************************************************
        */
        [ActionMethod]
        public bool isRequestType(string type)
        {
            //Switch to main frame when it is visible
            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);
            
            return UICommon.DoesSelectListValueExist("rta_type", type, driver);
        }

       
        /*
        * TenancyTypeList
        * ************************************************************************
        */

        [ActionMethod]
        public bool GetTenancyTypeListValue(string listValue)
        {
            UICommon.SetSelectListValue("rta_tenancy_type", listValue, driver);
            return true;

        }

        [ActionMethod]
        public void SetTenancyTypeListValue(string listValue)
        {
            UICommon.SetSelectListValue("rta_tenancy_type", listValue, driver);
        }

        
        /*
        * InitialRequestPartyElement
        * ************************************************************************
        */

   
        [ActionMethod]
        public void ClickInitialRequestPartyListValue(string party)
        {
            
            UICommon.ClickSearchableListValue("rta_initial_request_partyid", party, driver);

        }

          

        /*
        * ManagingPartyElement
        * ************************************************************************
        */
        [ActionMethod]
        public void SetManagingPartyListValue(string party)
        {
            UICommon.SetSearchableListValue("rta_managing_partyid", party, driver);
        }

        [ActionMethod]
        public void SetPrevManagingPartyListValue(string party)
        {
            UICommon.SetSearchableListValue("rta_previous_managing_partyid", party, driver);
        }

        /*
       * Weekly Rent
       * ************************************************************************
       */
        [ActionMethod]
        public void SetWeeklyRent(string amount)
        {
              UICommon.SetTextBoxValue("rta_weekly_rent", amount, driver);
        }
        
        /*
        * InitialContributionText
        * ************************************************************************
        */
        [ActionMethod]
        public void SetInitialConrtibution(string amount)
        {
            UICommon.SetTextBoxValue("rta_initial_contribution", amount, driver);
        }

        /*
        * AmountPaidText
        * ************************************************************************
        */

        [ActionMethod]
        public void SetAmountPaid(string amount)
        {
            UICommon.SetTextBoxValue("rta_amount_bond_paid", amount, driver);
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
      * Distribute
      * ************************************************************************
      */

        [ActionMethod]
        public string ClickDistributeButton()
        {

            this.driver.SwitchTo().DefaultContent();
            
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[alt='Distribute Contributions Equally']")));
            Actions action = new Actions(driver);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).Release().Build().Perform();
            Thread.Sleep(3000);

            string alertMessage = this.GetAlertMessage();
            return alertMessage;
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
            driver.SwitchTo().DefaultContent();
            UICommon.ClickNewButton(driver);
          
        }


        /*
        * RentalPremiseAddressERROR
        * ************************************************************************
        */

        [ActionMethod]
        public String GetRentalPremiseAddressErrorText()
        {
            
            return UICommon.GetTextFromElement("#rta_rental_premisesid_err", driver);

        }

        /*
       * TenancyTypeERROR
       * ************************************************************************
       */

        [ActionMethod]
        public String GetTenancyTypeErrorText()
        {
            return UICommon.GetTextFromElement("#rta_tenancy_type_err", driver);

        }

        /*
        * RTATenancyManagementType
        * ************************************************************************
        */
        
        
          
        [ActionMethod]
        public IList<IWebElement> GetTenancyManagementTypeListValues()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_tenancy_management_type_i")));

            SelectElement selectList = new SelectElement(elem);
            IList<IWebElement> options = selectList.Options;
            
            return options;

        }
         
        [ActionMethod]
        public void SetTenancyManagementTypeListValue(string RTATenancyManagementType)
        {
            UICommon.SetSelectListValue("rta_tenancy_management_type", RTATenancyManagementType, driver);
        }

        [ActionMethod]
        public bool GetTenancyManagementTypeListValue(string RTATenancyManagementType)
        {
            SetTenancyManagementTypeListValue( RTATenancyManagementType);
            return true;
        }



        /*
        * NumberOfBedrooms
        * ************************************************************************
        */
        
       
        [ActionMethod]
        public string GetNumberOfBedrooms()
        {
            
            RefreshPageFrame.RefreshPage(driver, frameId);
            return UICommon.GetTextFromElement("#rta_number_of_bedrooms>div>span", driver);
            
        }

        [ActionMethod]
        public string GetNumberOfBedroomsProperty(string property)
        {

            RefreshPageFrame.RefreshPage(driver, frameId);
            return UICommon.GetElementProperty("#rta_number_of_bedrooms", property, driver);

        }

        [ActionMethod]
        public void SetNumberOfBedrooms(string numBedrooms)
        {
            UICommon.SetTextBoxValue("rta_number_of_bedrooms", numBedrooms, driver);
            
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
            try
            {
                return UICommon.GetTextFromElement("#crmNotifications", driver);
            }catch(Exception)
            {
                return "";
            }
        }

       
        [ActionMethod]
        public void SetTenancyStartDate(string dateValue)
        {
            
            UICommon.SetDateValue("rta_tenancy_start_date", dateValue, driver);

        }


        /*
        * Anticipated End Date
        * ************************************************************************
        */
        [ActionMethod]
        public void SetAnticipatedEndDate(string dateValue)
        {

            UICommon.SetDateValue("rta_anticipated_tenancy_end_date", dateValue, driver);

        }

        /*
        * Date Received at RTA
        * ************************************************************************
        */
        [ActionMethod]
        public string GetPropertyDataControlModeRTADateReceivedAtRTA()
        {
            return UICommon.GetElementProperty("#rta_date_bond_received", "data-controlmode", driver);
   
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
            return UICommon.GetTextFromElement("#rta_funded_status>div>span", driver);
        }

        [ActionMethod]
        public string GetMaximumAllowedAmount()
        {
            return UICommon.GetTextFromElement("#rta_maximum_allowed_bond>div>span", driver);
        }

        [ActionMethod]
        public string GetMaximumAllowedAmountRefreshed(int refreshTimeOut)
        {
            
            for (int i = 0; i <= refreshTimeOut; i++)
            {
                string maxBondAmount = UICommon.GetTextFromElement("#rta_maximum_allowed_bond>div>span", driver);
                string statusReason = UICommon.GetTextFromElement("#statuscode>div>span", driver);

                if ((maxBondAmount == "--") & (statusReason == "Ready for validation"))
                {
                    Thread.Sleep(1000);
                    this.ClickSaveButton();
                }
                else
                {
                    if (statusReason == "Validation failed")
                    {
                        UICommon.SetSelectListValue("statuscode", "Ready for validation", driver);
                        this.ClickSaveButton();
                    }
                    else
                    {
                        return maxBondAmount;
                    }
                }
            } return "Validation failed";

        }

        [ActionMethod]
        public void SetAmountPaidWithLodgement(string amountPaid)
        {
            UICommon.SetTextBoxValue("rta_amount_bond_paid", amountPaid, driver);
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
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Subgrid_paymentref_openAssociatedGridViewImageButtonImage"))).Click();
            
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
            RefreshPageFrame.RefreshPage(driver, frameId);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("SubGrid_TenancyRequestParties_addImageButtonImage")));
            
            Actions action = new Actions(driver);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(2000);
            action.MoveToElement(elem).Release().Build().Perform();
            Thread.Sleep(3000);
           

        }

        /*
       * Return to Tenancy Request Home Page
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
       * search criteria
       * ************************************************************************
       */

        [ActionMethod]
        public void SetTenancyRequestPaymentSearchText(string searchValue)
        {
            driver.SwitchTo().Frame(FRAMEpaymentReference);
            UICommon.SetSearchText("crmGrid_findCriteria", searchValue, driver);

        }


        /*
        * search table 
        * ************************************************************************
        */

        [ActionMethod]
        public IWebElement GetPaymentSearchResultTable()
        {
            RefreshPageFrame.RefreshPage(driver, frameId);
            driver.SwitchTo().Frame(FRAMEpaymentReference);
            return UICommon.GetSearchResultTable(driver);
        }

       

        [ActionMethod]
        public IWebElement GetPaymentSummaryResultTable()
        {

            RefreshPageFrame.RefreshPage(driver, frameId);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement parent = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Subgrid_paymentref_divDataArea")));

            IWebElement webElementBody = parent.FindElement(By.Id("gridBodyTable"));
            return webElementBody;
        }


        [ActionMethod]
        public IWebElement GetAuditHistoryTable()
        {
            RefreshPageFrame.RefreshPage(driver, frameId, "areaAuditFrame");
            return UICommon.GetSearchResultTable(driver);
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
                else if (elem.Text == "Validation failed")
                {
                    UICommon.SetSelectListValue("statuscode", "Ready for validation", driver);
                    this.ClickSaveButton();
                }else
                {
                    break;
                }
            }
            WaitForPageToLoad.WaitToLoad(driver);
            return elem.Text;
        }
        
       
        [ActionMethod]
        public bool GetValidationStatusReasonNotExits(string reason)
        {
            RefreshPageFrame.RefreshPage(driver, frameId);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#statuscode>div")));

            Actions action = new Actions(driver);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(2000);
            action.MoveToElement(elem).Release().Build().Perform();
            WaitForPageToLoad.WaitToLoad(driver);


            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("statuscode_i")));
            //new SelectElement(elem).SelectByText(reason);
            IList<IWebElement> list = driver.FindElements(By.XPath("//select[@id='statuscode_i']/option[text()='" + reason + "']"));
            if (list.Count == 0)
            { return true;
            }
            else { return false;
            }

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

        [ActionMethod]
        public void SetPrevAcceptMethod(string type)
        {
            UICommon.SetSelectListValue("rta_acceptance_method_previous_managing_party", type, driver);
        }
        
        [ActionMethod]
        public void SetAcceptMethod(string type)
        {
            UICommon.SetSelectListValue("rta_acceptance_method_managing_party1", type, driver);
        }

        

        /*
       * Amount Matched
       * ************************************************************************
       */
        [ActionMethod]
        public string GetAmountMatched()
        {
            return UICommon.GetTextFromElement("#rta_amount_matched>div>span", driver);
        }


        public IWebDriver SwitchNewBrowser(IWebDriver driver, string BaseWindow)
        {
            return UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow,"Tenancy Request Party");
        }


        /*
        * Titlefooter_statuscontrol
        * ************************************************************************
        */

        [ActionMethod]
        public string GetRecordStatus()
        {
            return UICommon.GetTextFromElement("#titlefooter_statuscontrol", driver);
        }

        [ActionMethod]
        public string GetStatusReason()
        {
            
            return UICommon.GetTextFromElement("#statuscode>div>span", driver);
        }

        /*
        * Section Headers / Tables
        * ************************************************************************
        */

        [ActionMethod]
        public void ClickQueueReasons()
        {
            UICommon.ClickTabHeader("tab_RequestQueueReasons", driver);
        }

        [ActionMethod]
        public IWebElement GetQueueReasonTable()
        {
     
            RefreshPageFrame.RefreshPage(driver, frameId);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement parent = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Subgrid_RequestQueueReasons_divDataArea")));
            IWebElement webElementBody = parent.FindElement(By.Id("gridBodyTable"));
            return webElementBody;
        }


        [ActionMethod]
        public IWebElement GetRequestPartyTable()
        {
            RefreshPageFrame.RefreshPage(driver, frameId);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement parent = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("SubGrid_TenancyRequestParties_divDataArea")));

            IWebElement webElementBody = parent.FindElement(By.Id("gridBodyTable"));
            return webElementBody;
        }

        /*
        * Previous ManagingPartyElement
        * ************************************************************************
        */
        [ActionMethod]
        public void SetPreviousManagingPartyListValue(string party)
        {
            UICommon.SetSearchableListValue("rta_previous_managing_partyid", party, driver);
        }

        [ActionMethod]
        public void SetPreviousManagingPartyEndDateValue(string dateValue)
        {
            UICommon.SetDateValue("rta_date_previous_management_ended", dateValue, driver);
        }

        [ActionMethod]
        public void SetManagingPartyCommenceDateValue(string dateValue)
        {
            UICommon.SetDateValue("rta_date_new_management_commenced", dateValue, driver);
        }

        [ActionMethod]
        public void SetManagingPartyAcceptDateValue(string dateValue)
        {
            UICommon.SetDateValue("rta_date_accepted_managing_party1", dateValue, driver);
        }

        [ActionMethod]
        public void SetPrevManagingPartyAcceptDateValue(string dateValue)
        {
            UICommon.SetDateValue("rta_date_accepted_previous_managing_party", dateValue, driver);
        }
        

        /*
        * New ManagingPartyElement
        * ************************************************************************
        */
        [ActionMethod]
        public void SetNewManagingPartyListValue(string party)
        {
            UICommon.SetSearchableListValue("rta_managing_partyid1", party, driver);
        }

        [ActionMethod]
        public void SetNewManagingPartyEndDateValue(string dateValue)
        {
            UICommon.SetDateValue("rta_date_new_management_commenced", dateValue, driver);
        }

        [ActionMethod]
        public void SetSubsidy(string subsidy)
        {
            UICommon.SetSelectListValue("rta_rent_subsidy", subsidy, driver);
        }

        [ActionMethod]
        public bool ClickDeleteButtonIfDisplayed(IWebElement element)
        {
            IWebElement deleteElement = null;
            Actions actions = new Actions(driver);
            actions.MoveToElement(element).Build().Perform();
            Thread.Sleep(3000);
            
           WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
           try
           {
               wait.Until((d) =>
               {
                   deleteElement = element.FindElement(By.XPath("../td[contains(@id,'gridBodyTable_gridDelBtn')]//a[@title='Delete']/img"));
                   return deleteElement.Displayed;
               });

           }
           catch (Exception)
           {
               return false;
           }   
            if(deleteElement!=null)
            {
                actions.MoveToElement(deleteElement).ClickAndHold().Build().Perform();
                Thread.Sleep(1000);
                actions.MoveToElement(deleteElement).Release().Build().Perform();
                return true;
            }else
            {
                return false;
            }
           
        }

        [ActionMethod]
        public void SetInitialRequestPartyWithSearch(string initialRequestParty)
        {
            UICommon.SetSearchableListValue("rta_initial_request_partyid",initialRequestParty,driver);
        }
        
        [ActionMethod]
        public void ClickDeactivateButton()
        {
            driver.SwitchTo().DefaultContent();
            UICommon.ClickDeactivateButton(driver);
        }

        [ScenarioMethod]
        public void CreateNewAddress(string roadno,string roadname,string locality,string roomtype = "",string roomno="",string complexunitno="")
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
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_rental_premisesid")));
            Actions actions = new Actions(driver);
            actions.MoveToElement(elem).Build().Perform();
            Thread.Sleep(500);
            
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#rta_rental_premisesid>div>div>img"))).Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_rental_premisesid_i"))).Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//div[@id='Dialog_rta_rental_premisesid_i_IMenu']//a[contains(@title,'Create a new Address Detail.')]"))).Click();
            Thread.Sleep(5000);

            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "Address Detail");
        }

        public void CheckForErrors()
        {
            try
            {
                driver.SwitchTo().Frame(dialogFRAME);
                UICommon.ClickElementWithId("butBegin", driver);
                driver.SwitchTo().DefaultContent();
                driver.SwitchTo().Frame(frameId);
            }
            catch
            { }

        }
    }
}
