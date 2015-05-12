using ActionWordsLib.Attributes;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using RTA.Automation.CRM.UI;
using RTA.Automation.CRM.Utils;
using RTA.Automation.CRM.Pages.Clients;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using AutoItX3Lib;

namespace RTA.Automation.CRM.Pages
{
    [ActionClass]
    public class ClientPage : IFramePage
    {
        private static string frameId = "contentIFrame0";
        private static string dialogFRAME = "InlineDialog_Iframe";
        private static string connectionsFRAME = "areaConnectionsFrame";
        private static string contentFrameId = "areaActivitiesFrame";
        private static string auditFRAME = "areaAuditFrame";

        public static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;
        private static string pageTitle = "Client:";

        public ClientPage(IWebDriver driver)
            : base(driver, ClientPage.frameId)
        {
            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);


        }

        public ClientPage(IWebDriver driver,string pagetitle)
            : base(driver, ClientPage.frameId)
        {
            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until((d) => { return d.Title.Contains(pagetitle); });

            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);
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
        public IWebElement GetCurrentAlertsTable()
        {

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            //Wait for table to load
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#Subgrid_alerts_gridBodyContainer>div")));
            IWebElement webElementBody = elem.FindElement(By.Id("Subgrid_alerts_divDataBody"));

            return webElementBody;

        }


        [ActionMethod]
        public void ClickAddAlertElement()
        {

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Subgrid_alerts_addImageButtonImage")));
            elem.Click();

        }



        [ScenarioMethod]
        public void PopulateNewClient(string familyName)
        {
            //this.ClickFamilyNameTextBoxArea();
            SetUnknownClientListValues("No");
            SetFamilyName(familyName);
        }


        /*
        * NewIMG
        * ************************************************************************
        */

        [ActionMethod]
        public void ClickNewClientButton()
        {
            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickNewButton(driver);


        }

        /*
        * FamilyNameTextBox
        * ************************************************************************
        */

        //[ActionMethod]
        //public void ClickFamilyNameTextBoxArea()
        //{

        //    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
        //    wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#lastname>div>span"))).Click();

        //    RefreshPageFrame.RefreshPage(driver, frameId);


        //}

        [ActionMethod]
        public void SetFamilyName(string textValue)
        {
            UICommon.SetTextBoxValue("lastname", textValue, driver);
        }

        [ActionMethod]
        public void SetGivenName(string textValue)
        {
            UICommon.SetTextBoxValue("firstname", textValue, driver);
        }

        [ActionMethod]
        public void SetMiddleName(string textValue)
        {
            UICommon.SetTextBoxValue("middlename", textValue, driver);
        }

        [ActionMethod]
        public void SetABNACNValue(string textValue)
        {
            UICommon.SetTextBoxValue("rta_abn_acn", textValue, driver);

        }

        [ActionMethod]
        public void SetARBNValue(string textValue)
        {
            UICommon.SetTextBoxValue("rta_arbn", textValue, driver);

        }

        [ActionMethod]
        public Boolean CheckFamilyNameErrorPresent()
        {
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(20));
            List<IWebElement> warning = driver.FindElements(By.Id("lastname_warn")).ToList();
            if (warning.Count <= 1)
            {
                return true;
            }
            else
            {
                return true;
            }
        }


        [ActionMethod]
        public void SetEmailCorrespondenceValue(string textValue)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#rta_agreed_email_correspondence>div"))).Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_agreed_email_correspondence_i")));
            new SelectElement(driver.FindElement(By.Id("rta_agreed_email_correspondence_i"))).SelectByText(textValue);

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
        * ClientID
        * ************************************************************************
        */

        [ActionMethod]
        public string GetClientID()
        {
            
            return UICommon.GetTextFromElement("#header_rta_rta_client_id>div>span", driver);


        }

        /*
      * Preferences - Email Correspondence
      * ************************************************************************
      */

        [ActionMethod]
        public string GetEmailCorrespondenceValue()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            string textValue = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#rta_agreed_email_correspondence>div>span"))).Text;
            return textValue;
        }

        /*
       * Unknown Client
       * ************************************************************************
       */

        [ActionMethod]
        public string GetUnknownClientListValues()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_unknownclient"))).Click();
            string textValue = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_unknownclient_i"))).Text;
            return textValue;

        }

        [ActionMethod]
        public void SetUnknownClientListValues(string UnknownClient)
        {
            UICommon.SetSelectListValue("rta_unknownclient", UnknownClient, driver);
        }

        /*
       * Suffix
       * ************************************************************************
       */

        [ActionMethod]
        public void ClickSuffixSearchButton()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_name_suffixid_i"))).Click();
            //Wait for list to be displayed
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Dialog_rta_name_suffixid_i_IMenu"))).GetAttribute("id");

            IWebElement elem = driver.FindElement(By.Id("rta_name_suffixid_i"));
            elem.SendKeys(Keys.Backspace);

            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_name_suffixid_ledit"))).GetAttribute("id");

        }

        [ActionMethod]
        public void ClickSuffixElement(string suffix)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//li[a[contains(@title,'" + suffix + "')] and parent::*[@id='rta_name_suffixid_i_IMenu']]"))).Click();

        }


        [ActionMethod]
        public void SetSuffixElement(string suffix)
        {

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_name_suffixid_ledit")));
            elem.SendKeys(suffix);
            elem.SendKeys(Keys.Enter);
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_name_suffixid_i_IMenu"))).GetAttribute("id");


        }
        [ActionMethod]
        public bool GetSuffixText(string suffix)
        {
            ClickSuffixText();
            ClickSuffixSearchButton();
            SetSuffixElement(suffix);
            ClickSuffixElement(suffix);
            return true;

        }

        [ActionMethod]
        public void ClickSuffixText()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#rta_name_suffixid>div"))).Click();
            //Wait for search button to be available
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_name_suffixid_i"))).GetAttribute("id");

        }

        /*
         * Title
         * ************************************************************************
         */

        [ActionMethod]
        public void ClickTitleSearchButton()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_name_titleid_i"))).Click();
            //Wait for list to be displayed
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Dialog_rta_name_titleid_i_IMenu"))).GetAttribute("id");

            IWebElement elem = driver.FindElement(By.Id("rta_name_titleid_i"));
            elem.SendKeys(Keys.Backspace);

            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_name_titleid_ledit"))).GetAttribute("id");

        }
        [ActionMethod]
        public void ClickTitleElement(string title)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//li[a[contains(@title,'" + title + "')] and parent::*[@id='rta_name_titleid_i_IMenu']]"))).Click();

        }
        [ActionMethod]
        public void SetTitleElement(string title)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_name_titleid_ledit")));
            elem.SendKeys(title);
            elem.SendKeys(Keys.Enter);
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_name_titleid_i_IMenu"))).GetAttribute("id");


        }
        [ActionMethod]
        public bool GetTitleText(string title)
        {
            ClickTitleText();
            ClickTitleSearchButton();
            SetTitleElement(title);
            ClickTitleElement(title);
            return true;

        }
        [ActionMethod]
        public void ClickTitleText()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_name_titleid"))).Click();

        }

        [ActionMethod]
        public void ClickPageTitle()
        {
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(frameId);
            UICommon.ClickPageTitle(driver);

        }



        /*
        * Add New Client IMG
        * ************************************************************************
        */

        [ActionMethod]
        public void ClickAddNewClientPhoneImage()
        {
            WaitForPageToLoad.WaitToLoad(driver);
            driver.SwitchTo().Frame("rta_contact_client_rta_client_ph_no_clientidFrame");
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[alt='Add New Client Phone Number']")));
            Actions action = new Actions(driver);
            Thread.Sleep(2000);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(2000);
            action.MoveToElement(elem).Release().Build().Perform();

        }

        /*
       * Add New Client Names
       * ************************************************************************
       */

        [ActionMethod]
        public void ClickAddNewClientName()
        {
            WaitForPageToLoad.WaitToLoad(driver);
            driver.SwitchTo().Frame("rta_contact_rta_client_name_currently_known_asidFrame");
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[alt='Add New Client Name']")));

            Actions action = new Actions(driver);
            Thread.Sleep(2000);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            //action.MoveToElement(elem).Click().Build().Perform();
            Thread.Sleep(2000);
            action.MoveToElement(elem).Release().Build().Perform();
            Thread.Sleep(15000);

        }

        /*
        * Add New Client Id Artefact
        * ************************************************************************
        */

        [ActionMethod]
        public void ClickAddNewClientIdArtefact()
        {

            WaitForPageToLoad.WaitToLoad(driver);
            driver.SwitchTo().Frame("rta_contact_rta_client_id_artefact_clientidFrame");
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[alt='Add New Client Identification Artefact']")));

            Actions action = new Actions(driver);
            Thread.Sleep(2000);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            //action.MoveToElement(elem).Click().Build().Perform();
            Thread.Sleep(2000);
            action.MoveToElement(elem).Release().Build().Perform();
            Thread.Sleep(15000);


        }

        /*
         * Add New Address 
         * ************************************************************************
         */

        [ActionMethod]
        public void ClickAddNewClientAddressImage()
        {
            WaitForPageToLoad.WaitToLoad(driver);
            driver.SwitchTo().Frame("rta_contact_rta_client_address_clientidFrame");
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[alt='Add New Client Address']")));
            Actions action = new Actions(driver);
            Thread.Sleep(2000);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(2000);
            action.MoveToElement(elem).Release().Build().Perform();

        }

        [ActionMethod]
        public void ClickAddNewClientAddressImageIRSIT()
        {
            WaitForPageToLoad.WaitToLoad(driver);
            driver.SwitchTo().Frame("area_rta_contact_rta_client_address_clientidFrame");
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[alt='Add New Client Address']")));
            Actions action = new Actions(driver);
            Thread.Sleep(2000);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(2000);
            action.MoveToElement(elem).Release().Build().Perform();

        }

        /*
        * search criteria
        * ************************************************************************
        */

        [ActionMethod]
        public void SetClientSearchText(string searchValue)
        {

            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame("contentIFrame1");
            driver.SwitchTo().Frame("rta_contact_client_rta_client_ph_no_clientidFrame");

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            Thread.Sleep(3000);
            IWebElement element = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#crmGrid_rta_contact_client_rta_client_ph_no_clientid_SavedNewQuerySelector>span")));
            string title = element.GetAttribute("title");


            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("crmGrid_rta_contact_client_rta_client_ph_no_clientid_findCriteria")));
            element = this.driver.FindElement(By.Id("crmGrid_rta_contact_client_rta_client_ph_no_clientid_findCriteria"));
            element.Clear();
            element.SendKeys(searchValue.ToString());
            element.SendKeys(Keys.Enter);
            WaitForPageToLoad.WaitToLoad(driver);
            driver.SwitchTo().DefaultContent();
        }

        /*
        * search table 
        * ************************************************************************
        */

        [ActionMethod]
        public IWebElement GetSearchResultTable()
        {
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame("contentIFrame1");
            driver.SwitchTo().Frame("rta_contact_client_rta_client_ph_no_clientidFrame");
            WaitForPageToLoad.WaitToLoad(driver);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("gridBodyTable")));
            IWebElement webElementBody = driver.FindElement(By.Id("gridBodyTable"));
            return webElementBody;
        }

        [ActionMethod]
        public void ClickPreferencesTab()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//div/a/h2[contains(text(),'Preferences')]")));
            string text = elem.GetAttribute("text");
            Actions action = new Actions(driver);
            Thread.Sleep(2000);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(2000);
            action.MoveToElement(elem).Release().Build().Perform();

        }

        public void SetConnectList(string connectType)
        {

            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(frameId);
            driver.SwitchTo().Frame(connectionsFRAME);

            UICommon.SetConnectList(connectType, driver);
        }

        internal IWebDriver SwitchNewBrowser(IWebDriver driver, string BaseWindow, string newWindow)
        {
            return UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, newWindow);
        }

        /*
        * Email
        * ************************************************************************
        */
        private IWebElement GetEmail1Field()
        {
            return this.driver.FindElement(By.Id("emailaddress1"));
        }
        public void SetEmail1ID(String Email1)
        {
            UICommon.SetTextBoxValue("emailaddress1", Email1, driver);
        }
        public String GetEmail1ID()
        {
            return UICommon.GetTextFromElement("#emailaddress1",driver);
        }

        public void SetEmail2ID(String Email2)
        {
            UICommon.SetTextBoxValue("emailaddress2", Email2, driver);
        }

        public void ClearEmail1Id()
        {
            UICommon.ClearTextBoxValue("emailaddress1", driver);
        }

        public  void ClearEmail2Id()
        {
            UICommon.ClearTextBoxValue("emailaddress2", driver);
        }
        public String GetEmail2ID()
        {
            return UICommon.GetTextFromElement("#emailaddress2", driver);
        }

        public String GetBannerEmailValue()
        {
            return UICommon.GetTextFromElement("#header_emailaddress1>div>span>a", driver);
        }

        public String GetAddressValue(string cssSelector)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            return UICommon.GetTextFromElement("#" + cssSelector + ">div>span", driver);
        }

        public void setEmail1Address(string emailAddress)
        {
            UICommon.SetTextBoxValue("emailaddress1", emailAddress, driver);

        }

        /*
        * Client Address
        * ************************************************************************
        */

        public void ClickCreateNewClientAddressButton(string cssSelector)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elemList = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#" + cssSelector + ">div")));
            elemList.Click();

            elemList = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(cssSelector + "_i")));
            elemList.Click();

            elemList = wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("ms-crm-InlineLookup-FooterSection-AddAnchor")));
            elemList.Click();
        }

        /*
        * Date of Birth
        * ************************************************************************
        */
        public void SetDateOfBirthValue(string dateValue)
        {
            this.ClickPageTitle();
            UICommon.SetDateValue("birthdate", dateValue, driver);
        }

        public string GetBirthdayErrorText()
        {
            return UICommon.GetTextFromElement("#birthdate_err", driver);
        }

        public string GetEmail1ErrorText()
        {
            return UICommon.GetTextFromElement("#emailaddress1_err", driver);
        }

        [ActionMethod]
        public Boolean CheckDateOfBirthErrorPresent()
        {
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(20));
            List<IWebElement> warning = driver.FindElements(By.Id("birthdate_err")).ToList();
            if (warning.Count >= 1)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /*
        * Personal preferred Client number
        * ************************************************************************
        */

        public void ClickAddNewPersonalPreferredClientNumber()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elemList = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#rta_personal_preferredcontactid>div")));
            elemList.Click();

            elemList = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_personal_preferredcontactid_i")));
            elemList.Click();

            elemList = wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("ms-crm-InlineLookup-FooterSection-AddAnchor")));
            elemList.Click();
        }

        public string GetPersonalPreferredMobileNumber()
        {
            return UICommon.GetTextFromElement("#rta_personal_preferredcontactid>div>span", driver);
        }

        /*
        * Client Type
        * ************************************************************************
        */

        public void SetClientType(string clientType)
        {
            UICommon.SetSelectListValue("rta_client_type", clientType, driver);
        }

        /*
         * Organization Name Text Box
         * ************************************************************************
         */

        [ActionMethod]
        public void ClickOrganizationNameTextBoxArea()
        {

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#rta_organisation_name>div>span"))).Click();

            RefreshPageFrame.RefreshPage(driver, frameId);


        }

        [ActionMethod]
        public void SetOrganizationName(string textValue)
        {
            this.ClickOrganizationNameTextBoxArea();
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_organisation_name_i"))).SendKeys(textValue);

        }

        /*
         * Warning message on top of page e.g. Physical address is blank etc.
         * ************************************************************************
         */

        public string GetWarningMessage(string cssSelector)
        {
            return UICommon.GetWarningMessage(cssSelector, driver);
        }

        public bool VerifyWarningMessagePresent(string cssSelector)
        {
            return UICommon.VerifyWarningMessagePresent(cssSelector, driver);
        }

        /*
         * Activities Table
         * ************************************************************************
         */


        [ActionMethod]
        public void ClickActivitiesAddButton()
        {
            UICommon.ClickAddButton(driver, "Activities_addImageButtonImage");
        }

        [ActionMethod]
        public void ClickAddActivity(string ActivitycssSelectorId)
        {
            UICommon.ClickAddActivity(ActivitycssSelectorId, driver);
        }


        [ActionMethod]
        public IWebElement GetActivitiesTable()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement webElementBody = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("table[ologicalname='activitypointer']")));
            return webElementBody;
        }

        [ActionMethod]
        public IWebElement GetActivitiesHeaderTable()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement webElementBody = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#Activities_gridBar")));
            return webElementBody;
        }

        public void ClickSeeRecordsAssociatedWithThisViewButton(string association)
        {
            UICommon.ClickSeeRecordsAssociatedWithThisViewButton(association, driver);
        }

        /*
         *     Activities Associated View Table
         * ************************************************************************
         */
        [ActionMethod]
        public IWebElement GetActivitiesAssociatedViewTable(IWebDriver d)
        {
           
            return UICommon.GetSearchResultTable(d);
            
        }

        [ActionMethod]
        public void SetActivitiesSearchText(string SearchText)
        {
            UICommon.SetSearchText("crmGrid_findCriteria", SearchText, driver);
        }

        [ActionMethod]
        public void SetPageFilterList(string value, IWebDriver d)
        {

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement parent = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#crmGrid_Contact_ActivityPointers_SavedNewQuerySelector>span")));
            Actions action = new Actions(driver);
            action.MoveToElement(parent).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(parent).Release().Build().Perform();
            Thread.Sleep(1000);
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Dialog_0")));
            elem.FindElement(By.XPath("//li[a[contains(@title,'" + value + "')]]")).Click();
           

        }

        
        
        [ActionMethod]
        public void SetFilterOnList(string value, IWebDriver d)
        {

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("crmGrid_Contact_ActivityPointers_datefilter")));
            elem.Click();

            IReadOnlyCollection<IWebElement> OptionList = elem.FindElements(By.CssSelector("Option"));
            foreach(IWebElement option in OptionList)
            {
                if (option.Text.Contains(value))
                {
                    option.Click();
                }
            }
            
        }

        public void SwitchToFrame()
        {
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(frameId);
            driver.SwitchTo().Frame(contentFrameId);        
        }


        [ActionMethod]
        public string GetABNACNErrorText()
        {
            return UICommon.GetTextFromElement("#rta_abn_acn_err", driver);
        }
        [ActionMethod]
        public string GetARBNErrorText()
        {
            return UICommon.GetTextFromElement("#rta_arbn", driver);
        }

        [ActionMethod]
        public bool GetABNACNErrorTextNotExist()
        {
            return UICommon.GetElementNotExistTrue("#rta_abn_acn_err", driver);
        }
        [ActionMethod]
        public bool GetARBNErrorTextNotExist()
        {
            return UICommon.GetElementNotExistTrue("#rta_arbn", driver);
        }
        [ActionMethod]
        public bool GetEmailAddress1ErrorTextNotExist()
        {
            return UICommon.GetElementNotExistTrue("#emailaddress1_err", driver);
        }


        [ActionMethod]
        public void SetPostalAddress(string p)
        {
            UICommon.SetSearchableListValue("rta_postaladdressid", p, driver);
        }

        [ActionMethod]
        public string GetPostalAddress()
        {
            return UICommon.GetTextFromElement("#rta_postaladdressid>div>span", driver);
        }

        [ActionMethod]
        public void ClickNextPageAlertTable()
        {
            WebDriverWait  wait = new WebDriverWait(driver,TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#Subgrid_alerts_StatusBar img[alt='Load Next Page']"))).Click();
        }

        [ActionMethod]
        public void SetParentOrganization(string parentorg)
        {
            UICommon.SetSearchableListValue("parentcustomerid", parentorg, driver);
        }
        
        [ActionMethod]
        public void ClickStartDialogButton()
        {
            this.driver.SwitchTo().DefaultContent();
            UICommon.ClickStartDialogButton(driver);
            driver.SwitchTo().Frame(frameId);
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


        /*
         *     Verify search options present for Mobile, FAX, Personal Preferred number etc.
         * ************************************************************************
         */

        public void OpenMobileNumbeDropDown()
        {
            this.ClickPageTitle();
            this.ClickPageTitle();
            UICommon.OpenSelectOptionDropDown("rta_personal_mobileid", driver);
        }

        [ActionMethod]
        public bool VerifyMobileNumberOptionPresent(string SelectOption)
        {
            return UICommon.VerifySelectOptionPresent("rta_personal_mobileid", SelectOption, driver);
        }

        [ActionMethod]
        public void SetMobileNumber(string Number)
        {
            this.ClickPageTitle();
            UICommon.SetSearchableListValue("rta_personal_mobileid", Number, driver);
        }

        [ActionMethod]
        public void ClearMobileNumber()
        {
            UICommon.ClearListBoxValue("rta_personal_mobileid", driver);
        }

        [ActionMethod]
        public void OpenPersonalPreferredNumbeDropDown()
        {
            this.ClickPageTitle();
            this.ClickPageTitle();
            UICommon.OpenSelectOptionDropDown("rta_personal_preferredcontactid", driver);
        }

        [ActionMethod]
        public bool VerifyPersonalPreferredNumberOptionPresent(string SelectOption)
        {

            return UICommon.VerifySelectOptionPresent("rta_personal_preferredcontactid", SelectOption, driver);
        }

        [ActionMethod]
        public void SetPersonalPreferredMobileNumber(string Number)
        {
            this.ClickPageTitle();
            UICommon.SetSearchableListValue("rta_personal_preferredcontactid", Number, driver);
        }

        [ActionMethod]
        public void ClearPersonalPreferredMobileNumber()
        {
            UICommon.ClearListBoxValue("rta_personal_preferredcontactid", driver);
        }

        [ActionMethod]
        public void OpenWorkPreferredNumbeDropDown()
        {
            this.ClickPageTitle();
            this.ClickPageTitle();
            UICommon.OpenSelectOptionDropDown("rta_work_preferredid", driver);
        }

        [ActionMethod]
        public bool VerifyWorkPreferredOptionPresent(string SelectOption)
        {
            return UICommon.VerifySelectOptionPresent("rta_work_preferredid", SelectOption, driver);
        }

        [ActionMethod]
        public void SetWorkPreferredNumber(string Number)
        {
            this.ClickPageTitle();
            UICommon.SetSearchableListValue("rta_work_preferredid", Number, driver);
        }

        [ActionMethod]
        public void ClearWorkPreferredNumber()
        {
            UICommon.ClearListBoxValue("rta_work_preferredid", driver);
        }

        [ActionMethod]
        public void OpenHomeMainPhoneNumbeDropDown()
        {
            this.ClickPageTitle();
            this.ClickPageTitle();
            UICommon.OpenSelectOptionDropDown("rta_main_phoneid", driver);
        }

        [ActionMethod]
        public bool VerifyHomeMainPhoneOptionPresent(string SelectOption)
        {
            return UICommon.VerifySelectOptionPresent("rta_main_phoneid", SelectOption, driver);
        }

        [ActionMethod]
        public void OpenFaxPhoneNumbeDropDown()
        {
            this.ClickPageTitle();
            this.ClickPageTitle();
            UICommon.OpenSelectOptionDropDown("rta_faxid", driver);
        }

        [ActionMethod]
        public bool VerifyFaxOptionPresent(string SelectOption)
        {
            return UICommon.VerifySelectOptionPresent("rta_faxid", SelectOption, driver);
        }

        [ActionMethod]
        public void SetFaxNumber(string Number)
        {
            this.ClickPageTitle();
            UICommon.SetSearchableListValue("rta_faxid", Number, driver);
        }

        [ActionMethod]
        public void ClearFaxNumber()
        {
            UICommon.ClearListBoxValue("rta_faxid", driver);
        }

        [ActionMethod]
        public bool VerifyElementLocked(string CssSelector)
        {
            return UICommon.VerifyElementLocked(CssSelector, driver);
        }


        internal void ClickSaveFooter()
        {
            UICommon.ClickSaveFooter(driver);
        }
        public IWebElement GetAuditHistoryTable()
        {
            RefreshPageFrame.RefreshPage(driver, frameId);
            this.driver.SwitchTo().Frame(auditFRAME);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            IWebElement parent = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("crmGrid_divDataArea")));

            IWebElement webElementBody = parent.FindElement(By.Id("gridBodyTable"));
            return webElementBody;
        }
    }
}
