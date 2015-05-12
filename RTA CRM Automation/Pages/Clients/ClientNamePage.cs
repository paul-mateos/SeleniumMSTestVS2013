using ActionWordsLib.Attributes;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Interactions.Internal;
using System.Collections.ObjectModel;
using System.Threading;
using RTA.Automation.CRM.Utils;
using RTA.Automation.CRM.UI;

namespace RTA.Automation.CRM.Pages
{
   [ActionClass]
   public class ClientNamePage : IFramePage
    {
        //public static string WINDOW = "Client Name: New Client Name - Microsoft Dynamics CRM";
        private static string frameId = "contentIFrame0";
        //private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;
        private static int waitsec = Properties.Settings.Default.LONG_WAIT_SECONDS;
        private static string pageTitle = "Client Name";

        //protected IWebDriver driver = null;

        public ClientNamePage(IWebDriver driver)
            : base(driver, ClientNamePage.frameId)
        {

            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);

        }

        /*
        * Client Title
        * ************************************************************************
        */

        //[ActionMethod]
        //public void ClickTitleList()
        //{
        //    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
        //    //id="rta_name_titleid"
        //    //<input class="ms-crm-InlineInput ms-crm-InlineLookupEdit" id="rta_name_titleid_ledit" style="-ms-ime-mode: auto;" type="text" maxlength="1000" ime-mode="auto">
        //    wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_name_titleid"))).Click();
        //    wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_name_titleid_ledit"))).GetAttribute("id");
            
        //}

        [ActionMethod]
        public void SetTitleListValue(string listValue)
        {
            //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            //wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_name_titleid_ledit")));
            //IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_name_titleid_ledit")));
            //elem.Clear();
            //elem.SendKeys(listValue);

            UICommon.SetSelectListValue("rta_name_titleid", listValue, driver);
        }

        //[ActionMethod]
        //public string GetClientTitle()
        //{

        //    //<div style="display: none;" class="ms-crm-Inline-Value ms-crm-Inline-Lookup"><span tabindex="-1" title="MISS" onkeydown="Mscrm.ReadFormUtilities.keyDownHandler(new Sys.UI.DomEvent(event));" style="display: inline-block;" oid="{7A225DCE-6BB5-E311-80C0-005056B949EF}" otype="10041" otypename="rta_config_person_name_title" resolved="true" onclick="Mscrm.ReadFormUtilities.openLookup(true, new Sys.UI.DomEvent(event));" role="link" class="ms-crm-Lookup-Item" contenteditable="false">MISS<div class="ms-crm-Inline-GradientMask"></div></span><span style="display: none;" contenteditable="false">MISS</span><div style="display: none;" class="ms-crm-Inline-EditIcon"><img src="/_imgs/imagestrips/transparent_spacer.gif" class="ms-crm-ImageStrip-search_normal ms-crm-InlineLookupEdit ms-crm-EditLookup-Image" alt=""></div></div>
        //    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
        //    IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//./div/span[contains(@otypename,'rta_config_person_name_title')]")));
        //    string value = elem.Text;
        //    return value;
        //}

        //[ActionMethod]
        //public void ClickGivenName()
        //{
        //    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
        //    //<input class="ms-crm-InlineInput ms-crm-InlineLookupEdit" id="rta_name_titleid_ledit" style="-ms-ime-mode: auto;" type="text" maxlength="1000" ime-mode="auto">
        //    wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_given_name"))).Click();
        //    wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_given_name_i"))).GetAttribute("id");
            
        //}


        
        [ActionMethod]
        public void SetGivenNameValue(String Value)
        {
            //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            //IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_given_name_i")));
            //elem.Clear();
            //elem.SendKeys(Value);

            UICommon.SetTextBoxValue("rta_given_name", Value, driver);

        }

        //[ActionMethod]
        //public string GetClientName()
        //{
        //    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
        //    IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_given_name_i")));
        //    string value = elem.Text;
        //    return value;
        //}

        //[ActionMethod]
        //public void ClickMiddleName()
        //{
        //    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
        //    //<div class="ms-crm-Inline-Edit" style="display: block;"><input defaultvalue="" style="ime-mode: auto;" controlmode="normal" aria-labelledby="rta_middle_name_c rta_middle_name_w" class="ms-crm-InlineInput" title="" maxlength="50" attrpriv="7" attrname="rta_middle_name" id="rta_middle_name_i" type="text"></div>
        //    wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_middle_name"))).Click();
        //    wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_middle_name_i"))).GetAttribute("id");
           
        //}

        [ActionMethod]
        public void SetMiddleNameValue(String Value)
        {
            //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            ////wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_middle_name_i")));
            //IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_middle_name_i")));
            //elem.Clear();
            //elem.SendKeys(Value);

            UICommon.SetTextBoxValue("rta_middle_name", Value, driver);

        }

        //[ActionMethod]
        //public void ClickFamilyName()
        //{
        //    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
        //    //<div class="ms-crm-Inline-Edit"><input title="" class="ms-crm-InlineInput" id="rta_family_name_i" aria-labelledby="rta_family_name_c rta_family_name_w" style="-ms-ime-mode: auto;" type="text" maxlength="100" attrName="rta_family_name" attrPriv="7" controlmode="normal"></div>
        //    wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_family_name"))).Click();
        //    wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_family_name_i"))).GetAttribute("id");
            
        //}

        [ActionMethod]
        public void SetFamilyNameValue(String Value)
        {
            //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            ////wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_family_name_i")));
            //IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_family_name_i")));
            //elem.Clear();
            //elem.SendKeys(Value);

            UICommon.SetTextBoxValue("rta_family_name", Value, driver);

        }
        //[ActionMethod]
        //public void ClickSuffixList()
        //{
        //    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
        //    //id="rta_name_suffixid"
        //    //<input style="ime-mode: auto;" maxlength="1000" id="rta_name_suffixid_ledit" ime-mode="auto" class="ms-crm-InlineInput ms-crm-InlineLookupEdit" type="text">
        //    wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_name_suffixid"))).Click();
        //    wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_name_suffixid_ledit"))).GetAttribute("alt");
            
        //}

        [ActionMethod]
        public void SetSuffixListValue(string listValue)
        {
            //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            ////wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_name_suffixid_ledit")));
            //IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_name_suffixid_ledit")));
            //elem.Clear();
            //elem.SendKeys(listValue);
            UICommon.SetSearchableListValue("rta_name_suffixid", listValue, driver);
        }

       [ActionMethod]
       public string GetFormTitle()
       {
           //frameId = UICommon.FindVisibleIFrame(driver);
           //RefreshPageFrame.RefreshPage(driver, frameId);
           //System.Diagnostics.Debug.WriteLine("frameId: " + frameId.ToString());
           ////<div class="ms-crm-Form-Title-Data autoellipsis" id="FormTitle"><h1 title="MISS JACK AAAAAA FFFFFFFF BM" class="ms-crm-TextAutoEllipsis">MISS JACK AAAAAA FFFFFFFF BM</h1></div>
           //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
           //IWebElement element = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#FormTitle")));
           //System.Diagnostics.Debug.WriteLine("FormTitle text: " + element.Text);
           //string value = element.Text;
           //return value;

           return UICommon.GetPageTitle(driver);
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

        internal IWebDriver SwitchNewBrowser(IWebDriver driver, string BaseWindow)
        {
            return UICommon.SwitchToNewBrowser(driver, BaseWindow);
        }

    }
}
