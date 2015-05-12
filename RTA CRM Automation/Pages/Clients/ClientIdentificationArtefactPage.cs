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
   public class ClientIdentificationArtefactPage : IFramePage
    {
        //public static string WINDOW = "Client Identification Artefact: New Client Identification Artefact - Microsoft Dynamics CRM";
        private static string frameId = "contentIFrame0";
        //private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;
        private static int waitsec = Properties.Settings.Default.LONG_WAIT_SECONDS;
        private static string pageTitle = "Client Identification Artefact";

        //protected IWebDriver driver = null;

        public ClientIdentificationArtefactPage(IWebDriver driver)
            : base(driver, ClientIdentificationArtefactPage.frameId)
        {

            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);

        }

        /*
        * Client Id
        * ************************************************************************
        */

        //[ActionMethod]
        //public void ClickClientIdList()
        //{
        //    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
        //    //id="rta_clientid_titleid"
        //    wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_clientid"))).Click();
        //    wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_clientid_ledit"))).GetAttribute("id");
            
        //}

        //[ActionMethod]
        //public void SetClientIdListValue(string listValue)
        //{
        //    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
        //    wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_clientid_ledit")));
        //    IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_clientid_ledit")));
        //    elem.Clear();
        //    elem.SendKeys(listValue);
        //}

        //[ActionMethod]
        //public string GetClientId()
        //{

        //    //<div style="display: none;" class="ms-crm-Inline-Value ms-crm-Inline-Lookup"><span tabindex="-1" title="MISS" onkeydown="Mscrm.ReadFormUtilities.keyDownHandler(new Sys.UI.DomEvent(event));" style="display: inline-block;" oid="{7A225DCE-6BB5-E311-80C0-005056B949EF}" otype="10041" otypename="rta_config_person_name_title" resolved="true" onclick="Mscrm.ReadFormUtilities.openLookup(true, new Sys.UI.DomEvent(event));" role="link" class="ms-crm-Lookup-Item" contenteditable="false">MISS<div class="ms-crm-Inline-GradientMask"></div></span><span style="display: none;" contenteditable="false">MISS</span><div style="display: none;" class="ms-crm-Inline-EditIcon"><img src="/_imgs/imagestrips/transparent_spacer.gif" class="ms-crm-ImageStrip-search_normal ms-crm-InlineLookupEdit ms-crm-EditLookup-Image" alt=""></div></div>
        //    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
        //    IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_clientid_ledit")));
        //    string value = elem.Text;
        //    return value;
        //}

        //[ActionMethod]
        //public void ClickClientIdProvided()
        //{
        //    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
        //    //<input class="ms-crm-InlineInput ms-crm-InlineLookupEdit" id="rta_name_titleid_ledit" style="-ms-ime-mode: auto;" type="text" maxlength="1000" ime-mode="auto">
        //    wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_identification_providedid"))).Click();
        //    wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_identification_providedid_ledit"))).GetAttribute("id");
            
        //}


        
        [ActionMethod]
        public void SetClientIdProvided(String Value)
        {
            UICommon.SetSearchableListValue("rta_identification_providedid", Value, driver);
            
            //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
            //IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_identification_providedid_ledit")));
            //elem.Clear();
            //elem.SendKeys(Value);

        }

        //[ActionMethod]
        //public string GetClientIdProvided()
        //{
        //    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
        //    IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("rta_identification_providedid_ledit")));
        //    string value = elem.Text;
        //    return value;
        //}

       [ActionMethod]
       public string GetFormTitle(string IdProvided)
       {
           //frameId = UICommon.FindVisibleIFrame(driver);
           //RefreshPageFrame.RefreshPage(driver, frameId);
           //System.Diagnostics.Debug.WriteLine("frameId: " + frameId.ToString());
           //System.Diagnostics.Debug.WriteLine("IdProvided: " + IdProvided);
           ////<div id="FormTitle" class="ms-crm-Form-Title-Data autoellipsis"><h1 class="ms-crm-TextAutoEllipsis" title="BLAIR TEST: Australian Birth Certificate (full)">BLAIR TEST: Australian Birth Certificate (full)</h1></div>
           //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitsec));
           //Thread.Sleep(5000);//Css and xpath does not appear to find the FormTitle element, therefore, add 5 sec sleep workround
           //System.Diagnostics.Debug.WriteLine("FormTitle text: " + driver.FindElement(By.CssSelector("#FormTitle")).Text);
           ////Css and xpath does not appear to find the FormTitle element
           ////IWebElement element = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//./div[contains(@id,'FormTitle')]/h1[contains(@title,'" + IdProvided + "')]")));
           ////wait.Until(de => !de.FindElement(By.CssSelector("#FormTitle")).Text.Contains(IdProvided.Trim()));
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

        internal IWebDriver SwitchNewBrowserWithTitle(IWebDriver driver, string BaseWindow, string pageTitle)
        {
            return UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, pageTitle);
        }

    }
}
