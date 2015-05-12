using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using RTA.Automation.CRM.Pages;
using RTA.Automation.CRM.Utils;
//using System.Linq;

namespace RTA.Automation.CRM.UI
{
    public class UICommon
    {
        public static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;

        public static string FindVisibleIFrame(IWebDriver driver)
        {
            WaitForPageToLoad.WaitToLoad(driver);
            
            driver.SwitchTo().DefaultContent();
            ReadOnlyCollection<IWebElement> iframes = driver.FindElements(By.TagName("iframe"));
                foreach (IWebElement frame in iframes)
                {
                    if (iframes.Count == 1)
                    {
                        return frame.GetAttribute("id");
                    }
                    else
                    {
                        if (frame.GetAttribute("style").Contains("visibility: visible;"))
                        {
                            return frame.GetAttribute("id");
                        }

                    }
                }
                throw new Exception("No frame found");
           
        }


        
        public static void ClickSaveButton(IWebDriver d)
        {

            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[alt='Save']")));
            Actions action = new Actions(d);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).Release().Build().Perform();
            Thread.Sleep(3000);


        }

        public static void ClickStartDialogButton(IWebDriver d)
        {

            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[alt='Start Dialog']")));
            Actions action = new Actions(d);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).Release().Build().Perform();
            Thread.Sleep(3000);
        }

        public static void ClickAddToQueueButton(IWebDriver d)
        {

            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[alt='Add to Queue']")));
            Actions action = new Actions(d);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).Release().Build().Perform();
            Thread.Sleep(3000);
        }
        

        public static void ClickSelectValueForQueueButton(IWebDriver d)
        {

            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[alt='Click to select a value for Queue.']")));
            Actions action = new Actions(d);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).Release().Build().Perform();
            Thread.Sleep(3000);
        }

        public static void ClickSaveCloseButton(IWebDriver d)
        {

            
            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[alt='Save & Close']")));
            Actions action = new Actions(d);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).Release().Build().Perform();
            Thread.Sleep(3000);
            


        }

        public static void ClickClosePhoneCallButton(IWebDriver d)
        {

            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[alt='Close Phone Call']")));
            Actions action = new Actions(d);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            action.MoveToElement(elem).Release().Build().Perform();
            Thread.Sleep(3000);
        }

     
        public static string GetNewReferenceNumber(IWebDriver d)
        {

            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            wait.Until(de => !de.FindElement(By.CssSelector("#FormTitle")).Text.Contains("New"));
            IWebElement elem = d.FindElement(By.Id("FormTitle"));
            string value = elem.Text;
            return value;
        }


        public static void ClickNewButton(IWebDriver d)
        {
            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[alt='New']")));
            Actions action = new Actions(d);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).Release().Build().Perform();
            Thread.Sleep(3000);
        }

        public static void ClickImageButton(IWebDriver d, string altSelector)
        {
            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[alt='" + altSelector + "']")));
            Actions action = new Actions(d);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).Release().Build().Perform();
            Thread.Sleep(3000);
        }
         public static void ClickTableRefreshButton(IWebDriver d)
        {


            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("grid_refresh")));
            Actions action = new Actions(d);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).Release().Build().Perform();
            Thread.Sleep(3000);
        }

        

        
        public static IWebElement GetSearchResultTable(IWebDriver d)
        {

            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement webElementBody = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("gridBodyTable")));
            return webElementBody;
        }

        public static IWebElement GetHeaderSearchResultTable(IWebDriver d)
        {

            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement webElementBody = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("crmGrid_gridBar")));
            return webElementBody;
        }

       
        public static void SetSearchText(string elementID, string searchValue, IWebDriver d)
        {
            //crmGrid_findCriteria
            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement element = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(elementID)));
            d.FindElement(By.Id(elementID));
            element.Clear();
            element.SendKeys(searchValue.ToString());
            element.SendKeys(Keys.Enter);

        }
        public static void SetSearchTextCss(string CssSelector, string searchValue, IWebDriver d)
        {

            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement element = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#" + CssSelector)));
            d.FindElement(By.CssSelector("#" + CssSelector));
            element.Clear();
            element.SendKeys(searchValue.ToString());
            element.SendKeys(Keys.Enter);

        }

        public static void ClickPageTitle(IWebDriver d)
        {
            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("FormTitle")));
            Actions action = new Actions(d);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).Release().Build().Perform();
           
        }

        public static string GetPageTitle(IWebDriver d)
        {
            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("FormTitle")));
            return elem.Text;

        }

        public static void ClickSaveFooter(IWebDriver d)
        {
            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("savefooter_statuscontrol")));
            Actions action = new Actions(d);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).Release().Build().Perform();
           
        }

         
        public static void DoubleClickElement(IWebElement element, IWebDriver d)
        {
            Actions action = new Actions(d);
            action.MoveToElement(element).DoubleClick().Build().Perform();
        }

        public static void ClickElementWithId(string elementID, IWebDriver d)
        {
            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(elementID)));
            Actions action = new Actions(d);
            action.MoveToElement(elem).Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(2000);
            action.MoveToElement(elem).Release().Build().Perform();

        }

       

        public static string GetAlertMessage(IWebDriver d)
        {
            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            wait.Until((de) => d.SwitchTo().Alert());

            string alertMessage = d.SwitchTo().Alert().Text;
            d.SwitchTo().Alert().Accept();
            return alertMessage;

        }

        public static IWebDriver SwitchToNewBrowser(IWebDriver d, string BaseWindow)
        {
            //*****************This needs to be moved out of here********************************************
            //string NewWindow = ""; //prepares for the new window handle

            ReadOnlyCollection<string> handles = null;
            for (int i = 1; i < 30; i++)
            {
                if (d.WindowHandles.Count == 1)
                { Thread.Sleep(1000); }
                else { break; }
            }
            handles = d.WindowHandles;
           
            d = d.SwitchTo().Window(handles.Last());
            return d;

            //foreach (string handle in handles)
            //{
            //    var Handles = handle;
            //    if (BaseWindow != handle)
            //    {
            //        NewWindow = handle;

            //        d = d.SwitchTo().Window(NewWindow);
            //        return d;
            //    }
            //} throw new Exception("Error switching to new browser");
            //**********************************************************************************************
        }

        public static IWebDriver SwitchToNewBrowserWithTitle(IWebDriver d, string BaseWindow, string title)
        {
            string NewWindow; //prepares for the new window handle
            ReadOnlyCollection<string> handles = null;
            for (int i = 1; i < 30; i++)
            {
                if (d.WindowHandles.Count == 1)
                { Thread.Sleep(1000); }
                else { break; }
            }
            handles = d.WindowHandles;


            foreach (string handle in handles)
            {
                //var Handles = handle;
                if (BaseWindow != handle)
                {
                    NewWindow = handle;
                    WaitForPageToLoad.WaitToLoad(d);
                    if (d.SwitchTo().Window(handle).Title.Contains(title))
                    {
                         return d;
                    }
                }
            } throw new Exception("Error switching to new browser");

        }

        public static void ClickRibbonButton(string buttonID, IWebDriver d)
        {
            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(buttonID)));
            Actions action = new Actions(d);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).Release().Build().Perform();
            Thread.Sleep(2000);

        }

        public static void ClickRibbonTab(string cssTabID, IWebDriver d)
        {
            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(cssTabID)));
            Actions action = new Actions(d);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).Release().Build().Perform();
            Thread.Sleep(2000);

        }



        public static void HoverXRibbonTab(string name, IWebDriver d)
        {
            d.SwitchTo().DefaultContent();
            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//span[@id='TabNode_tab0Tab' and contains(@title, '" + name + "')]")));  //
            Actions action = new Actions(d);
            action.MoveToElement(elem).Build().Perform();
            
            
        }


        public static void HoverRibbonTab(string ribbonTab, IWebDriver d)
        {
            d.SwitchTo().DefaultContent();
            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(ribbonTab)));
            Actions action = new Actions(d);
            action.MoveToElement(elem).Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).Build().Perform();
        }


        public static string GetTextFromElement(string CssSelector, IWebDriver d)
        {

            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(CssSelector)));
            return elem.Text;
        }

        public static string GetElementProperty(string CssSelector, string property, IWebDriver d)
        {

            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(CssSelector)));
            return elem.GetAttribute(property);

        }

        public static bool GetElementNotExistTrue(string CssSelector, IWebDriver d)
        {
            ReadOnlyCollection<IWebElement> elemList = d.FindElements(By.CssSelector(CssSelector));
            if (elemList.Count == 0)
                return true;
            else
                return false;

        }
        
        public static void SetSelectListValue(string elementId, string listValue, IWebDriver d)
        {
            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#" + elementId + ">div"))).Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//select[@id='" + elementId + "_i']/option[text()='" + listValue + "']"))).Click();

        }

        public static bool DoesSelectListValueExist(string elementId, string listValue, IWebDriver d)
        {
            try
            {
                WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
                wait.Until(ExpectedConditions.ElementExists(By.Id(elementId))).Click();
                wait.Until(ExpectedConditions.ElementExists(By.XPath("//select[@id='" + elementId + "_i']/option[text()='" + listValue + "']")));
                return true;
            }
            catch
            {
                return false;
            }

            
        }

        public static void SetTextBoxValue(string elementId, string textValue, IWebDriver d)
        {
            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#" + elementId + ">div"))).Click();
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(elementId + "_i")));
            elem.Clear();
            elem.SendKeys(textValue);
            elem.SendKeys(Keys.Enter);
        }

        public static void ClearTextBoxValue(string elementId, IWebDriver d)
        {
            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#" + elementId + ">div")));
            Actions actions = new Actions(d);
            actions.MoveToElement(elem,0,0).Click().Build().Perform();
            
            Thread.Sleep(500);
            IWebElement el = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(elementId + "_i")));
            actions.MoveToElement(elem, 0, 0).Click().Build().Perform();
            el.Clear();
            el.SendKeys(Keys.Backspace);
        }

        public static void ClearListBoxValue(string elementId, IWebDriver d)
        {
            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(elementId)));
            Actions actions = new Actions(d);
            actions.MoveToElement(elem).Build().Perform();
            Thread.Sleep(500);
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#" + elementId + ">div>div>img"))).Click();
            elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(elementId + "_i")));
            elem.Click();
            elem.SendKeys(Keys.Backspace);
            elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(elementId + "_ledit")));
            elem.SendKeys(Keys.Clear);
        }


        public static void SetDateValue(string elementId, string dateValue, IWebDriver d)
        {
            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#"+elementId+">div"))).Click();
            WebDriverWait waitforDateInput = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement tableElement = waitforDateInput.Until(ExpectedConditions.ElementIsVisible(By.Id(elementId+"_i")));
            IWebElement elem = tableElement.FindElement(By.Id("DateInput"));
            elem.Clear();
            elem.SendKeys(dateValue.ToString());
            elem.SendKeys(Keys.Enter);
            
        }

        public static void SetSearchableListValue(string elementId, string listValue, IWebDriver d)
        {

            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(elementId)));
            Actions actions = new Actions(d);
            actions.MoveToElement(elem).Build().Perform();
            Thread.Sleep(500);
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#" + elementId + ">div>div>img"))).Click();
            Thread.Sleep(1000);
            elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(elementId +"_i")));
            elem.Click();
            elem.SendKeys(Keys.Backspace);
            elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(elementId + "_ledit")));
            elem.SendKeys(Keys.Clear);
            elem.SendKeys(listValue);
            elem.SendKeys(Keys.Enter);
            if (listValue.Substring(0, 1) == "*")
            {
                listValue = listValue.Substring(1);
            }
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//li[a[contains(@title,'" + listValue + "')] and parent::*[@id='" + elementId + "_i_IMenu']]"))).Click();
        }

        public static void SetSearchableMultiListValue(string elementId, string listValue, IWebDriver d)
        {

            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(elementId)));
            Actions actions = new Actions(d);
            actions.MoveToElement(elem).Build().Perform();
            Thread.Sleep(500);
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#" + elementId + ">div>div>img"))).Click();
            elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(elementId + "_i")));
            elem.Click();
            elem.SendKeys(Keys.Backspace);
            elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(elementId + "_ledit_multi")));
            elem.SendKeys(Keys.Clear);
            elem.SendKeys(listValue);
            elem.SendKeys(Keys.Enter);
            IWebElement parent = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(elementId + "_i_IMenu")));
            parent.FindElement(By.XPath("//li/a[contains(@title,'" + listValue + "')]")).Click();
        }

     

        public static void ClickSearchableListValue(string elementId, string listValue, IWebDriver d)
        {

            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(elementId)));
            Actions actions = new Actions(d);
            actions.MoveToElement(elem).Build().Perform();
            Thread.Sleep(500);
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#" + elementId + ">div>div>img"))).Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id(elementId + "_i"))).Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//li[a[contains(@title,'" + listValue + "')] and parent::*[@id='" + elementId + "_i_IMenu']]"))).Click();
        }

        public static bool GetPaymentRefernceRefreshTable(string reference, Table t, IWebDriver d)
        {


            int i = 1;
            while (i <= 60) //waits for 60sec
            {

                if (t.GetCellValue("Name", reference, "Status Reason") == "Pending Financials")
                {
                    return true;

                }
                else
                {
                    Thread.Sleep(1000);
                    WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
                    wait.Until(ExpectedConditions.ElementIsVisible(By.Id("grid_refresh"))).Click();
                    i++;
                }

            } throw new Exception(String.Format("Status Reason has not changed to Pending Financials"));
        }

        public static void ClickTabHeader(string elementName, IWebDriver d)
        {
            Actions action = new Actions(d);
            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("div[name='" + elementName + "'] a.ms-crm-InlineTabHeaderText h2")));
            elem.Click();
        }

        public static bool CheckElementExists(string CssSelector, IWebDriver d)
        {
            IList<IWebElement> eleList = d.FindElements(By.CssSelector(CssSelector));
            if (eleList.Count > 0)
            {
                return true;
            }
            return false;
        }

        /*
         * Warning message on top of page e.g. Physical address is blank etc.
         * ************************************************************************
         */

        public static string GetWarningMessage(string cssSelector, IWebDriver d) 
        {
             WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("div[notificationid='" + cssSelector + "']")));
            return elem.Text;
        }

        public static bool VerifyWarningMessagePresent(string cssSelector, IWebDriver d)
        {
            IList<IWebElement> warningList = d.FindElements(By.CssSelector("div[notificationid='" + cssSelector + "']"));
            if (warningList.Count > 0)
            {
                return true;
            }
            return false;
        }

        public static void ClickDeactivateButton(IWebDriver d)
        {
            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[alt='Deactivate']")));
            Actions action = new Actions(d);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).Release().Build().Perform();
            Thread.Sleep(3000);
        }

        public static bool VerifyElementLocked(string CssSelector, IWebDriver d)
        {
            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#" + CssSelector)));
            try {
                String value = elem.GetAttribute("data-controlmode");
                if (value.Contains("locked"))
                {
                    return true;
                }
            } 
            catch (Exception) {}
            return false;
        }

        public static void ClickAddButton(IWebDriver d, string ButtonID)
        {
            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("FormTitle"))).Click();

            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(ButtonID)));// "Activities_addImageButtonImage")));

            Actions action = new Actions(d);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).Release().Build().Perform();
        }

        public static bool ClickAddActivity(string Activity, IWebDriver d)
        {
            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement parent = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("moreActivitiesList")));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//li[span/a/img[@alt='"+Activity+"'] and parent::*[@id='moreActivitiesList']]")));
            for(int i = 0; i<=30; i++)
            {
                if(d.WindowHandles.Count == 1)
                {
                    Actions action = new Actions(d);
                    action.MoveToElement(elem).ClickAndHold().Build().Perform();
                    Thread.Sleep(1000);
                    action.MoveToElement(elem).Release().Build().Perform();
                    Thread.Sleep(3000);
                    
                }
                else
                {
                    return true;
                    
                }
            } throw new Exception ("Can not click on Activity");
      
            
        }

        public static void ClickSeeRecordsAssociatedWithThisViewButton(string association, IWebDriver d)
        {
            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(association+"_openAssociatedGridViewImageButtonImage")));
            Actions action = new Actions(d);
            action.MoveToElement(elem).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(elem).Release().Build().Perform();
            Thread.Sleep(2000);
        }

        public static string GetRandomString(int length)
        {
            string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
            var randomString = new StringBuilder();
            var random = new Random();

            for (int i = 0; i < length; i++)
                randomString.Append(chars[random.Next(chars.Length)]);

            return randomString.ToString().ToUpper();
        }

        public static void SetRadioButton(IWebDriver d, string elementid)
        {
             WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
             IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(elementid)));
             if(!elem.Selected)
             {
                 elem.Click();
             }
        }

        public static void OpenSelectOptionDropDown(string elementId, IWebDriver d)
        {
            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(elementId)));
            Actions actions = new Actions(d);
            actions.MoveToElement(elem).Build().Perform();
            Thread.Sleep(500);
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#" + elementId + ">div>div>img"))).Click();
            elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(elementId + "_i")));
            elem.Click();
            elem = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(elementId + "_ledit")));
            elem.SendKeys(Keys.Enter);
        }
        public static bool VerifySelectOptionPresent(string elementId, string listValue, IWebDriver d)
        {


            IReadOnlyCollection<IWebElement> ListOptions = d.FindElements(By.CssSelector("#" + elementId + "_i_IMenu>li"));  

            if (ListOptions.Count > 0)
            {
                foreach (IWebElement ListItem in ListOptions)
                {
                    if (ListItem.Text.Contains(listValue))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        public static void SetPageFilterList(string value, IWebDriver d)
        {
            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#crmGrid_SavedNewQuerySelector>span"))).Click();
            IWebElement parent = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("Dialog_0")));
            parent.FindElement(By.XPath("//li[a[contains(@title,'" + value + "')]]")).Click();
        }


        public static void SetQueue(string QueueName, IWebDriver d)
        {
            UICommon.ClickSelectValueForQueueButton(d);

            // Set the Queue name
            IList<IWebElement> queueList = d.FindElements(By.ClassName("ms-crm-MenuItem-NoOutline"));

            foreach (IWebElement queue in queueList)
            {
                if (queue.Text.Contains(QueueName))
                {
                    queue.Click();
                    return;
                }
            }
            throw new Exception("Queue not found!!!");
        }

        public static void ClickDialogAddButton(IWebDriver d)
        {
            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("butBegin"))).Click();
        }

        public static void SetConnectList(string connectType, IWebDriver d)
        {

            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            IWebElement elemList = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("connection|OneToMany|SubGridAssociated|Mscrm.SubGrid.connection.AddConnection")));

            IWebElement arrow = elemList.FindElement(By.XPath("//li/span/a/img[contains(@src,'CommandBarMenuDown.png')]"));//  //span/a/img[contains(@src,'/CommandBarMenuDown.png']"));
            Actions action = new Actions(d);
            action.MoveToElement(arrow).ClickAndHold().Build().Perform();
            Thread.Sleep(1000);
            action.MoveToElement(arrow).Release().Build().Perform();

            IWebElement elemMenu = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("connection_OneToMany_SubGridAssociated_Mscrm_SubGrid_connection_AddConnectionMenu")));
            elemMenu.FindElement(By.XPath("//li/span/a[contains(.,'" + connectType + "')]")).Click();

        }

        public static void ClickHomePageTile(string tileSrcPath, IWebDriver d)
        {
            d.SwitchTo().Frame("contentIFrame0");

            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//ul[@id='uxMenu']/li[3]")));
            IWebElement elem = d.FindElement(By.XPath("//ul[@id='uxMenu']/li[3]"));
            d.FindElement(By.XPath("*//img[contains(@src,' " + tileSrcPath + "')]")).Click();

            d.SwitchTo().DefaultContent();
        }

        public static void ClickClientServicesHomePageTile(string tileSrcPath, IWebDriver d)
        {
            d.SwitchTo().Frame("contentIFrame0");

            WebDriverWait wait = new WebDriverWait(d, TimeSpan.FromSeconds(waitsec));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//ul[@id='uxMenu']/li[1]")));
            IWebElement elem = d.FindElement(By.XPath("//ul[@id='uxMenu']/li[1]"));
            d.FindElement(By.XPath("*//img[contains(@src,' " + tileSrcPath + "')]")).Click();

            d.SwitchTo().DefaultContent();
        }
    }
}
