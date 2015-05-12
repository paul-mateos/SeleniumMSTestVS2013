﻿using ActionWordsLib.Attributes;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RTA.Automation.CRM.Utils;
using System.Threading;
using RTA.Automation.CRM.UI;

namespace RTA.Automation.CRM.Pages
{
    [ActionClass]
    public class ClientsSearchPage : IFramePage
    {
        private static string frameId = "contentIFrame0";
        private static int waitsec = Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;
        private static string pageTitle = "Clients";

        public ClientsSearchPage(IWebDriver driver)
            : base(driver, ClientsSearchPage.frameId)
        {
            //Wait for title to be displayed
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait.Until((d) => { return d.Title.Contains(pageTitle); });

            //Switch to main frame when it is visible
            frameId = UICommon.FindVisibleIFrame(driver);
            RefreshPageFrame.RefreshPage(driver, frameId);
            
        }


        /*
        * NewIMG
        * ************************************************************************
        */

        [ActionMethod]
        public void ClickNewClientButton()
        {
            driver.SwitchTo().DefaultContent();
            UICommon.ClickNewButton(driver);
            
            
        }

   
        /*
        * search criteria
        * ************************************************************************
        */

        [ActionMethod]
        public void SetClientSearchText(string searchValue)
        {

            UICommon.SetSearchText("crmGrid_findCriteria", searchValue, driver);

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

        [ActionMethod]
        public void SetPageFilterList(string value)
        {
            UICommon.SetPageFilterList(value, driver);
        }

        [ActionMethod]
        public IWebElement GetHeaderSearchResultTable()
        {
            return UICommon.GetHeaderSearchResultTable(driver);
        }
    }
}
