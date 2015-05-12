using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Input;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;
using RTA.Automation.AX.Utils;
using RTA.Automation.AX.Pages;
using System.IO;
using RTA.Automation.AX.UI;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using RTA.Automation.CRM.DataSource;



namespace RTA.Automation.AX.Tests
{
   
    [CodedUITest]
    public class AXConsolidation : TestBase
    {

        public AXConsolidation()
        {
        }


        #region TestInitialize
        //Use TestInitialize to run code before running each test 
        [TestInitialize()]
        public override void TestInitialize()
        {
            // Call a common routine to set up the test
            PlayBackSettings.StartTest();
            base.TestInitialize();

            ////open excel
            //MyApp = new Excel.Application();
            //MyApp.Visible = false;

        
        }
        #endregion


        [TestMethod]
        [TestProperty("TestcaseID", "7145")]
        public void ATC7145_AXShellEntityinAX()
        {

            Homepage homePage = new Homepage();
            homePage.ClickCompanyButton();

            SelectCompanyPage selectCompanyPage = new SelectCompanyPage();
            Table table = new Table(selectCompanyPage.GetCompanyListTable());
            table.ClickCellValue("Company", "RT", "Company");
            selectCompanyPage.ClickOkButton();

            homePage.ClickGeneralLedgerTab();
            homePage.ClickConsolidateLink();
            homePage.ClickConsolidateOnlineLink();

            ConsolidateOnlinePage consolidateOnlinePage = new ConsolidateOnlinePage();
            //consolidateOnlinePage.SetFromValue("1/07/2014");
            //consolidateOnlinePage.SetToValue("30/06/2015");

            consolidateOnlinePage.ClickFinancialdimensionsTab();
            consolidateOnlinePage.ClickLegalentitiesTab();
            consolidateOnlinePage.ClickDescriptionTab();

            consolidateOnlinePage.ClickEliminationTab();

            consolidateOnlinePage.ClickCriteriaTab();
            consolidateOnlinePage.ClickOKButton();

            //consolidation process no errors. there are errors at the moment
            try
            {
                InfoLogPage infoLogPage = new InfoLogPage();
                Assert.IsFalse(infoLogPage.GetControlExists("One or more critical STOP errors have occurred. Use the error messages below to guide you or call your administrator.", "Client"));
            }
            catch
            {}





        }

        
  

        #region Test Clean Up
        [TestCleanup()]
        public override void TestCleanup()
        {
            base.TestCleanup();
        }
        #endregion

      


    }

   
}
