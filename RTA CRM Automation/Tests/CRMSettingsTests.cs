using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.TestTools.UITesting;
using RTA.Automation.CRM.UI;
using RTA.Automation.CRM.Environment;
using OpenQA.Selenium;
using RTA.Automation.CRM.Pages;
using System.Threading;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Input;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;
using System.Collections.ObjectModel;

namespace RTA.Automation.CRM.Tests
{
    [TestClass]
    public class CRMSettingsTests : BaseTest
    {

        public CRMSettingsTests()
        {
        }

        [TestMethod]
        [TestProperty("TestcaseID", "6641")]
        public void ATC6641_CRMInvestigationCaseCasePendingNominatednumberofmonthscanbeconfigured()
        {

            //Login in as role
            User user = this.environment.GetUser(SecurityRole.SystemAdministrator);
            new LoginDialog().Login(user.Id, user.Password);

            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickSettingsRibbonButton();
            homePage.HoverSettingsRibbonTab();
            homePage.ClickRightScrollRibbonButton();
            homePage.ClickProcessesRibbonButton();

            ProcessesSearchPage processesSearchPage = new ProcessesSearchPage(driver);

            //processesSearchPage.ClickProcessesViewButton();
            
            processesSearchPage.SetProcessesSearchText("GetConfigurationValues");
            Table table = new Table(processesSearchPage.GetSearchResultTable());
            StringAssert.Contains(table.GetCellValue("Process Name", "GetConfigurationValues", "Process Name"), "GetConfigurationValues");
            
            string BaseWindow = driver.CurrentWindowHandle; //Records the current window handle
            
            table.ClickCellValue("Process Name", "GetConfigurationValues", "Process Name");
            

            //Enter Request Party details

            driver = UICommon.SwitchToNewBrowserWithTitle(driver, BaseWindow, "Process:");

            string title = driver.Title;

            ProcessesPage processPage = new ProcessesPage(driver);


            processPage.ClickDeactivateButton();
            WarningDialogueFramePage warningDialogueFramePage = new WarningDialogueFramePage(driver);
            warningDialogueFramePage.ClickProcessBeginButton();

            processPage.ClickActivateButton();
            warningDialogueFramePage = new WarningDialogueFramePage(driver);
            warningDialogueFramePage.ClickProcessBeginButton();

            processPage.ClickDeactivateButton();
            warningDialogueFramePage = new WarningDialogueFramePage(driver);
            warningDialogueFramePage.ClickProcessBeginButton();

            processPage.ClickActivateButton();
            warningDialogueFramePage = new WarningDialogueFramePage(driver);
            warningDialogueFramePage.ClickProcessBeginButton();

            processPage.ClickCloseButton();

            driver = driver.SwitchTo().Window(BaseWindow);


        }

       
       
   
    }
}
