using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.TestTools.UITesting;
using RTA.Automation.CRM.UI;
using RTA.Automation.CRM.Environment;
using OpenQA.Selenium;
using RTA.Automation.CRM.Pages;
using System.Collections.ObjectModel;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reflection;
using RTA.Automation.CRM.DataSource;
using System.Windows.Forms;
using System.Threading;

namespace RTA.Automation.CRM.Tests
{
    [TestClass]
    public class CRMTenancyRequestBL : BaseTest
    {

        [TestMethod]
        [TestProperty("TestcaseID", "9875")]
        public void ATC9875_CRMTenancyRequestBLManagementTypeIsMandatory()
        {
            //Login as RBS Operations Standard User Role.
            User user = this.environment.GetUser(SecurityRole.RBSOfficer);
            new LoginDialog().Login(user.Id, user.Password);

            //Navigate to: RBS group > Tenancy Requests
            HomePage homePage = new HomePage(driver);
            homePage.HoverCRMRibbonTab();
            homePage.ClickRBSRibbonButton();
            homePage.HoverRBSRibbonTab();
            homePage.ClickRtaTenancyRequestRibbonButton();

            //To finish when I find out where new is on the page - GD

        }
    }
}

