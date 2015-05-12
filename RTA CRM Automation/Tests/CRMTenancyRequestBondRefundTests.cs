using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.TestTools.UITesting;
using RTA.Automation.CRM.UI;
using RTA.Automation.CRM.Environment;
using OpenQA.Selenium;
using System.Threading;
using RTA.Automation.CRM.Pages;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reflection;
using RTA.Automation.CRM.DataSource;


namespace RTA.Automation.CRM.Tests
{
    [TestClass]
    public class CRMTenancyRequestBondRefundTests : BaseTest
    {
       

        [TestInitialize]
        public void DataSourceSetup()
        {
            //open excel
            MyApp = new Excel.Application();
            MyApp.Visible = false;

        }

        //The manual steps for this test need review.  Are incomplete/inconsistent.
        [TestMethod]
        [TestProperty("TestID","9204")]
        public void ATC9204_CRMInputDataToRefundRequestsAC1InputAdditionalData()
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

            TenancyRequestsSearchPage tenancyRequestSearchPage = new TenancyRequestsSearchPage(driver);

            //Set view to Active Bond Refunds.
            tenancyRequestSearchPage.SetPageFilterList("Active Bond Refunds");

            //Double-click on record to open it.
            UICommon.DoubleClickElement(tenancyRequestSearchPage.GetSearchResultRow(), driver);

            //Ensure following fields are editable and mandatory i.e. Ensure cannot save record when cleared:
            //Assert: This field is Mandatory - Amount to be refunded per contributor (M)

            /*
             * Ensure following fields are editable and optional and requirements are reflected in software:
                - Date Tenants/residents vacated (O)
                - Expiry date of notice (O)
                - Forwarding address of every contributor (O)
                - date signed by every contributor and the managing party (O) 
                - Details of Claim with Amounts (O, Claim reason list - see Picture+free text - 300 (?)
                +) 
                Amounts are optional, a CRM user should be able to choose multiple reasons associated with the full amount of claim
             */

            //Modify Address and dollar values fields.

            //Select the [AUDIT HISTORY] item from the entity navigation menu in the global ribbon.

            //Inspect audit history.


        }
       
    }
}
