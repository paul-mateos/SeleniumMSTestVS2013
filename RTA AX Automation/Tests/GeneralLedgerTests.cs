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
    public class GeneralLedgerTests : TestBase
    {

      
        
        public GeneralLedgerTests()
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

            //open excel
            MyApp = new Excel.Application();
            MyApp.Visible = false;

        }
        #endregion

        [TestMethod]
        [TestProperty("TestcaseID", "7131")]
        public void ATC7131_AXTaxPosting()
        {

           
            Homepage homePage = new Homepage();
            homePage.ClickCompanyButton();

            SelectCompanyPage selectCompanyPage = new SelectCompanyPage();
            Table table = new Table(selectCompanyPage.GetCompanyListTable());
            table.ClickCellValue("Company", "RTA", "Company");
            selectCompanyPage.ClickOkButton();

            //Input journal
            homePage.ClickHomeTab();
            homePage.ClickGeneralLedgerTab();

            homePage.ClickSalesTaxLink();
            homePage.ClickLedgerPostingGroupsLink();

            LedgerPostingGroupsPage ledgerPostingsGroupsPage = new LedgerPostingGroupsPage();
            
            ledgerPostingsGroupsPage.ClickCloseButton();

            homePage.ClickGeneralLedgerTab();
            homePage.ClickGeneralJournalLink();

            GeneralJournalPage generalJournalPage = new GeneralJournalPage();
            generalJournalPage.ClickNewMenuItem();
            generalJournalPage.SetNameValue("GL");
            //generalJournalPage.SetDescriptionValue("General Journal RTA");
            generalJournalPage.ClickLinesMenuItem();

            JournalVoucherPage journalVoucherPage = new JournalVoucherPage();
            table = new Table(journalVoucherPage.GetJournalValueTable());
            table.ClickCellValue("Account type", "Ledger", "Account");
            journalVoucherPage.SetAccountSeg1Value("74533");
            journalVoucherPage.SetAccountSeg2Value("10");
            journalVoucherPage.SetAccountSeg3Value("1010");

            table.ClickCellValue("Account type", "Ledger", "Debit");

            journalVoucherPage.SetDebitValue("500");
            table.ClickCellValue("Account type", "Ledger", "Offset account");
            journalVoucherPage.SetAccountSeg1Value("15350");
            //Ily
            journalVoucherPage.SetSalesTaxGroupValue("ACQ");
            journalVoucherPage.SetItemGSTGroupValue("GST");
            StringAssert.Contains(journalVoucherPage.GetCalculatedSalesTaxAmountText(), "45.45");

            //Validate approval
            journalVoucherPage.ClickApprovalItemMenuButton();
            journalVoucherPage.ClickReportAsReadyItemMenuItem();
            journalVoucherPage.ClickApprovalItemMenuButton();
            journalVoucherPage.ClickApproveItemMenuItem();


            journalVoucherPage.ClickPostMenuButton();
            journalVoucherPage.ClickPostItemMenuItem();

            InfoLogPage infoLog = new InfoLogPage();
            //Needs to be fixed
            Assert.IsTrue(infoLog.GetTreeItemExists("Number of vouchers posted to the journal:", ""));
            infoLog.ClickCloseButton();

            journalVoucherPage.ClickInquiriesMenuButton();
            journalVoucherPage.ClickVoucherItemMenuItem();

            VoucherTransactionsPage voucherTransactionsPage = new VoucherTransactionsPage();
            journalVoucherPage.ClickCloseButton();
            generalJournalPage.ClickCloseButton();



            //Input journal
            homePage.ClickGeneralLedgerTab();
            homePage.ClickGeneralJournalLink();

            generalJournalPage = new GeneralJournalPage();
            generalJournalPage.ClickNewMenuItem();
            generalJournalPage.SetNameValue("GL");
            //generalJournalPage.SetDescriptionValue("General Journal RTA");
            generalJournalPage.ClickLinesMenuItem();

            journalVoucherPage = new JournalVoucherPage();
            table = new Table(journalVoucherPage.GetJournalValueTable());
            table.ClickCellValue("Account type", "Ledger", "Account");
            journalVoucherPage.SetAccountSeg1Value("74960");
            journalVoucherPage.SetAccountSeg2Value("10");
            journalVoucherPage.SetAccountSeg3Value("1010");

            table.ClickCellValue("Account type", "Ledger", "Debit");

            journalVoucherPage.SetDebitValue("500");
            table.ClickCellValue("Account type", "Ledger", "Offset account");
            journalVoucherPage.SetAccountSeg1Value("15350");


            journalVoucherPage.SetSalesTaxGroupValue("ACQ");
            journalVoucherPage.SetItemGSTGroupValue("INPUT");
            StringAssert.Contains(journalVoucherPage.GetCalculatedSalesTaxAmountText(), "45.45");

            //Validate approval
            journalVoucherPage.ClickApprovalItemMenuButton();
            journalVoucherPage.ClickReportAsReadyItemMenuItem();
            journalVoucherPage.ClickApprovalItemMenuButton();
            journalVoucherPage.ClickApproveItemMenuItem();

            journalVoucherPage.ClickPostMenuButton();
            journalVoucherPage.ClickPostItemMenuItem();

            infoLog = new InfoLogPage();
            //Needs to be fixed
            Assert.IsTrue(infoLog.GetTreeItemExists("Number of vouchers posted to the journal:", ""));
            infoLog.ClickCloseButton();


            journalVoucherPage.ClickCloseButton();
            generalJournalPage.ClickCloseButton();
        }

        [TestMethod]
        [TestProperty("TestcaseID", "7143")]
        public void ATC7143_AXAPRecordofInvoices1expenseaccount()
        {

            Homepage homePage = new Homepage();
            homePage.ClickCompanyButton();

            SelectCompanyPage selectCompanyPage = new SelectCompanyPage();
            Table table = new Table(selectCompanyPage.GetCompanyListTable());
            table.ClickCellValue("Company", "RTA", "Company");
            selectCompanyPage.ClickOkButton();

            homePage.ClickHomeTab();
            homePage.ClickAccountPayableTab();

            //Choose vendor
            homePage.ClickVendorsLink();

            homePage.ClickHomeTab();
            homePage.ClickAccountPayableTab();
            homePage.ClickVendorsSetupLink();


            homePage.ClickGeneralLedgerTab();
            homePage.ClickSalesTaxLink();
            homePage.ClickLedgerPostingGroupsLink();

            LedgerPostingGroupsPage ledgerPostingsGroupsPage = new LedgerPostingGroupsPage();
            ledgerPostingsGroupsPage.ClickCloseButton();


            homePage.ClickGeneralJournalLink();
            GeneralJournalPage generalJournalPage = new GeneralJournalPage();
            generalJournalPage.ClickNewMenuItem();
            generalJournalPage.SetNameValue("GL");
            //generalJournalPage.SetDescriptionValue("General Journal RTA");
            generalJournalPage.ClickLinesMenuItem();

            JournalVoucherPage journalVoucherPage = new JournalVoucherPage();
            table = new Table(journalVoucherPage.GetJournalValueTable());
            table.ClickCellValue("Account type", "Ledger", "Account");
            journalVoucherPage.SetAccountSeg1Value("74533");
            journalVoucherPage.SetAccountSeg2Value("10");
            journalVoucherPage.SetAccountSeg3Value("1010");

            table.ClickCellValue("Account type", "Ledger", "Debit");

            journalVoucherPage.SetDebitValue("500");
            table.ClickCellValue("Account type", "Ledger", "Offset account");
            journalVoucherPage.SetAccountSeg1Value("15350");
            journalVoucherPage.SetSalesTaxGroupValue("ACQ");
            journalVoucherPage.SetItemGSTGroupValue("GST");
            StringAssert.Contains(journalVoucherPage.GetCalculatedSalesTaxAmountText(), "45.45");

            //Validate approval
            journalVoucherPage.ClickApprovalItemMenuButton();
            journalVoucherPage.ClickReportAsReadyItemMenuItem();
            journalVoucherPage.ClickApprovalItemMenuButton();
            journalVoucherPage.ClickApproveItemMenuItem();

            journalVoucherPage.ClickPostMenuButton();
            journalVoucherPage.ClickPostItemMenuItem();

            InfoLogPage infoLogPage = new InfoLogPage();
            Assert.IsTrue(infoLogPage.GetTreeItemExists("Number of vouchers posted to the journal: 1", "Message"));
            infoLogPage.ClickClearButton();
            infoLogPage.ClickCloseButton();

            journalVoucherPage.ClickCloseButton();
            generalJournalPage.ClickCloseButton();


        }


        [TestMethod]
        [TestProperty("TestcaseID", "7292")]
        public void ATC7292_AXValidateTotalAccountsNoentry()
        {

            Homepage homePage = new Homepage();
            homePage.ClickCompanyButton();

            SelectCompanyPage selectCompanyPage = new SelectCompanyPage();
            Table table = new Table(selectCompanyPage.GetCompanyListTable());
            table.ClickCellValue("Company", "RTA", "Company");
            selectCompanyPage.ClickOkButton();

            homePage.ClickGeneralLedgerTab();
            homePage.ClickGeneralJournalLink();


            GeneralJournalPage generalJournalPage = new GeneralJournalPage();
            generalJournalPage.ClickNewMenuItem();
            generalJournalPage.SetNameValue("GL");
            generalJournalPage.SetDescriptionValue("General Journal RTA");

            generalJournalPage.ClickLinesMenuItem();
            JournalVoucherPage journalVoucherPage = new JournalVoucherPage();
            table = new Table(journalVoucherPage.GetJournalValueTable());
            table.ClickCellValue("Account type", "Ledger", "Account");

            journalVoucherPage.SetAccountSeg1Value("11300");

            InfoLogPage infoLogPage = new InfoLogPage();
            infoLogPage.GetTreeItemName("Value 11300 is not allowed for manual entry", "Message");
            infoLogPage.ClickCloseButton();

            journalVoucherPage.SetAccountSeg1Value("71000");

            infoLogPage = new InfoLogPage();
            infoLogPage.GetTreeItemName("Value 71000 is not allowed for manual entry", "Message");
            infoLogPage.ClickCloseButton();


            journalVoucherPage.SetAccountSeg1Value("11520");
            table.ClickCellValue("Account type", "Ledger", "Offset account");
            journalVoucherPage.SetAccountSeg1Value("11540");
            table.ClickCellValue("Account type", "Ledger", "Debit");
            journalVoucherPage.SetDebitValue("1000");


            journalVoucherPage.ClickApprovalItemMenuButton();
            journalVoucherPage.ClickReportAsReadyItemMenuItem();
            journalVoucherPage.ClickApprovalItemMenuButton();
            journalVoucherPage.ClickApproveItemMenuItem();

            journalVoucherPage.ClickPostMenuButton();
            journalVoucherPage.ClickPostItemMenuItem();

            infoLogPage = new InfoLogPage();
            infoLogPage.GetTreeItemName("Number of vouchers posted to the journal: 1", "Message");
            infoLogPage.ClickCloseButton();
            journalVoucherPage.ClickCloseButton();
            generalJournalPage.ClickCloseButton();
            

        }

      
    }

   
}
