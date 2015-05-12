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
    /// <summary>
    /// Summary description for CodedUITest1
    /// </summary>
    [CodedUITest]
    public class AXEFTFileProcessing : TestBase
    {

      
        
        public AXEFTFileProcessing()
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
        [TestProperty("TestcaseID", "7081")]
        public void ATC7081b_AXEFTreceiptforSingleBond()
        {

            #region Start Up Excel
            MyBook = MyApp.Workbooks.Open(DatasourceDir + @"\TenancyRequests.xlsx", 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            MySheet = (Excel.Worksheet)MyBook.Sheets[Properties.Settings.Default.ENVIRONMENT.ToString()];
            MyRange = MySheet.UsedRange;

            //Get specific row for the data
            int testDataRows = MyRange.Rows.Count;
            int MyRow = 0;
            for (int i = 2; i <= testDataRows; i++)
            {
                if (MyRange.Cells[i, 1].Value.ToString() == "7081")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string filelocation = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("OUTFILE")].Value.ToString();
            string paymentreference = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAY_REF_NUMBER")].Value.ToString();
            string tenancyRequestReference = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value.ToString();
            string managingParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();
            string contributor = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MISC1")].Value.ToString();
            string initialContribution = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString();
            string combinedContribution = ((double)Int32.Parse((MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value * 2).ToString())).ToString("C");
            string initialContribAmount = ((double)Int32.Parse(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString())).ToString("C");
            string amountPaidLodgement = ((double)Int32.Parse(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString())).ToString("C");


            Homepage homePage = new Homepage();
            homePage.ClickCompanyButton();

            SelectCompanyPage selectCompanyPage = new SelectCompanyPage();
            Table table = new Table(selectCompanyPage.GetCompanyListTable());
            table.ClickCellValue("Company", "RTB", "Company");
            selectCompanyPage.ClickOkButton();

            homePage.ClickHomeTab();
            homePage.ClickCashandBankManagementTab();
            homePage.ClickBankStatementsLink();

            homePage.ClickImportStatementButton();
            BAI2FileImportPage bai2FileImportPage = new BAI2FileImportPage();
            bai2FileImportPage.SetStatmentFormatText("CBA BAI2");
            bai2FileImportPage.SetImportFileFolderCheckBox(false);
            bai2FileImportPage.SetFileNamelocationText(filelocation);
            bai2FileImportPage.SetReconcileImportCheckBox(true);
            bai2FileImportPage.ClickOKButton();

            InfoLogPage infoLogPage = new InfoLogPage();
            string bvdTreeItem = infoLogPage.GetTreeItemName("Journal BDV", "Posting");
            string BVD = bvdTreeItem.Substring(8, 9);
            //Assert.IsTrue(infoLogPage.GetTreeItemExists("1 files have been imported in total.", "Posting"));
            infoLogPage.ClickClearButton();
            infoLogPage.ClickCloseButton();

            homePage.ClickBackNavButton();
            homePage.ClickHomeTab();
            homePage.ClickBondManagementTab();

            homePage.ClickCRMOutboundNotificationsLink();
            OutboundCRMIntegrationPage crmOutboundIntegrationPage = new OutboundCRMIntegrationPage();
            crmOutboundIntegrationPage.ClickOKButton();

            homePage.ClickReceiptJournalsLink();
            BondReceiptJournalPage bondReceiptJournalPage = new BondReceiptJournalPage();
            bondReceiptJournalPage.SetShowAllText("All");
            Keyboard.SendKeys("{ENTER}");
            
            table = new Table(bondReceiptJournalPage.GetBondReceiptTable());
            table.FilterCellValue("Bond journal");
            bondReceiptJournalPage.ClickFilterMenuItem();
            FilterPage filterPage = new FilterPage();
            filterPage.SetFilterText("Bond journal",BVD);
            filterPage.ClickOkButton();



           // StringAssert.Contains(table.GetCellValue("Bond journal", BVD, "Description"), RTB);
            table.ClickCellValue("Bond journal", BVD, "Bond journal");

            bondReceiptJournalPage.ClickLinesMenuItem();

            BondReceiptJournalLinesPage bondReceiptJournalLines = new BondReceiptJournalLinesPage();
            bondReceiptJournalLines.ClickInquiryButton();

            bondReceiptJournalLines.ClickBondTransactionsMenuItem();

            BondTransactionsPage bondTransactionsPage = new BondTransactionsPage();
            table = new Table(bondTransactionsPage.GetBondTransactionTable());

            StringAssert.Contains(table.GetCellContainsValue("Bond client", managingParty, "Bond request"), tenancyRequestReference);
            StringAssert.Contains(initialContribAmount, table.GetCellContainsValue("Bond client", managingParty, "Amount"));
            //StringAssert.Contains(table.GetCellContainsValue("Bond client", "BLAIR TEST", "Bond client"), "C23");
            StringAssert.Contains(table.GetCellContainsValue("Bond client", contributor, "Bond request"), tenancyRequestReference);
            StringAssert.Contains(initialContribAmount, table.GetCellContainsValue("Bond client", contributor, "Amount"));
//            StringAssert.Contains(table.GetCellContainsValue("Bond client", "BRIAN JOHN GILLAN", "Bond client"), "C31");
            StringAssert.Contains(table.GetCellValue("Amount", combinedContribution.Substring(1), "Bond client"), "Bond Receipts Client Account");

            bondTransactionsPage.ClickCloseButton();
            bondReceiptJournalLines.ClickCloseButton();
            bondReceiptJournalPage.ClickCloseButton();

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

        [TestMethod]
        [TestProperty("TestcaseID", "TopUp_TestData")]
        public void ATC_TopUpTestDatab()
        {

            #region Start Up Excel
            MyBook = MyApp.Workbooks.Open(DatasourceDir + @"\TenancyRequests.xlsx", 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            MySheet = (Excel.Worksheet)MyBook.Sheets[Properties.Settings.Default.ENVIRONMENT.ToString()];
            MyRange = MySheet.UsedRange;

            //Get specific row for the data
            int testDataRows = MyRange.Rows.Count;
            int MyRow = 0;
            for (int i = 2; i <= testDataRows; i++)
            {
                if (MyRange.Cells[i, 1].Value.ToString() == "TopUp_TestData")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string filelocation = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("OUTFILE")].Value.ToString();
            string paymentreference = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAY_REF_NUMBER")].Value.ToString();
            string tenancyRequestReference = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value.ToString();
            string managingParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();
            string initialContribution = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString();


            Homepage homePage = new Homepage();
            homePage.ClickCompanyButton();

            SelectCompanyPage selectCompanyPage = new SelectCompanyPage();
            Table table = new Table(selectCompanyPage.GetCompanyListTable());
            table.ClickCellValue("Company", "RTB", "Company");
            selectCompanyPage.ClickOkButton();

            homePage.ClickHomeTab();
            homePage.ClickCashandBankManagementTab();
            homePage.ClickBankStatementsLink();

            homePage.ClickImportStatementButton();
            BAI2FileImportPage bai2FileImportPage = new BAI2FileImportPage();
            bai2FileImportPage.SetStatmentFormatText("CBA BAI2");
            bai2FileImportPage.SetImportFileFolderCheckBox(false);
            bai2FileImportPage.SetFileNamelocationText(filelocation);
            bai2FileImportPage.SetReconcileImportCheckBox(true);
            bai2FileImportPage.ClickOKButton();

            InfoLogPage infoLogPage = new InfoLogPage();
            string bvdTreeItem = infoLogPage.GetTreeItemName("Journal BDV","Posting");
            string BVD = bvdTreeItem.Substring(8, 9);

            Assert.IsTrue(infoLogPage.GetTreeItemExists("1 files have been imported in total.", "Posting"));
            infoLogPage.ClickClearButton();
            infoLogPage.ClickCloseButton();

            homePage = new Homepage();
            homePage.ClickHomeTab();
            homePage.ClickBondManagementTab();

            //CRM Outbound Messages
            homePage.ClickSystemLink();
            homePage.ClickOutboundCRMMessagesLink();
            OutboundCRMMessagePage outboundCRMMessagePage = new OutboundCRMMessagePage();
            table = new Table(outboundCRMMessagePage.GetNotificationTable());
            table.FilterCellValue("Payment reference");
            outboundCRMMessagePage.ClickFilterMenuItem();
            FilterPage filterPage = new FilterPage();
            filterPage.SetFilterText("Payment reference", paymentreference);
            filterPage.ClickOkButton();
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("BOND_REF")].Value = table.GetCellValue("Payment reference", paymentreference, "Bond");
            

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

        [TestMethod]
        [TestProperty("TestcaseID", "8589")]
        public void ATC85891b_AX7098EFTReceiptRecordEFTreceiptfrombankstatementitem()
        {

            #region Start Up Excel
            MyBook = MyApp.Workbooks.Open(DatasourceDir + @"\TenancyRequests.xlsx", 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            MySheet = (Excel.Worksheet)MyBook.Sheets[Properties.Settings.Default.ENVIRONMENT.ToString()];
            MyRange = MySheet.UsedRange;

            //Get specific row for the data
            int testDataRows = MyRange.Rows.Count;
            int MyRow = 0;
            for (int i = 2; i <= testDataRows; i++)
            {
                if (MyRange.Cells[i, 1].Value == 85891)
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string filelocation = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("OUTFILE")].Value.ToString();
            string paymentreference = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAY_REF_NUMBER")].Value.ToString();
            string tenancyRequestReference = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value.ToString();
            string managingParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();
            string initialContribution = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString();

            Homepage homePage = new Homepage();
            homePage.ClickCompanyButton();

            SelectCompanyPage selectCompanyPage = new SelectCompanyPage();
            Table table = new Table(selectCompanyPage.GetCompanyListTable());
            table.ClickCellValue("Company", "RTB", "Company");
            selectCompanyPage.ClickOkButton();

            homePage.ClickHomeTab();
            homePage.ClickCashandBankManagementTab();
            homePage.ClickBankStatementsLink();

            homePage.ClickImportStatementButton();
            BAI2FileImportPage bai2FileImportPage = new BAI2FileImportPage();
            bai2FileImportPage.SetStatmentFormatText("CBA BAI2");
            bai2FileImportPage.SetImportFileFolderCheckBox(false);
            bai2FileImportPage.SetFileNamelocationText(filelocation);
            bai2FileImportPage.SetReconcileImportCheckBox(true);
            bai2FileImportPage.ClickOKButton();

            InfoLogPage infoLogPage = new InfoLogPage();
            string bvdTreeItem = infoLogPage.GetTreeItemName("Journal BDV", "Posting");
            string BVD = bvdTreeItem.Substring(8, 9);

            Assert.IsTrue(infoLogPage.GetTreeItemExists("1 files have been imported in total.", "Posting"));
            infoLogPage.ClickClearButton();
            infoLogPage.ClickCloseButton();


            homePage.ClickHomeTab();
            homePage.ClickBondManagementTab();

            homePage.ClickReceiptJournalsLink();


            BondReceiptJournalPage bondReceiptJournalPage = new BondReceiptJournalPage();
            bondReceiptJournalPage.SetShowAllText("All");
            Keyboard.SendKeys("{ENTER}");

            table = new Table(bondReceiptJournalPage.GetBondReceiptTable());
            table.FilterCellValue("Bond journal");
            bondReceiptJournalPage.ClickFilterMenuItem();
            FilterPage filterPage = new FilterPage();
            filterPage.SetFilterText("Bond journal", BVD);
            filterPage.ClickOkButton();

            //StringAssert.Contains(table.GetCellValue("Bond journal", BVD, "Description"), RTB);
            table.ClickCellValue("Bond journal", BVD, "Bond journal");

            bondReceiptJournalPage.ClickLinesMenuItem();

            BondReceiptJournalLinesPage bondReceiptJournalLines = new BondReceiptJournalLinesPage();
            table = new Table(bondReceiptJournalLines.GetBondReceiptJournalLinesTable());
            StringAssert.Contains(table.GetCellValue("Payment reference",paymentreference, "Allocated amount"), 
                MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString());
  
            
            bondReceiptJournalLines.ClickInquiryButton();
            bondReceiptJournalLines.ClickBondTransactionsMenuItem();

            BondTransactionsPage bondTransactionsPage = new BondTransactionsPage();
            table = new Table(bondTransactionsPage.GetBondTransactionTable());

            StringAssert.Contains(table.GetCellContainsValue("Bond client", managingParty, "Bond request"), tenancyRequestReference);
            StringAssert.Contains(table.GetCellContainsValue("Bond client", managingParty, "Amount"), initialContribution);
            StringAssert.Contains(table.GetCellValue("Type", "Receipt", "Amount"), initialContribution);
            StringAssert.Contains(table.GetCellValue("Type", "Receipt allocation", "Amount"), "-"+initialContribution);


            bondTransactionsPage.ClickCloseButton();
            bondReceiptJournalLines.ClickCloseButton();
            bondReceiptJournalPage.ClickCloseButton();


            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion




        }

        [TestMethod]
        [TestProperty("TestcaseID", "8589")]
        public void ATC85892d_AX7098EFTReceiptRecordEFTreceiptfrombankstatementitem()
        {

            #region Start Up Excel
            MyBook = MyApp.Workbooks.Open(DatasourceDir + @"\TenancyRequests.xlsx", 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            MySheet = (Excel.Worksheet)MyBook.Sheets[Properties.Settings.Default.ENVIRONMENT.ToString()];
            MyRange = MySheet.UsedRange;

            //Get specific row for the data
            int testDataRows = MyRange.Rows.Count;
            int MyRow = 0;
            for (int i = 2; i <= testDataRows; i++)
            {
                if (MyRange.Cells[i, 1].Value == 858923)
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string filelocation = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("OUTFILE")].Value.ToString();
            string paymentreference = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAY_REF_NUMBER")].Value.ToString();
            string batchRequestReference = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value.ToString();
            string managingParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();
            string initialContribution = Convert.ToDouble(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value).ToString("N2");

            Homepage homePage = new Homepage();
            homePage.ClickCompanyButton();

            SelectCompanyPage selectCompanyPage = new SelectCompanyPage();
            Table table = new Table(selectCompanyPage.GetCompanyListTable());
            table.ClickCellValue("Company", "RTB", "Company");
            selectCompanyPage.ClickOkButton();

            homePage.ClickHomeTab();
            homePage.ClickCashandBankManagementTab();
            homePage.ClickBankStatementsLink();

            homePage.ClickImportStatementButton();
            BAI2FileImportPage bai2FileImportPage = new BAI2FileImportPage();
            bai2FileImportPage.SetStatmentFormatText("CBA BAI2");
            bai2FileImportPage.SetImportFileFolderCheckBox(false);
            bai2FileImportPage.SetFileNamelocationText(filelocation);
            bai2FileImportPage.SetReconcileImportCheckBox(true);
            bai2FileImportPage.ClickOKButton();

            InfoLogPage infoLogPage = new InfoLogPage();
            string bvdTreeItem = infoLogPage.GetTreeItemName("Journal BDV", "Posting");
            string BVD = bvdTreeItem.Substring(8, 9);

            Assert.IsTrue(infoLogPage.GetTreeItemExists("1 files have been imported in total.", "Posting"));
            infoLogPage.ClickClearButton();
            infoLogPage.ClickCloseButton();


            homePage.ClickHomeTab();
            homePage.ClickBondManagementTab();

            homePage.ClickReceiptJournalsLink();


            BondReceiptJournalPage bondReceiptJournalPage = new BondReceiptJournalPage();
            bondReceiptJournalPage.SetShowAllText("All");
            Keyboard.SendKeys("{ENTER}");

            table = new Table(bondReceiptJournalPage.GetBondReceiptTable());
            table.FilterCellValue("Bond journal");
            bondReceiptJournalPage.ClickFilterMenuItem();
            FilterPage filterPage = new FilterPage();
            filterPage.SetFilterText("Bond journal", BVD);
            filterPage.ClickOkButton();

            //StringAssert.Contains(table.GetCellValue("Bond journal", BVD, "Description"), RTB);
            table.ClickCellValue("Bond journal", BVD, "Bond journal");

            bondReceiptJournalPage.ClickLinesMenuItem();

            BondReceiptJournalLinesPage bondReceiptJournalLines = new BondReceiptJournalLinesPage();
            table = new Table(bondReceiptJournalLines.GetBondReceiptJournalLinesTable());
            StringAssert.Contains(table.GetCellValue("Payment reference", paymentreference, "Allocated amount"), initialContribution);


            bondReceiptJournalLines.ClickInquiryButton();
            bondReceiptJournalLines.ClickBondTransactionsMenuItem();

            BondTransactionsPage bondTransactionsPage = new BondTransactionsPage();
            table = new Table(bondTransactionsPage.GetBondTransactionTable());

            //Add tenancy requests
            //Get specific row for the data
            int TR1Row = 0;
            for (int i = 2; i <= testDataRows; i++)
            {
                if (MyRange.Cells[i, 1].Value == 858921)
                {
                    TR1Row = i;
                    break;
                }
            }
            string TR1reference = MyRange.Cells[TR1Row, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value.ToString();
            string TR1managingParty = MyRange.Cells[TR1Row, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();
            string TR1initialContribution = Convert.ToDouble(MyRange.Cells[TR1Row, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value).ToString("N2");
            //Get specific row for the data
            int TR2Row = 0;
            for (int i = 2; i <= testDataRows; i++)
            {
                if (MyRange.Cells[i, 1].Value == 858922)
                {
                    TR2Row = i;
                    break;
                }
            }
            string TR2reference = MyRange.Cells[TR2Row, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value.ToString();
            string TR2managingParty = MyRange.Cells[TR2Row, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();
            string TR2initialContribution = Convert.ToDouble(MyRange.Cells[TR2Row, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value).ToString("N2");


            StringAssert.Contains(table.GetCellContainsValue("Bond client", TR1managingParty, "Bond request"), TR1reference);
            StringAssert.Contains(table.GetCellContainsValue("Bond client", TR1managingParty, "Amount"), TR1initialContribution);
            StringAssert.Contains(table.GetCellContainsValue("Bond client", TR2managingParty, "Bond request"), TR2reference);
            StringAssert.Contains(table.GetCellContainsValue("Bond client", TR2managingParty, "Amount"), TR2initialContribution);

            StringAssert.Contains(table.GetCellValue("Type", "Receipt", "Amount"), initialContribution);
            

            bondTransactionsPage.ClickCloseButton();
            bondReceiptJournalLines.ClickCloseButton();
            bondReceiptJournalPage.ClickCloseButton();


            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion




        }

        [TestMethod]
        [TestProperty("TestcaseID", "8589")]
        public void ATC85893_AX7098EFTReceiptRecordEFTreceiptfrombankstatementitem()
        {


            //Create Bank Statement file with invalid reference
            string filelocation = Utils.BAI2FileCreator.bAI2InvalidRefFileCreator();
            string paymentreference = filelocation.Substring(filelocation.Length - 12, 8);
            string initialContribution = "500.00";

            Homepage homePage = new Homepage();
            homePage.ClickCompanyButton();

            SelectCompanyPage selectCompanyPage = new SelectCompanyPage();
            Table table = new Table(selectCompanyPage.GetCompanyListTable());
            table.ClickCellValue("Company", "RTB", "Company");
            selectCompanyPage.ClickOkButton();

            homePage.ClickHomeTab();
            homePage.ClickCashandBankManagementTab();
            homePage.ClickBankStatementsLink();

            homePage.ClickImportStatementButton();
            BAI2FileImportPage bai2FileImportPage = new BAI2FileImportPage();
            bai2FileImportPage.SetStatmentFormatText("CBA BAI2");
            bai2FileImportPage.SetImportFileFolderCheckBox(false);
            bai2FileImportPage.SetFileNamelocationText(filelocation);
            bai2FileImportPage.SetReconcileImportCheckBox(true);
            bai2FileImportPage.ClickOKButton();

            InfoLogPage infoLogPage = new InfoLogPage();
            string bvdTreeItem = infoLogPage.GetTreeItemName("Journal BDV", "Posting");
            string BVD = bvdTreeItem.Substring(8, 9);

            Assert.IsTrue(infoLogPage.GetTreeItemExists("1 files have been imported in total.", "Posting"));
            infoLogPage.ClickClearButton();
            infoLogPage.ClickCloseButton();

            homePage.ClickBackNavButton();
            homePage.ClickHomeTab();
            homePage.ClickBondManagementTab();

            homePage.ClickReceiptJournalsLink();


            BondReceiptJournalPage bondReceiptJournalPage = new BondReceiptJournalPage();
            bondReceiptJournalPage.SetShowAllText("All");
           

            Keyboard.SendKeys("{ENTER}");

            table = new Table(bondReceiptJournalPage.GetBondReceiptTable());
            table.FilterCellValue("Bond journal");
            bondReceiptJournalPage.ClickFilterMenuItem();
            FilterPage filterPage = new FilterPage();
            filterPage.SetFilterText("Bond journal",BVD);
            filterPage.ClickOkButton();

            //StringAssert.Contains(table.GetCellValue("Bond journal", BVD, "Description"), RTB);
            table.ClickCellValue("Bond journal", BVD, "Bond journal");

            bondReceiptJournalPage.ClickLinesMenuItem();

            BondReceiptJournalLinesPage bondReceiptJournalLines = new BondReceiptJournalLinesPage();
            table = new Table(bondReceiptJournalLines.GetBondReceiptJournalLinesTable());
            StringAssert.Contains(table.GetCellValue("Payment reference", paymentreference, "Allocated amount"), "0.00");
            StringAssert.Contains(table.GetCellValue("Payment reference", paymentreference, "Amount"), initialContribution);


            bondReceiptJournalLines.ClickInquiryButton();
            bondReceiptJournalLines.ClickBondTransactionsMenuItem();

            BondTransactionsPage bondTransactionsPage = new BondTransactionsPage();
            table = new Table(bondTransactionsPage.GetBondTransactionTable());


            StringAssert.Contains(table.GetCellContainsValue("Bond client", "[UNKNOWN]", "Amount"), initialContribution);
            StringAssert.Contains(table.GetCellValue("Type", "Receipt", "Amount"), initialContribution);
            StringAssert.Contains(table.GetCellValue("Type", "Receipt allocation", "Amount"), "-" + initialContribution);


            bondTransactionsPage.ClickCloseButton();
            bondReceiptJournalLines.ClickCloseButton();
            bondReceiptJournalPage.ClickCloseButton();


        }

        [TestMethod]
        [TestProperty("TestcaseID", "4477")]
        public void ATC4477_AX2849DuplicateControlandArchivingBankStatementfile()
        {


            //Create Bank Statement file with invalid reference
            string filelocation;
            Random random = new Random();
            int randomNum;
            randomNum = random.Next(30000000, 39999999);
            string dateValue;
            

            dateValue = DateTime.Today.ToString("yyMMdd");
            string timeValue;
            timeValue = DateTime.Now.ToString("HHmm");
            filelocation = Utils.BAI2FileCreator.bAI2UnknownCreator(randomNum, "100", dateValue, timeValue);
            string fileName;
            fileName = filelocation.Substring(filelocation.Length - 35);

            Homepage homePage = new Homepage();
            homePage.ClickCompanyButton();

            SelectCompanyPage selectCompanyPage = new SelectCompanyPage();
            Table table = new Table(selectCompanyPage.GetCompanyListTable());
            table.ClickCellValue("Company", "RTB", "Company");
            selectCompanyPage.ClickOkButton();

            homePage.ClickHomeTab();
            homePage.ClickCashandBankManagementTab();
            homePage.ClickBankStatementsLink();

            homePage.ClickImportStatementButton();
            BAI2FileImportPage bai2FileImportPage = new BAI2FileImportPage();
            bai2FileImportPage.SetStatmentFormatText("CBA BAI2");
            bai2FileImportPage.SetImportFileFolderCheckBox(false);
            bai2FileImportPage.SetFileNamelocationText(filelocation);
            bai2FileImportPage.SetReconcileImportCheckBox(true);
            bai2FileImportPage.ClickOKButton();

            InfoLogPage infoLogPage = new InfoLogPage();
            string bvdTreeItem = infoLogPage.GetTreeItemName("Journal BDV", "Posting");
            string BVD = bvdTreeItem.Substring(8, 9);

            Assert.IsTrue(infoLogPage.GetTreeItemExists("1 files have been imported in total.", "Posting"));
            infoLogPage.ClickClearButton();
            infoLogPage.ClickCloseButton();

            filelocation = Utils.BAI2FileCreator.bAI2UnknownCreator(randomNum, "100", dateValue, timeValue);

            homePage.ClickImportStatementButton();
            bai2FileImportPage = new BAI2FileImportPage();
            bai2FileImportPage.SetStatmentFormatText("CBA BAI2");
            bai2FileImportPage.SetImportFileFolderCheckBox(false);
            bai2FileImportPage.SetFileNamelocationText(filelocation);
            bai2FileImportPage.SetReconcileImportCheckBox(true);
            bai2FileImportPage.ClickOKButton();

            //Need to get message of duplicate error
            infoLogPage = new InfoLogPage();
            string DuplicateMessageTreeItem;
            DuplicateMessageTreeItem = infoLogPage.GetTreeItemName("Duplicate file", "Message (");
            string ErrorMessageTreeItem;
            ErrorMessageTreeItem = infoLogPage.GetTreeItemName("Error importing bank statement file ", "Message (");
            StringAssert.Contains(ErrorMessageTreeItem, fileName);
            infoLogPage.ClickClearButton();
            infoLogPage.ClickCloseButton();

            homePage.ClickBackNavButton();
            homePage.ClickHomeTab();
            homePage.ClickCashandBankManagementTab();
            homePage.ClickBankStatementFileImportExceptionLink();

            BSFileImportExceptionPage bsFileImportExceptionPage = new BSFileImportExceptionPage();
            table = new Table(bsFileImportExceptionPage.GetBSFileTable());
            table.FilterCellValue("File name");
            bsFileImportExceptionPage.ClickFilterMenuItem();
            FilterPage filterPage = new FilterPage();
            filterPage.SetFilterText("File name", "*" + fileName);
            filterPage.ClickOkButton();
            StringAssert.Contains(filelocation, table.GetCellContainsValue("File name", fileName, "File name"));
            bsFileImportExceptionPage.ClickShowLogButton();

            //Need to get message of duplicate error
            infoLogPage = new InfoLogPage();
            DuplicateMessageTreeItem = infoLogPage.GetTreeItemName("Duplicate file", "Message (");
            ErrorMessageTreeItem = infoLogPage.GetTreeItemName("Error importing bank statement file ", "Message (");
            StringAssert.Contains(ErrorMessageTreeItem, fileName);
            infoLogPage.ClickClearButton();
            infoLogPage.ClickCloseButton();
            bsFileImportExceptionPage.ClickCloseButton();

        }


        [TestMethod]
        [TestProperty("TestcaseID", "6911b")]
        public void ATC6911b_TenancytopupvalidationSameaddresssamemanagingparty()
        {

            #region Start Up Excel
            MyBook = MyApp.Workbooks.Open(DatasourceDir + @"\TenancyRequests.xlsx", 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            MySheet = (Excel.Worksheet)MyBook.Sheets[Properties.Settings.Default.ENVIRONMENT.ToString()];
            MyRange = MySheet.UsedRange;

            //Get specific row for the data
            int testDataRows = MyRange.Rows.Count;
            int MyRow = 0;
            for (int i = 2; i <= testDataRows; i++)
            {
                if (MyRange.Cells[i, 1].Value.ToString() == "6911")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string filelocation = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("OUTFILE")].Value.ToString();
            string paymentreference = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAY_REF_NUMBER")].Value.ToString();
            string tenancyRequestReference = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value.ToString();
            string managingParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();
            string initialContribution = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString();


            Homepage homePage = new Homepage();
            homePage.ClickCompanyButton();

            SelectCompanyPage selectCompanyPage = new SelectCompanyPage();
            Table table = new Table(selectCompanyPage.GetCompanyListTable());
            table.ClickCellValue("Company", "RTB", "Company");
            selectCompanyPage.ClickOkButton();

            homePage.ClickHomeTab();
            homePage.ClickCashandBankManagementTab();
            homePage.ClickBankStatementsLink();

            homePage.ClickImportStatementButton();
            BAI2FileImportPage bai2FileImportPage = new BAI2FileImportPage();
            bai2FileImportPage.SetStatmentFormatText("CBA BAI2");
            bai2FileImportPage.SetImportFileFolderCheckBox(false);
            bai2FileImportPage.SetFileNamelocationText(filelocation);
            bai2FileImportPage.SetReconcileImportCheckBox(true);
            bai2FileImportPage.ClickOKButton();

            InfoLogPage infoLogPage = new InfoLogPage();
            string bvdTreeItem = infoLogPage.GetTreeItemName("Journal BDV", "Posting");
            string BVD = bvdTreeItem.Substring(8, 9);

            Assert.IsTrue(infoLogPage.GetTreeItemExists("1 files have been imported in total.", "Posting"));
            infoLogPage.ClickClearButton();
            infoLogPage.ClickCloseButton();

            homePage.ClickBackNavButton();

            homePage.ClickHomeTab();
            homePage.ClickBondManagementTab();

            homePage.ClickCRMOutboundNotificationsLink();
            OutboundCRMIntegrationPage crmOutboundIntegrationPage = new OutboundCRMIntegrationPage();
            crmOutboundIntegrationPage.ClickOKButton();

            //CRM Outbound Messages
            homePage.ClickSystemLink();
            homePage.ClickOutboundCRMMessagesLink();
            OutboundCRMMessagePage outboundCRMMessagePage = new OutboundCRMMessagePage();
            table = new Table(outboundCRMMessagePage.GetNotificationTable());
            table.FilterCellValue("Payment reference");
            outboundCRMMessagePage.ClickFilterMenuItem();
            FilterPage filterPage = new FilterPage();
            filterPage.SetFilterText("Payment reference", paymentreference);
            filterPage.ClickOkButton();
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("BOND_REF")].Value = table.GetCellValue("Payment reference", paymentreference, "Bond");


            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }

        [TestMethod]
        [TestProperty("TestcaseID", "6745")]
        public void ATC6745b_AXEFTFeedbacktoCRMMultipleBonds()
        {

            #region Start Up Excel
            MyBook = MyApp.Workbooks.Open(DatasourceDir + @"\TenancyRequests.xlsx", 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            MySheet = (Excel.Worksheet)MyBook.Sheets[Properties.Settings.Default.ENVIRONMENT.ToString()];
            MyRange = MySheet.UsedRange;

            //Get specific row for the data
            int testDataRows = MyRange.Rows.Count;
            int MyRow = 0;
            for (int i = 2; i <= testDataRows; i++)
            {
                if (MyRange.Cells[i, 1].Value.ToString() == "6745")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string filelocation = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("OUTFILE")].Value.ToString();
            string paymentreference = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAY_REF_NUMBER")].Value.ToString();
            string tenancyRequestReference = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value.ToString();
            string managingParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();
            string contributor = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MISC1")].Value.ToString();
            string initialContribution = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString();
            string combinedContribution = ((double)Int32.Parse((MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value * 2).ToString())).ToString("C");
            string initialContribAmount = ((double)Int32.Parse(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString())).ToString("C");
            string amountPaidLodgement = ((double)Int32.Parse(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString())).ToString("C");
            string amountOtherParty = (MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value +
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MISC2")].Value).ToString();

            Homepage homePage = new Homepage();
            homePage.ClickCompanyButton();

            SelectCompanyPage selectCompanyPage = new SelectCompanyPage();
            Table table = new Table(selectCompanyPage.GetCompanyListTable());
            table.ClickCellValue("Company", "RTB", "Company");
            selectCompanyPage.ClickOkButton();

            homePage.ClickHomeTab();
            homePage.ClickCashandBankManagementTab();
            homePage.ClickBankStatementsLink();

            homePage.ClickImportStatementButton();
            BAI2FileImportPage bai2FileImportPage = new BAI2FileImportPage();
            bai2FileImportPage.SetStatmentFormatText("CBA BAI2");
            bai2FileImportPage.SetImportFileFolderCheckBox(false);
            bai2FileImportPage.SetFileNamelocationText(filelocation);
            bai2FileImportPage.SetReconcileImportCheckBox(true);
            bai2FileImportPage.ClickOKButton();

            InfoLogPage infoLogPage = new InfoLogPage();
            string bvdTreeItem = infoLogPage.GetTreeItemName("Journal BDV", "Posting");
            string BVD = bvdTreeItem.Substring(8, 9);
            //Assert.IsTrue(infoLogPage.GetTreeItemExists("1 files have been imported in total.", "Posting"));
            infoLogPage.ClickClearButton();
            infoLogPage.ClickCloseButton();

            homePage.ClickBackNavButton();
            homePage.ClickHomeTab();
            homePage.ClickBondManagementTab();

            homePage.ClickCRMOutboundNotificationsLink();
            OutboundCRMIntegrationPage crmOutboundIntegrationPage = new OutboundCRMIntegrationPage();
            crmOutboundIntegrationPage.ClickOKButton();

            homePage.ClickSystemLink();
            homePage.ClickOutboundCRMMessagesLink();

            OutboundCRMMessagePage outboundCRMMessagePage = new OutboundCRMMessagePage();
            table = new Table(outboundCRMMessagePage.GetNotificationTable());
            table.FilterCellValue("Bond request");
            outboundCRMMessagePage.ClickFilterMenuItem();
            FilterPage filterPage = new FilterPage();
            filterPage.SetFilterText("Request document number", tenancyRequestReference);
            filterPage.ClickOkButton();

            StringAssert.Contains(table.GetCellValue("Bond request", tenancyRequestReference, "Payment reference"), paymentreference);
            StringAssert.Contains(table.GetCellValue("Bond request", tenancyRequestReference, "Bond Balance"),
                ((double)Int32.Parse(amountOtherParty)).ToString("C").Substring(1));

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
        }


        [TestMethod]
        [TestProperty("TestcaseID", "6746")]
        public void ATC6746b_AXEFTreceiptforSingleBond()
        {

            #region Start Up Excel
            MyBook = MyApp.Workbooks.Open(DatasourceDir + @"\TenancyRequests.xlsx", 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            MySheet = (Excel.Worksheet)MyBook.Sheets[Properties.Settings.Default.ENVIRONMENT.ToString()];
            MyRange = MySheet.UsedRange;

            //Get specific row for the data
            int testDataRows = MyRange.Rows.Count;
            int MyRow = 0;
            for (int i = 2; i <= testDataRows; i++)
            {
                if (MyRange.Cells[i, 1].Value.ToString() == "6746")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string filelocation = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("OUTFILE")].Value.ToString();
            string paymentreference = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAY_REF_NUMBER")].Value.ToString();
            string tenancyRequestReference = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value.ToString();
            string managingParty = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MANAGING_PARTY")].Value.ToString();
            string contributor = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MISC1")].Value.ToString();
            string initialContribution = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString();
            string combinedContribution = ((double)Int32.Parse((MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value * 2).ToString())).ToString("C");
            string initialContribAmount = ((double)Int32.Parse(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value.ToString())).ToString("C");
            string amountPaidLodgement = ((double)Int32.Parse(MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("AMOUNT_PAID_LODGEMENT")].Value.ToString())).ToString("C");
            string amountOtherParty = (MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("INITIAL_CONTRIBUTION")].Value +
               MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MISC2")].Value).ToString();

            Homepage homePage = new Homepage();
            homePage.ClickCompanyButton();

            SelectCompanyPage selectCompanyPage = new SelectCompanyPage();
            Table table = new Table(selectCompanyPage.GetCompanyListTable());
            table.ClickCellValue("Company", "RTB", "Company");
            selectCompanyPage.ClickOkButton();

            homePage.ClickHomeTab();
            homePage.ClickCashandBankManagementTab();
            homePage.ClickBankStatementsLink();

            homePage.ClickImportStatementButton();
            BAI2FileImportPage bai2FileImportPage = new BAI2FileImportPage();
            bai2FileImportPage.SetStatmentFormatText("CBA BAI2");
            bai2FileImportPage.SetImportFileFolderCheckBox(false);
            bai2FileImportPage.SetFileNamelocationText(filelocation);
            bai2FileImportPage.SetReconcileImportCheckBox(true);
            bai2FileImportPage.ClickOKButton();

            InfoLogPage infoLogPage = new InfoLogPage();
            string bvdTreeItem = infoLogPage.GetTreeItemName("Journal BDV", "Posting");
            string BVD = bvdTreeItem.Substring(8, 9);
            //Assert.IsTrue(infoLogPage.GetTreeItemExists("1 files have been imported in total.", "Posting"));
            infoLogPage.ClickClearButton();
            infoLogPage.ClickCloseButton();

            homePage.ClickBackNavButton();
            homePage.ClickHomeTab();
            homePage.ClickBondManagementTab();

            homePage.ClickCRMOutboundNotificationsLink();
            OutboundCRMIntegrationPage crmOutboundIntegrationPage = new OutboundCRMIntegrationPage();
            crmOutboundIntegrationPage.ClickOKButton();

            homePage.ClickSystemLink();
            homePage.ClickOutboundCRMMessagesLink();

            OutboundCRMMessagePage outboundCRMMessagePage = new OutboundCRMMessagePage();
            table = new Table(outboundCRMMessagePage.GetNotificationTable());
            table.FilterCellValue("Bond request");
            outboundCRMMessagePage.ClickFilterMenuItem();
            FilterPage filterPage = new FilterPage();
            filterPage.SetFilterText("Request document number", tenancyRequestReference);
            filterPage.ClickOkButton();

            StringAssert.Contains(table.GetCellValue("Bond request", tenancyRequestReference, "Payment reference"), paymentreference);
            StringAssert.Contains(table.GetCellValue("Bond request", tenancyRequestReference, "Bond Balance"),
                ((double)Int32.Parse(amountOtherParty)).ToString("C").Substring(1));

            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("BOND_REF")].Value = table.GetCellValue("Bond request", tenancyRequestReference, "Bond");

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion
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
