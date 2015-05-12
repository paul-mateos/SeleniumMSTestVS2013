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
    public class AXBPayFileProcessing : TestBase
    {

       
        
        public AXBPayFileProcessing()
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
        [TestProperty("TestcaseID", "7120")]
        public void ATC7120b_AXAC1Fileformatverification()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "7120")
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
            homePage.ClickHomeTab();
            homePage.ClickBondManagementTab();
            homePage.ClickPaymentsLink();
            homePage.ClickBPayFileLink();


            BPayFilePage bpayFile = new BPayFilePage();
            bpayFile.ClickImportMenuItem();

            BPayFileImportPage bpayFileImportPage = new BPayFileImportPage();
            Assert.IsTrue(bpayFileImportPage.GetWindowExistStatus());
            bpayFileImportPage.SetMoveFileCheckBox(true);



            bpayFileImportPage.SetProcessFileCheckBox(true);
            bpayFileImportPage.SetImportPathText("");
            bpayFileImportPage.SetFileNameEdit(filelocation);
            bpayFileImportPage.ClickOKButton();

            InfoLogPage infoLogPage = new InfoLogPage();
            //check that file imported successful
            //Assert.IsTrue(infoLogPage.GetTreeItemNOTExists("Cannot create a record in", "Importing BPAY file"));
            Assert.IsTrue(infoLogPage.GetTreeItemExists("Posting journal BDV", "Importing BPAY file"));
            infoLogPage.ClickClearButton();
            infoLogPage.ClickCloseButton();
            bpayFile.ClickCloseButton();

            homePage.ClickCRMOutboundNotificationsLink();
            OutboundCRMIntegrationPage outboundCRMIntegrationPage = new OutboundCRMIntegrationPage();
            outboundCRMIntegrationPage.ClickOKButton();

            //CRM Outbound Messages
            homePage.ClickSystemLink();
            homePage.ClickOutboundCRMMessagesLink();
            OutboundCRMMessagePage outboundCRMMessagePage = new OutboundCRMMessagePage();
            Table table = new Table(outboundCRMMessagePage.GetNotificationTable());
            table.FilterCellValue("Payment reference");
            outboundCRMMessagePage.ClickFilterMenuItem();
            FilterPage filterPage = new FilterPage();
            filterPage.SetFilterText("Payment reference", paymentreference);
            filterPage.ClickOkButton();

            StringAssert.Contains(table.GetCellValue("Payment reference", paymentreference, "Bond Balance"), initialContribution);

            ////Navigate to Payments/BPay file exceptions
            //homePage.ClickPaymentsLink();
            //homePage.ClickBPayFileExceptionLink();

            ////Order by created date

            ////Confirm customer referance number exists
            //BPayFileExceptionsPage bPayFileExceptionPage = new BPayFileExceptionsPage();

            //Table table = new Table(bPayFileExceptionPage.GetFileExceptionTable());

            //table.FilterCellValue("Customer reference number");

            //bPayFileExceptionPage.ClickFilterMenuItem();

            //FilterPage filterPage = new FilterPage();
            //filterPage.SetCustomerReferenceNumberText(paymentreference);
            //filterPage.ClickOkButton();

            //StringAssert.Contains(table.GetCellValue("Customer reference number", paymentreference, "Rejection reason"), "Invalid CRN");
            //bPayFileExceptionPage.ClickCloseButton();

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }


        [TestMethod]
        [TestProperty("TestcaseID", "7121")]
        public void ATC7121b_AXDuplicateBPAYpayment()
        {
            #region Start Up Excel
            //MyApp = new Excel.Application();
            //MyApp.Visible = false;
            //MyBook = MyApp.Workbooks.Open(DatasourceDir + "\\TenancyRequests.xlsx");
            MyBook = MyApp.Workbooks.Open(DatasourceDir + @"\TenancyRequests.xlsx", 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            MySheet = (Excel.Worksheet)MyBook.Sheets[Properties.Settings.Default.ENVIRONMENT.ToString()];
            MyRange = MySheet.UsedRange;

            //Get specific row for the data
            int testDataRows = MyRange.Rows.Count;
            int MyRow = 0;
            for (int i = 2; i <= testDataRows; i++)
            {
                if (MyRange.Cells[i, 1].Value.ToString() == "7121")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion

            string referenceNumber = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAY_REF_NUMBER")].Value.ToString();
            string tenancyrequest = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value.ToString();
            string fileLocation = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("OUTFILE")].Value.ToString();
            string fileLocation2 = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("MISC1")].Value.ToString();

            Homepage homePage = new Homepage();
            homePage.ClickHomeTab();
            homePage.ClickBondManagementTab();
            homePage.ClickPaymentsLink();
            homePage.ClickBPayFileLink(); ;

            BPayFilePage bPayFilePage = new BPayFilePage();

            //import 1st file
            bPayFilePage.ClickImportMenuItem();
            BPayFileImportPage bPayFileImportPage = new BPayFileImportPage();
            Assert.IsTrue(bPayFileImportPage.GetWindowExistStatus());
            bPayFileImportPage.SetMoveFileCheckBox(true);
            bPayFileImportPage.SetProcessFileCheckBox(true);
            bPayFileImportPage.SetImportPathEdit("");
            bPayFileImportPage.SetFileNameEdit(fileLocation);
            bPayFileImportPage.ClickOKButton();
            InfoLogPage infoLogPage = new InfoLogPage();
            //check that file imported successful
            Assert.IsTrue(infoLogPage.GetTreeItemExists("File imported with identifier", "Importing BPAY file"));
            infoLogPage.ClickClearButton();
            infoLogPage.ClickCloseButton();


            //import 2nd file
            bPayFilePage.ClickImportMenuItem();
            bPayFileImportPage = new BPayFileImportPage();
            Assert.IsTrue(bPayFileImportPage.GetWindowExistStatus());
            bPayFileImportPage.SetMoveFileCheckBox(true);
            bPayFileImportPage.SetProcessFileCheckBox(true);
            bPayFileImportPage.SetImportPathEdit("");
            bPayFileImportPage.SetFileNameEdit(fileLocation2);
            bPayFileImportPage.ClickOKButton();
            infoLogPage = new InfoLogPage();

            //check that file imported successful
            Assert.IsTrue(infoLogPage.GetTreeItemExists("File imported with identifier", "Importing BPAY file"));
            infoLogPage.ClickClearButton();
            infoLogPage.ClickCloseButton();

            bPayFilePage.ClickCloseButton();

            //Navigate to Outbound CRM and click ok
            homePage.ClickCRMOutboundNotificationsLink();
            OutboundCRMIntegrationPage outBoundCRMIntegration = new OutboundCRMIntegrationPage();
            Assert.IsTrue(outBoundCRMIntegration.GetWindowExistStatus());
            outBoundCRMIntegration.ClickOKButton();

            //Navigate to Payments/BPay file exceptions
            homePage.ClickPaymentsLink();
            homePage.ClickBPayFileExceptionLink();
            //Order by created date

            //Confirm customer referance number exists
            BPayFileExceptionsPage bPayFileExceptionPage = new BPayFileExceptionsPage();

            Table table = new Table(bPayFileExceptionPage.GetFileExceptionTable());
            table.FilterCellValue("Customer reference number");
            bPayFileExceptionPage.ClickFilterMenuItem();

            FilterPage filterPage = new FilterPage();
            filterPage.SetFilterText("Customer reference number", referenceNumber);
            filterPage.ClickOkButton();
            StringAssert.Contains(table.GetCellValue("Customer reference number", referenceNumber, "Rejection reason"), "Duplicate CRN");
            bPayFileExceptionPage.ClickCloseButton();

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }

        [TestMethod]
        [TestProperty("TestcaseID", "7124")]
        public void ATC7124_AXFinancialPostingsforBPAYpaymentsnumberGL()
        {

            
            string dateValue = DateTime.Today.AddDays(-1).ToString("yyyyMMdd");
            string transDesc = "BPAY " + DateTime.Today.AddDays(-1).ToString("d/MM/yyyy");
            Random random = new Random();
            int randomNum = random.Next(1000, 9999);
            string fileLocation = Utils.BPayFileCreator.bPayUnknownClientFileCreator(dateValue, randomNum);
            string referenceNumber = Utils.BPayFileReaderClass.GetPaymentReference1File(fileLocation);

            Homepage homePage = new Homepage();
            homePage.ClickCompanyButton();

            SelectCompanyPage selectCompanyPage = new SelectCompanyPage();
            Table table = new Table(selectCompanyPage.GetCompanyListTable());
            table.ClickCellValue("Company", "RTB", "Company");
            selectCompanyPage.ClickOkButton();

            homePage.ClickHomeTab();
            homePage.ClickBondManagementTab();
            


            homePage.ClickPaymentsLink();
            homePage.ClickBPayFileLink(); ;

            BPayFilePage bPayFilePage = new BPayFilePage();
            bPayFilePage.ClickImportMenuItem();

            BPayFileImportPage bPayFileImportPage = new BPayFileImportPage();
            Assert.IsTrue(bPayFileImportPage.GetWindowExistStatus());
            bPayFileImportPage.SetMoveFileCheckBox(true);
            bPayFileImportPage.SetProcessFileCheckBox(true);
            //bPayFileImportPage.SetImportPathEdit(@"P:\Dynamics AX\Bank files\Bpay\Paul");
            bPayFileImportPage.SetImportPathEdit("");
            bPayFileImportPage.SetFileNameEdit(fileLocation);
            bPayFileImportPage.ClickOKButton();

            InfoLogPage infoLogPage = new InfoLogPage();
            string rtbTreeItem = infoLogPage.GetTreeItemName("Processing BPAY file RTB-", "Importing BPAY file");
            string RTB = rtbTreeItem.Substring(21, 10);
            string bvdTreeItem = infoLogPage.GetTreeItemName("Posting journal BDV", "Importing BPAY file");
            string BVD = bvdTreeItem.Substring(16, 9);

            Assert.IsTrue(infoLogPage.GetTreeItemExists("1 file(s) were processed", "Importing BPAY file"));
            infoLogPage.ClickClearButton();
            infoLogPage.ClickCloseButton();
            bPayFilePage.ClickCloseButton();


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



            StringAssert.Contains(table.GetCellValue("Bond journal", BVD, "Description"), RTB);
            table.ClickCellValue("Bond journal", BVD, "Bond journal");

            bondReceiptJournalPage.ClickLinesMenuItem();

            BondReceiptJournalLinesPage bondReceiptJournalLines = new BondReceiptJournalLinesPage();
            bondReceiptJournalLines.ClickInquiryButton();

            bondReceiptJournalLines.ClickVoucherTransactionsMenuItem();

            VoucherTransactionsPage voucherTransactionsPage = new VoucherTransactionsPage();
            table = new Table(voucherTransactionsPage.GetVoucherTransactionTable());

            StringAssert.Contains(table.GetCellValue("Ledger account", "11330", "Amount"), "700.00");
            table.ClickCellValue("Ledger account", "11330", "Ledger account");
            StringAssert.Contains(voucherTransactionsPage.GetDescriptionText(), transDesc);
            StringAssert.Contains(voucherTransactionsPage.GetAccountNameText(), "Cash at Bank - Rental Bond BPAY Account");

            StringAssert.Contains(table.GetCellValue("Ledger account", "32130", "Amount"), "700.00");
            table.ClickCellValue("Ledger account", "32130", "Ledger account");
            StringAssert.Contains(voucherTransactionsPage.GetDescriptionText(), referenceNumber);
            StringAssert.Contains(voucherTransactionsPage.GetAccountNameText(), "Rental bonds - Receipts unallocated");

            voucherTransactionsPage.ClickCloseButton();
            bondReceiptJournalLines.ClickCloseButton();
            bondReceiptJournalPage.ClickCloseButton();


        }

        [TestMethod]
        [TestProperty("TestcaseID", "7125")]
        public void ATC7125b_AXTESTINGEndtoEndSingleFormBPayAXSuccess()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "7125")
                {
                    MyRow = i;
                    break;
                }
            }
            #endregion


            string referenceNumber = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("PAY_REF_NUMBER")].Value.ToString();
            string tenancyrequest = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("TR_NUMBER")].Value.ToString();
            string fileLocation = MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("OUTFILE")].Value.ToString();

            Homepage homePage = new Homepage();
            homePage.ClickHomeTab();
            homePage.ClickBondManagementTab();
            homePage.ClickPaymentsLink();
            homePage.ClickBPayFileLink();

            BPayFilePage bpayFile = new BPayFilePage();
            bpayFile.ClickImportMenuItem();

            BPayFileImportPage bpayFileImportPage = new BPayFileImportPage();
            Assert.IsTrue(bpayFileImportPage.GetWindowExistStatus());
            bpayFileImportPage.SetMoveFileCheckBox(true);

            bpayFileImportPage.SetProcessFileCheckBox(true);

            bpayFileImportPage.SetImportPathText("");
            bpayFileImportPage.SetFileNameEdit(fileLocation);
            bpayFileImportPage.ClickOKButton();

            InfoLogPage infoLogPage = new InfoLogPage();
            //check that file imported successful
            Assert.IsTrue(infoLogPage.GetTreeItemNOTExists("Cannot create a record in", "Importing BPAY file"));

            string rtbTreeItem = infoLogPage.GetTreeItemName("Processing BPAY file RTB-", "Importing BPAY file");
            string RTB = rtbTreeItem.Substring(21, 10);
            string bvdTreeItem = infoLogPage.GetTreeItemName("Posting journal BDV", "Importing BPAY file");
            string BVD = bvdTreeItem.Substring(16, 9);
            //Assert.IsTrue(infoLogPage.GetTreeItemExists("1 file(s) were processed","Importing BPAY file"));

            infoLogPage.ClickClearButton();
            infoLogPage.ClickCloseButton();
            bpayFile.ClickCloseButton();

            homePage.ClickCRMOutboundNotificationsLink();
            OutboundCRMIntegrationPage outboundCRMIntegrationPage = new OutboundCRMIntegrationPage();
            outboundCRMIntegrationPage.ClickOKButton();

            //CRM Outbound Messages
            homePage.ClickSystemLink();
            homePage.ClickOutboundCRMMessagesLink();
            OutboundCRMMessagePage outboundCRMMessagePage = new OutboundCRMMessagePage();
            Table table = new Table(outboundCRMMessagePage.GetNotificationTable());
            table.FilterCellValue("Payment reference");
            outboundCRMMessagePage.ClickFilterMenuItem();
            FilterPage filterPage = new FilterPage();
            filterPage.SetFilterText("Payment reference", referenceNumber);
            filterPage.ClickOkButton();

            StringAssert.Contains(table.GetCellValue("Payment reference", referenceNumber, "Bond request"), tenancyrequest);
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("BOND_REF")].Value = table.GetCellValue("Payment reference", referenceNumber, "Bond");

            outboundCRMMessagePage.ClickCloseButton();


            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion


        }


        [TestMethod]
        [TestProperty("TestcaseID", "4467")]
        public void ATC4467_AX1817DuplicateControlandArchivingBPayfile()
        {
            string dateValue = DateTime.Today.AddDays(-1).ToString("yyyyMMdd");
            string transDesc = "BPAY " + DateTime.Today.AddDays(-1).ToString("d/MM/yyyy");
            Random random = new Random();
            int randomNum = random.Next(1000, 9999);
            string fileLocation; 
            fileLocation = Utils.BPayFileCreator.bPayUnknownClientFileCreator(dateValue, randomNum);
            string fileName;
            fileName = fileLocation.Substring(fileLocation.Length - 33);
            string referenceNumber = Utils.BPayFileReaderClass.GetPaymentReference1File(fileLocation);

            Homepage homePage = new Homepage();
            homePage.ClickCompanyButton();

            SelectCompanyPage selectCompanyPage = new SelectCompanyPage();
            Table table = new Table(selectCompanyPage.GetCompanyListTable());
            table.ClickCellValue("Company", "RTB", "Company");
            selectCompanyPage.ClickOkButton();

            homePage.ClickHomeTab();
            homePage.ClickBondManagementTab();
            
            homePage.ClickPaymentsLink();
            homePage.ClickBPayFileLink(); 

            BPayFilePage bPayFilePage = new BPayFilePage();
            bPayFilePage.ClickImportMenuItem();

            BPayFileImportPage bPayFileImportPage = new BPayFileImportPage();
            Assert.IsTrue(bPayFileImportPage.GetWindowExistStatus());
            bPayFileImportPage.SetMoveFileCheckBox(true);
            bPayFileImportPage.SetProcessFileCheckBox(true);
            bPayFileImportPage.SetImportPathEdit("");
            bPayFileImportPage.SetFileNameEdit(fileLocation);
            bPayFileImportPage.ClickOKButton();

            InfoLogPage infoLogPage = new InfoLogPage();
            string rtbTreeItem; 
            rtbTreeItem = infoLogPage.GetTreeItemName("Processing BPAY file RTB-", "Importing BPAY file");
            string RTB;
            RTB = rtbTreeItem.Substring(21, 10);
            string bvdTreeItem;
            bvdTreeItem = infoLogPage.GetTreeItemName("Posting journal BDV", "Importing BPAY file");
            string BVD;
            BVD = bvdTreeItem.Substring(16, 9);

            Assert.IsTrue(infoLogPage.GetTreeItemExists("1 file(s) were processed", "Importing BPAY file"));
            infoLogPage.ClickClearButton();
            infoLogPage.ClickCloseButton();
            bPayFilePage.ClickCloseButton();


            homePage.ClickHomeTab();
            homePage.ClickCashandBankManagementTab();
            homePage.ClickPaymentsLink();
            homePage.ClickBPayFileLink();
            
            table = new Table(bPayFilePage.GetBPayFileTable());
            table.FilterCellValue("BPAY file identifier");
            bPayFilePage.ClickFilterMenuItem();
            FilterPage filterPage = new FilterPage();
            filterPage.SetFilterText("BPAY file identifier",RTB);
            filterPage.ClickOkButton();
            StringAssert.Contains(fileLocation, table.GetCellValue("BPAY file identifier", RTB, "File name"));
            
            //Import second duplicate file
            fileLocation = Utils.BPayFileCreator.bPayUnknownClientFileCreator(dateValue, randomNum);
            bPayFilePage.ClickImportMenuItem();
            bPayFileImportPage = new BPayFileImportPage();
            Assert.IsTrue(bPayFileImportPage.GetWindowExistStatus());
            bPayFileImportPage.SetMoveFileCheckBox(true);
            bPayFileImportPage.SetProcessFileCheckBox(true);
            bPayFileImportPage.SetImportPathEdit("");
            bPayFileImportPage.SetFileNameEdit(fileLocation);
            bPayFileImportPage.ClickOKButton();

            infoLogPage = new InfoLogPage();
            string ImportErrorTreeItem = infoLogPage.GetTreeItemName("was already imported on", "Importing BPAY file");
            string MovedFileMessageTreeItem = infoLogPage.GetTreeItemName("has been moved to sub-folder 'Error'", "Importing BPAY file");

            StringAssert.Contains(ImportErrorTreeItem, fileName);
            StringAssert.Contains(MovedFileMessageTreeItem, fileName);

            infoLogPage.ClickClearButton();
            infoLogPage.ClickCloseButton();
            bPayFilePage.ClickCloseButton();


        }

        [TestMethod]
        [TestProperty("TestcaseID", "6365b")]
        public void ATC6365b_AXTESTINGEndtoEndSingleFormBPayAXSuccess()
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
                if (MyRange.Cells[i, 1].Value.ToString() == "6365")
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
            homePage.ClickBondManagementTab();
            homePage.ClickPaymentsLink();
            homePage.ClickBPayFileLink();


            BPayFilePage bpayFile = new BPayFilePage();
            bpayFile.ClickImportMenuItem();

            BPayFileImportPage bpayFileImportPage = new BPayFileImportPage();
            Assert.IsTrue(bpayFileImportPage.GetWindowExistStatus());
            bpayFileImportPage.SetMoveFileCheckBox(true);



            bpayFileImportPage.SetProcessFileCheckBox(true);
            bpayFileImportPage.SetImportPathText("");
            bpayFileImportPage.SetFileNameEdit(filelocation);
            bpayFileImportPage.ClickOKButton();

            InfoLogPage infoLogPage = new InfoLogPage();
            //check that file imported successful
            Assert.IsTrue(infoLogPage.GetTreeItemExists("File imported with identifier", "Importing BPAY file"));

            infoLogPage.ClickClearButton();
            infoLogPage.ClickCloseButton();
            bpayFile.ClickCloseButton();

            homePage.ClickCRMOutboundNotificationsLink();
            OutboundCRMIntegrationPage outboundCRMIntegrationPage = new OutboundCRMIntegrationPage();
            outboundCRMIntegrationPage.ClickOKButton();

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

            StringAssert.Contains(table.GetCellValue("Payment reference", paymentreference, "Bond Balance"), initialContribution);
            MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("BOND_REF")].Value = table.GetCellValue("Payment reference", paymentreference, "Bond");

            

            #region Shut down Excel
            MyBook.Save();
            MyBook.Close();
            MyApp.Quit();
            #endregion

        }


        [TestMethod]
        [TestProperty("TestcaseID", "6370")]
        public void ATC6370_AXFinpostingBpayreversal()
        {


            string dateValue = DateTime.Today.AddDays(-1).ToString("yyyyMMdd");
            string transDesc = "BPAY " + DateTime.Today.AddDays(-1).ToString("d/MM/yyyy");
            Random random = new Random();
            int randomNum = random.Next(1000, 9999);
            string fileLocation = Utils.BPayFileCreator.bPayUnknownClientFileCreator(dateValue, randomNum);
            string referenceNumber = Utils.BPayFileReaderClass.GetPaymentReference1File(fileLocation);

            Homepage homePage = new Homepage();
            homePage.ClickCompanyButton();

            SelectCompanyPage selectCompanyPage = new SelectCompanyPage();
            Table table = new Table(selectCompanyPage.GetCompanyListTable());
            table.ClickCellValue("Company", "RTB", "Company");
            selectCompanyPage.ClickOkButton();

            homePage.ClickHomeTab();
            homePage.ClickBondManagementTab();
            


            homePage.ClickPaymentsLink();
            homePage.ClickBPayFileLink(); ;

            BPayFilePage bPayFilePage = new BPayFilePage();
            bPayFilePage.ClickImportMenuItem();

            BPayFileImportPage bPayFileImportPage = new BPayFileImportPage();
            Assert.IsTrue(bPayFileImportPage.GetWindowExistStatus());
            bPayFileImportPage.SetMoveFileCheckBox(true);
            bPayFileImportPage.SetProcessFileCheckBox(true);
            //bPayFileImportPage.SetImportPathEdit(@"P:\Dynamics AX\Bank files\Bpay\Paul");
            bPayFileImportPage.SetImportPathEdit("");
            bPayFileImportPage.SetFileNameEdit(fileLocation);
            bPayFileImportPage.ClickOKButton();

            InfoLogPage infoLogPage = new InfoLogPage();
            string rtbTreeItem = infoLogPage.GetTreeItemName("Processing BPAY file RTB-", "Importing BPAY file");
            string RTB = rtbTreeItem.Substring(21, 10);
            string bvdTreeItem = infoLogPage.GetTreeItemName("Posting journal BDV", "Importing BPAY file");
            string BVD = bvdTreeItem.Substring(16, 9);
            
            //Assert.IsTrue(infoLogPage.GetTreeItemExists("1 file(s) were processed", "Importing BPAY file"));
            infoLogPage.ClickClearButton();
            infoLogPage.ClickCloseButton();
            bPayFilePage.ClickCloseButton();


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



            StringAssert.Contains(table.GetCellValue("Bond journal", BVD, "Description"), RTB);
            table.ClickCellValue("Bond journal", BVD, "Bond journal");

            bondReceiptJournalPage.ClickLinesMenuItem();

            BondReceiptJournalLinesPage bondReceiptJournalLines = new BondReceiptJournalLinesPage();
            bondReceiptJournalLines.ClickInquiryButton();

            bondReceiptJournalLines.ClickVoucherTransactionsMenuItem();

            VoucherTransactionsPage voucherTransactionsPage = new VoucherTransactionsPage();
            table = new Table(voucherTransactionsPage.GetVoucherTransactionTable());

            StringAssert.Contains(table.GetCellValue("Ledger account", "11330", "Amount"), "700.00");
            table.ClickCellValue("Ledger account", "11330", "Ledger account");
            StringAssert.Contains(voucherTransactionsPage.GetDescriptionText(), transDesc);
            StringAssert.Contains(voucherTransactionsPage.GetAccountNameText(), "Cash at Bank - Rental Bond BPAY Account");

            StringAssert.Contains(table.GetCellValue("Ledger account", "32130", "Amount"), "700.00");
            table.ClickCellValue("Ledger account", "32130", "Ledger account");
            StringAssert.Contains(voucherTransactionsPage.GetDescriptionText(), referenceNumber);
            StringAssert.Contains(voucherTransactionsPage.GetAccountNameText(), "Rental bonds - Receipts unallocated");

            voucherTransactionsPage.ClickCloseButton();
            bondReceiptJournalLines.ClickCloseButton();
            bondReceiptJournalPage.ClickCloseButton();

        }
        

         [TestMethod]
        [TestProperty("TestcaseID", "6740")]
        public void ATC6740_AXPostBPAYpaymentswithunknownreferencenumber()
        {

            
            string dateValue = DateTime.Today.AddDays(-1).ToString("yyyyMMdd");
            string transDesc = "BPAY " + DateTime.Today.AddDays(-1).ToString("d/MM/yyyy");
            Random random = new Random();
            int randomNum = random.Next(1000, 9999);
            string fileLocation = Utils.BPayFileCreator.bPayUnknownClientFileCreator(dateValue, randomNum);
            string referenceNumber = Utils.BPayFileReaderClass.GetPaymentReference1File(fileLocation).ToString();

            Homepage homePage = new Homepage();
            homePage.ClickCompanyButton();

            SelectCompanyPage selectCompanyPage = new SelectCompanyPage();
            Table table = new Table(selectCompanyPage.GetCompanyListTable());
            table.ClickCellValue("Company", "RTB", "Company");
            selectCompanyPage.ClickOkButton();

            homePage.ClickHomeTab();
            homePage.ClickBondManagementTab();
            


            homePage.ClickPaymentsLink();
            homePage.ClickBPayFileLink(); ;

            BPayFilePage bPayFilePage = new BPayFilePage();
            bPayFilePage.ClickImportMenuItem();

            BPayFileImportPage bPayFileImportPage = new BPayFileImportPage();
            Assert.IsTrue(bPayFileImportPage.GetWindowExistStatus());
            bPayFileImportPage.SetMoveFileCheckBox(true);
            bPayFileImportPage.SetProcessFileCheckBox(true);
            //bPayFileImportPage.SetImportPathEdit(@"P:\Dynamics AX\Bank files\Bpay\Paul");
            bPayFileImportPage.SetImportPathEdit("");
            bPayFileImportPage.SetFileNameEdit(fileLocation);
            bPayFileImportPage.ClickOKButton();

            InfoLogPage infoLogPage = new InfoLogPage();
            string rtbTreeItem = infoLogPage.GetTreeItemName("Processing BPAY file RTB-", "Importing BPAY file");
            string RTB = rtbTreeItem.Substring(21, 10);
            string bvdTreeItem = infoLogPage.GetTreeItemName("Posting journal BDV", "Importing BPAY file");
            string BVD = bvdTreeItem.Substring(16, 9);
            Assert.IsTrue(infoLogPage.GetTreeItemExists("BPAY exception 'Invalid CRN' has been recorded for this line", "Importing BPAY file"));
            infoLogPage.ClickClearButton();
            infoLogPage.ClickCloseButton();
            bPayFilePage.ClickCloseButton();


            homePage.ClickHomeTab();
            homePage.ClickBondManagementTab();

            homePage.ClickCRMOutboundNotificationsLink();
            OutboundCRMIntegrationPage outboundCRMIntegrationPage = new OutboundCRMIntegrationPage();
            outboundCRMIntegrationPage.ClickOKButton();

            //Navigate to Payments/BPay file exceptions
            homePage.ClickPaymentsLink();
            homePage.ClickBPayFileExceptionLink();

            //Confirm customer referance number exists
            BPayFileExceptionsPage bPayFileExceptionPage = new BPayFileExceptionsPage();

            table = new Table(bPayFileExceptionPage.GetFileExceptionTable());
            table.FilterCellValue("Customer reference number");
            bPayFileExceptionPage.ClickFilterMenuItem();
            FilterPage filterPage = new FilterPage();
            filterPage.SetFilterText("Customer reference number", referenceNumber);
            filterPage.ClickOkButton();

            StringAssert.Contains(table.GetCellValue("Customer reference number", referenceNumber, "Rejection reason"), "Invalid CRN");
            bPayFileExceptionPage.ClickCloseButton();

        }

         [TestMethod]
         [TestProperty("TestcaseID", "6837b")]
         public void ATC6837b_AXCancelBPayRequestBatchWithTopup()
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
                 if (MyRange.Cells[i, 1].Value.ToString() == "Topup_6837")
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
             homePage.ClickBondManagementTab();
             homePage.ClickPaymentsLink();
             homePage.ClickBPayFileLink();


             BPayFilePage bpayFile = new BPayFilePage();
             bpayFile.ClickImportMenuItem();

             BPayFileImportPage bpayFileImportPage = new BPayFileImportPage();
             Assert.IsTrue(bpayFileImportPage.GetWindowExistStatus());
             bpayFileImportPage.SetMoveFileCheckBox(true);



             bpayFileImportPage.SetProcessFileCheckBox(true);
             bpayFileImportPage.SetImportPathText("");
             bpayFileImportPage.SetFileNameEdit(filelocation);
             bpayFileImportPage.ClickOKButton();

             InfoLogPage infoLogPage = new InfoLogPage();
             //check that file imported successful
             Assert.IsTrue(infoLogPage.GetTreeItemExists("File imported with identifier", "Importing BPAY file"));

             infoLogPage.ClickClearButton();
             infoLogPage.ClickCloseButton();
             bpayFile.ClickCloseButton();

             homePage.ClickCRMOutboundNotificationsLink();
             OutboundCRMIntegrationPage outboundCRMIntegrationPage = new OutboundCRMIntegrationPage();
             outboundCRMIntegrationPage.ClickOKButton();

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

             StringAssert.Contains(table.GetCellValue("Payment reference", paymentreference, "Bond Balance"), initialContribution);
             MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("BOND_REF")].Value = table.GetCellValue("Payment reference", paymentreference, "Bond");



             #region Shut down Excel
             MyBook.Save();
             MyBook.Close();
             MyApp.Quit();
             #endregion

         }

         [TestMethod]
         [TestProperty("TestcaseID", "6904b")]
         public void ATC6904b_AXTopupExcessBondValidation()
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
                 if (MyRange.Cells[i, 1].Value.ToString() == "Topup_6904")
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
             homePage.ClickBondManagementTab();
             homePage.ClickPaymentsLink();
             homePage.ClickBPayFileLink();


             BPayFilePage bpayFile = new BPayFilePage();
             bpayFile.ClickImportMenuItem();

             BPayFileImportPage bpayFileImportPage = new BPayFileImportPage();
             Assert.IsTrue(bpayFileImportPage.GetWindowExistStatus());
             bpayFileImportPage.SetMoveFileCheckBox(true);



             bpayFileImportPage.SetProcessFileCheckBox(true);
             bpayFileImportPage.SetImportPathText("");
             bpayFileImportPage.SetFileNameEdit(filelocation);
             bpayFileImportPage.ClickOKButton();

             InfoLogPage infoLogPage = new InfoLogPage();
             //check that file imported successful
             Assert.IsTrue(infoLogPage.GetTreeItemExists("File imported with identifier", "Importing BPAY file"));

             infoLogPage.ClickClearButton();
             infoLogPage.ClickCloseButton();
             bpayFile.ClickCloseButton();

             homePage.ClickCRMOutboundNotificationsLink();
             OutboundCRMIntegrationPage outboundCRMIntegrationPage = new OutboundCRMIntegrationPage();
             outboundCRMIntegrationPage.ClickOKButton();

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

             StringAssert.Contains(table.GetCellValue("Payment reference", paymentreference, "Bond Balance"), initialContribution);
             MyRange.Cells[MyRow, TenancyRequestSchema.GetColumnIndex("BOND_REF")].Value = table.GetCellValue("Payment reference", paymentreference, "Bond");



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

        public UIMap UIMap
        {
            get
            {
                if ((this.map == null))
                {
                    this.map = new UIMap();
                }

                return this.map;
            }
        }

        private UIMap map;

      

    }

   
}
