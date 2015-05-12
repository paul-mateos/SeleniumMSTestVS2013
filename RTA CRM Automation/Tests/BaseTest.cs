using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using RTA.Automation.CRM.Environment;
using Microsoft.VisualStudio.TestTools.UITesting;
using System.Diagnostics;
using System.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Input;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reflection;
using System.Threading;

namespace RTA.Automation.CRM.Tests
{
    [TestClass]
    public class BaseTest
    {

        protected static IWebDriver driver;
        protected TestEnvironment environment;
        public static Excel.Workbook MyBook = null;
        public static Excel.Application MyApp = null;
        public static Excel.Worksheet MySheet = null;
        public static Excel.Range MyRange = null;
        public static string outPutDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase);
        public static string DatasourceDir = @"P:\Test Automation\SharedDatasource";

        public BaseTest()
        {
            //driver = new BrowserContext().WebDriver;
        }

        [ClassInitialize]
        public static void InitClass() {
           
        }

        [TestInitialize]
        public void TestSetup() 
        {
            KillProcess("iexplore");
            KillProcess("EXCEL");
                        
            if (Properties.Settings.Default.BROWSER == BrowserType.Ie && driver == null)
            {
                driver = new BrowserContext().WebDriver;
                driver.Manage().Cookies.DeleteAllCookies();
                this.environment = TestEnvironment.GetTestEnvironment();
            }

            driver.Navigate().GoToUrl(this.environment.Url);
        }


        [TestCleanup]
        public void TestCleanup()
        {

            if (Properties.Settings.Default.BROWSER == BrowserType.Ie)
            {
                driver.Quit();
                driver = null;
            }

            try
            {
                if (MyBook.Name != "")
                {
                    MyBook.Save();
                    MyBook.Close();
                }
                if (MyApp.Name != "")
                {
                    MyApp.Quit();
                }
            }catch
            { }

            KillProcess("EXCEL");
            KillProcess("IEDriverServer");
            KillProcess("iexplore");
            

        }

        


        public void KillProcess(string processName)
        {
            var processes = Process.GetProcessesByName(processName);

            foreach (var process in processes)
                {
                    try
                    {
                        process.Kill();
                        Thread.Sleep(1000);
                    }
                    catch { }
                }
            
        }

        //**************************************Use this for highlighting elements**************************************
        //var jsDriver = (IJavaScriptExecutor)driver;
        //var element = elem;
        //string highlightJavascript = @"$(arguments[0]).css({ ""border-width"" : ""2px"", ""border-style"" : ""solid"", ""border-color"" : ""red"" });";
        //jsDriver.ExecuteScript(highlightJavascript, new object[] { element });

        
    }
}
