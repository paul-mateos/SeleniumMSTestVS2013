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
using System.Diagnostics;
using System.IO;
using RTA.Automation.AX.Environments;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Threading;

namespace RTA.Automation.AX.Utils
{
    /// <summary>
    /// Summary description for CodedUITest1
    /// </summary>
    [CodedUITest]
    public class TestBase
    {
        protected TestEnvironment environment;

        public static Excel.Workbook MyBook = null;
        public static Excel.Application MyApp = null;
        public static Excel.Worksheet MySheet = null;
        public static Excel.Range MyRange = null;
        public static string outPutDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase);
        public static string DatasourceDir = @"P:\Test Automation\SharedDatasource";

        public TestBase()
        {
        }

        public static TestContext testContext
        {
            get { return TestBase.testContext; }
            set { TestBase.testContext = value; }
        }

        #region TestInitialize
        //Use TestInitialize to run code before running each test 
        [TestInitialize()]
        public virtual void TestInitialize()
        {
            KillProcess("Ax32");
            
            var pi = new ProcessStartInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), TestEnvironment.GetTestEnvironment()));
            pi.UseShellExecute = true;
            var process = Process.Start(pi);
            

        }
        #endregion
       
        
        #region TestCleanup
        //Use TestCleanup to run code after each test has run
        [TestCleanup()]
        public virtual void TestCleanup()
        {
            KillProcess("Ax32");

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
            }
            catch
            { }
            
        }
        #endregion


       
        #region TestContext
        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }
        private TestContext testContextInstance;
        #endregion

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
    }
}
