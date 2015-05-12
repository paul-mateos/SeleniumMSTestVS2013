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
using RTA_AX_Automation.Utils;
using RTA_AX_Automation.UIMaps;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using RTA_AX_Automation.Pages;


namespace RTA_AX_Automation.Tests
{

    //[CodedUITest]
    [TestClass]
    public class LoginToHomepage : TestBase
    {
        public LoginToHomepage()
        {
        }


        #region TestInitialize
        //Use TestInitialize to run code before running each test 
        [TestInitialize()]
        public override void TestInitialize()
        {
            Console.WriteLine("Initialize");
            base.TestInitialize();
        }
        #endregion

        #region Scripts
        [TestMethod]
        //[Description("This test will Login To AX")]
        //[Owner("PaulMateos")]
        //[Priority(0)]
        //[TestProperty("TestcaseID", "12341")]
        //[DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV", "|DataDirectory|\\TestData\\SearchData.csv", "SearchData#csv", DataAccessMethod.Sequential), DeploymentItem("CodedUISampleFramework\\TestData\\SearchData.csv"), TestMethod]
        public void LoginToHomapage()
        {
           //WinWindow DynamicsAXWindow = new WinWindow();

            Homepage.EnterSearchText("");
           
        }
        #endregion

        [TestCleanup()]
        public override void TestCleanup()
        {
            base.TestCleanup();
        }



    }  
}
