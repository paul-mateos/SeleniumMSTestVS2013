using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ActionWordsLib.Execution;
using RTAAutomation.Utils;
using OpenQA.Selenium;
using ActionWordsLib.Validate;


namespace RTA_Automation_Solution
{
    [TestClass]
    public class CRMNewTenancyRequestAW
    {
        private IWebDriver driver;

        [TestMethod]
        public void CRMNewTenancyRequestTestAW()
        {
            string startupPath = System.IO.Directory.GetParent(System.IO.Directory.GetCurrentDirectory()).Parent.FullName+"\\Scripts";
            this.driver = DriverFactory.getIEDriver();
            //this.driver = new NavigateToURLWithAuth().navigateToURLWithAuth(this.driver,"http://srcrm51-te/MSCRMRTA08/main.aspx","florezj", "Dermnbr1");
            ScriptRunner runner = new ScriptRunner(new ClassConstructor(createPage));
            runner.OnException += onException;
            runner.RunAllScripts(startupPath);//"C:\\RTA Automation Solution\\RTA CRM Automation\\Scripts");
        }

        private void onException(Script script, ScriptCommand command, Exception e)
        {
            if (e is ValidateException)
            {
                Assert.Fail(e.Message);
            }
            else
            {
                throw (e);
            }
        }

        private object createPage(string className)
        {
            Type classType = Type.GetType(className);
            object classInstance = Activator.CreateInstance(classType, this.driver);
            return classInstance;

        }

        [TestCleanup()]
        public void MyTestCleanup()
        {
            this.driver.Dispose();
        }
    }
}
