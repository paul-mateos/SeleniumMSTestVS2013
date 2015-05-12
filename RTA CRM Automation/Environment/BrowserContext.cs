using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using System;
using System.IO;
using System.Reflection;

namespace RTA.Automation.CRM.Environment
{
    public class BrowserContext
    {
        private string driversLocation = @"P:\LabsDeploymentItems";
        
        public BrowserContext() 
        {
            InitDriver(); 
        }
        
        public void InitDriver() 
        {
            if (!Directory.Exists(@"P:\LabsDeploymentItems")) throw new Exception(@"Unable to locate P:\LabsDeploymentItems");

            switch (Properties.Settings.Default.BROWSER)
            {
                case BrowserType.Chrome:
                    WebDriver = new ChromeDriver(driversLocation);
                    break;
                case BrowserType.Ie:
                    InternetExplorerOptions opts = new InternetExplorerOptions();
                    opts.EnsureCleanSession = true;
                    opts.IntroduceInstabilityByIgnoringProtectedModeSettings = true;
                    WebDriver = new InternetExplorerDriver(driversLocation, opts);
                    break;
                case BrowserType.Firefox:
                    WebDriver = new FirefoxDriver();
                    break;
                default:
                    throw new ArgumentException("Invalid BROWSER Setting has been used");
            }

            WebDriver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(Properties.Settings.Default.IMPLICIT_WAIT_SECONDS));
            WebDriver.Manage().Window.Maximize();
        }

        public IWebDriver WebDriver { get; private set; }
    }
}
