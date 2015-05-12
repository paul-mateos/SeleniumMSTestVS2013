using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Remote;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RTAAutomation.Utils
{
    public class DriverFactory
    {
        private static int waitsec = RTA.Automation.CRM.Properties.Settings.Default.IMPLICIT_WAIT_SECONDS;
        public static IWebDriver getIEDriver()
        {
            DriverFactory.DeleteIECookiesAndData();
            InternetExplorerOptions options = new InternetExplorerOptions();
            options.IntroduceInstabilityByIgnoringProtectedModeSettings = true;
            IWebDriver driver = new InternetExplorerDriver(options);
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(waitsec));
            return driver;
        }

        public static IWebDriver getFirefoxDriver()
        {
            FirefoxProfile profile = new FirefoxProfile();
            profile.SetPreference("network.proxy.type", 0);
            IWebDriver driver = new FirefoxDriver(profile);
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(waitsec));
            return driver;
        }

        private static void DeleteIECookiesAndData()
        {
            Process p = new Process();
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.FileName = "RunDll32.exe";
            p.StartInfo.Arguments = "InetCpl.cpl,ClearMyTracksByProcess 2";
            p.Start();
            p.StandardOutput.ReadToEnd();
            p.WaitForExit();
        }
    }
}
