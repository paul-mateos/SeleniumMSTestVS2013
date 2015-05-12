using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RTA.Automation.CRM.Utils
{
    public class ExpectedConditionsExtensions
    {

        public static Func<IWebDriver, bool> ElementIsInvisible(By locator) 
        {
            return (driver) =>
            {
                if (driver.FindElement(locator).Displayed)
                {
                    return false;
                }

                return true;
            };
        }

        public static Func<IWebDriver, bool> ElementIsNotZero(By locator)
        {
            return (driver) =>
            {
                string text = "";
                text = driver.FindElement(locator).Text;
                if ( text != "0" && text != "")
                {
                    return false;
                }

                return true;
            };
        }

        public static Func<IWebDriver, bool> ElementNotContains(By locator, string value)
        {
            return (driver) =>
            {
                string text = "";
                text = driver.FindElement(locator).GetAttribute("title");
                if (text != value && text != "")
                {
                    return false;
                }

                return true;
            };
        }
    }
}
