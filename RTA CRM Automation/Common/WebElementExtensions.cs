using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RTA.Automation.CRM.Common
{
    public static class WebElementExtensions
    {
        public static bool IsTextInTable(this IWebElement table, string text)
        {
            var rows = table.GetTableRows();
            var isFound = false;

            foreach (var row in rows)
            {
                if (row.GetRowCells().Any(a => a.Text.Contains(text)))
                {
                    isFound = true;
                    break;
                }
            }

            return isFound;
        }

        public static ReadOnlyCollection<IWebElement> GetTableRows(this IWebElement table)
        {
            return table.FindElements(By.CssSelector("tbody tr"));
        }

        public static ReadOnlyCollection<IWebElement> GetRowCells(this IWebElement row)
        {
            return row.FindElements(By.CssSelector("td"));
        }
    }
}
