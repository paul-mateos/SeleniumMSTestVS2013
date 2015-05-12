using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace RTA.Automation.CRM.UI
{
    public class Table
    {
        private IWebElement element;

        public Table(IWebElement element)
        {
            this.element = element;
        }

        public string GetCellValue(string lookupColumn, string lookupValue, string returnColumn)
        {
            if(!this.element.Text.Contains("records are available in this view.")) //No Records are available
            {
                int lookupColumnIndex = this.GetColumnIndex(lookupColumn);
                int returnColumnIndex = this.GetColumnIndex(returnColumn);

                IReadOnlyCollection<IWebElement> rows = this.element.FindElements(By.CssSelector("tbody tr"));
                foreach (IWebElement row in rows)
                {
                    if(row.Text != "")
                    {
                        IReadOnlyCollection<IWebElement> cells = row.FindElements(By.TagName("td"));
                        if (cells.ElementAt(lookupColumnIndex).Text == lookupValue)
                        {
                            return cells.ElementAt(returnColumnIndex).Text;
                        }
                    }
                }
            }
            else
            {
                throw new Exception(String.Format("Unable to find record"));
            }
            throw new Exception(String.Format("Unable to find table row {0} with value {1}", lookupColumn, lookupValue));
        }

        public string GetCellContainsValue(string lookupColumn, string lookupValue, string returnColumn)
        {
            int lookupColumnIndex = this.GetColumnIndex(lookupColumn);
            int returnColumnIndex = this.GetColumnIndex(returnColumn);

            IReadOnlyCollection<IWebElement> rows = this.element.FindElements(By.CssSelector("tbody tr"));
            foreach (IWebElement row in rows)
            {
                if (row.Text != "")
                {
                    IReadOnlyCollection<IWebElement> cells = row.FindElements(By.TagName("td"));
                    if (cells.ElementAt(lookupColumnIndex).Text.Contains(lookupValue))
                    {
                        return cells.ElementAt(returnColumnIndex).Text;
                    }
                }
            }

            throw new Exception(String.Format("Unable to find table row {0} with value {1}", lookupColumn, lookupValue));
        }

        public string GetCellContainsValue(string lookupColumn, int returnRowValue)
        {
            int lookupColumnIndex = this.GetColumnIndex(lookupColumn);

            IReadOnlyCollection<IWebElement> rows = this.element.FindElements(By.CssSelector("tbody tr"));
            IReadOnlyCollection<IWebElement> cells = rows.ElementAt(returnRowValue).FindElements(By.TagName("td"));
            return cells.ElementAt(lookupColumnIndex).Text;

            throw new Exception(String.Format("Unable to find table row {0}", returnRowValue));
        }

        public string GetCellContainsValueRefreshed(string lookupColumn, string lookupValue, string returnColumn, 
            string refreshValue, int refreshTime, IWebDriver driver)
        {
            int lookupColumnIndex = this.GetColumnIndex(lookupColumn);
            int returnColumnIndex = this.GetColumnIndex(returnColumn);

            IReadOnlyCollection<IWebElement> rows = this.element.FindElements(By.CssSelector("tbody tr"));
            foreach (IWebElement row in rows)
            {
                if (row.Text != "")
                {
                    for (int i = 1; i <= refreshTime; i++)
                    {

                        IReadOnlyCollection<IWebElement> cells = row.FindElements(By.TagName("td"));
                        if (cells.ElementAt(lookupColumnIndex).Text.Contains(lookupValue))
                        {
                            if (cells.ElementAt(returnColumnIndex).Text.Contains(refreshValue))
                            {
                                return cells.ElementAt(returnColumnIndex).Text;
                            }
                    
                            else
                            {
                                Thread.Sleep(3000);
                                Actions action = new Actions(driver);
                                action.ContextClick(row).Perform();
                                IWebElement elementOpen = driver.FindElement(By.LinkText("Refresh List"));
                                elementOpen.Click();
                                rows = this.element.FindElements(By.CssSelector("tbody tr"));
                            }
                        }
                    }
                }
            }

            throw new Exception(String.Format("Unable to find table row {0} with value {1}", lookupColumn, lookupValue));
        }


        public IWebElement GetCellElementContainsValue(string lookupColumn, string lookupValue, string returnColumn)
        {
            int lookupColumnIndex = this.GetColumnIndex(lookupColumn);
            int returnColumnIndex = this.GetColumnIndex(returnColumn);

            IReadOnlyCollection<IWebElement> rows = this.element.FindElements(By.CssSelector("tbody tr"));
            foreach (IWebElement row in rows)
            {
                if (row.Text != "")
                {
                    IReadOnlyCollection<IWebElement> cells = row.FindElements(By.TagName("td"));
                    if (cells.ElementAt(lookupColumnIndex).Text.Contains(lookupValue))
                    {
                        return cells.ElementAt(returnColumnIndex);
                    }
                }
            }

            throw new Exception(String.Format("Unable to find table row {0} with value {1}", lookupColumn, lookupValue));
        }

        public void ClickCellValue(string lookupColumn, string lookupValue, string clickColumnValue)
        {
            int lookupColumnIndex = this.GetColumnIndex(lookupColumn);
            int returnColumnIndex = this.GetColumnIndex(clickColumnValue);

            IReadOnlyCollection<IWebElement> rows = this.element.FindElements(By.CssSelector("tbody tr"));
            int rowCount = rows.Count;
            int currentCount = 1;
            foreach (IWebElement row in rows)
            {
                if (row.Text != "")
                {
                    IReadOnlyCollection<IWebElement> cells = row.FindElements(By.TagName("td"));
                    if (cells.ElementAt(lookupColumnIndex).Text == lookupValue)
                    {
                        IWebElement elem = cells.ElementAt(returnColumnIndex);//.FindElement(By.XPath("//a")).Click();
                        string linkText = elem.Text;
                        //elem.FindElement(By.XPath("//td[contains(.,'" + linkText + "')]")).Click();
                        elem.FindElement(By.XPath("//a[contains(@title,'" + linkText + "')]")).Click();
                        break;
                    }
                    else if (currentCount == rowCount)
                    {
                        throw new Exception(String.Format("Unable to find table row {0} with value {1}", lookupColumn, lookupValue));
                    }
                }
            }
     
        }

        public void ClickCellContainsValue(string lookupColumn, string lookupValue, string clickColumnValue)
        {
            int lookupColumnIndex = this.GetColumnIndex(lookupColumn);
            int returnColumnIndex = this.GetColumnIndex(clickColumnValue);

            IReadOnlyCollection<IWebElement> rows = this.element.FindElements(By.CssSelector("tbody tr"));
            int rowCount = rows.Count;
            int currentCount = 1;
            foreach (IWebElement row in rows)
            {
                if (row.Text != "")
                {
                    IReadOnlyCollection<IWebElement> cells = row.FindElements(By.TagName("td"));
                    if (cells.ElementAt(lookupColumnIndex).Text.Contains(lookupValue))
                    {
                        IWebElement elem = cells.ElementAt(returnColumnIndex);//.FindElement(By.XPath("//a")).Click();
                        string linkText = elem.Text;
                        //elem.FindElement(By.XPath("//td[contains(.,'" + linkText + "')]")).Click();
                        elem.FindElement(By.XPath("//a[contains(@title,'" + linkText + "')]")).Click();
                        break;
                    }
                    else if (currentCount == rowCount)
                    {
                        throw new Exception(String.Format("Unable to find table row {0} with value {1}", lookupColumn, lookupValue));
                    }
                }
            }

        }

        public void ClickCellContainsValueEnterRow(string lookupColumn, string lookupValue, string clickColumnValue)
        {
            int lookupColumnIndex = this.GetColumnIndex(lookupColumn);
            int returnColumnIndex = this.GetColumnIndex(clickColumnValue);

            IReadOnlyCollection<IWebElement> rows = this.element.FindElements(By.CssSelector("tbody tr"));
            int rowCount = rows.Count;
            int currentCount = 1;
            foreach (IWebElement row in rows)
            {
                if (row.Text != "")
                {
                    IReadOnlyCollection<IWebElement> cells = row.FindElements(By.TagName("td"));
                    if (cells.ElementAt(lookupColumnIndex).Text.Contains(lookupValue))
                    {
                        IWebElement elem = cells.ElementAt(returnColumnIndex);
                        elem.Click();
                        elem.SendKeys(Keys.Enter);
                        break;
                    }
                    else if (currentCount == rowCount)
                    {
                        throw new Exception(String.Format("Unable to find table row {0} with value {1}", lookupColumn, lookupValue));
                    }
                }
            }

        }

        public void ClickCell(string lookupColumn, string lookupValue, string clickColumn)
        {
            int lookupColumnIndex = this.GetColumnIndex(lookupColumn);
            int returnColumnIndex = this.GetColumnIndex(clickColumn);

            IReadOnlyCollection<IWebElement> rows = this.element.FindElements(By.CssSelector("tbody tr"));
            int rowCount = rows.Count;
            int currentCount = 1;
            foreach (IWebElement row in rows)
            {
                if (row.Text != "")
                {
                    IReadOnlyCollection<IWebElement> cells = row.FindElements(By.TagName("td"));
                    if (cells.ElementAt(lookupColumnIndex).Text == lookupValue)
                    {
                        cells.ElementAt(returnColumnIndex).Click();
                        break;
                    }
                    else if (currentCount == rowCount)
                    {
                        throw new Exception(String.Format("Unable to find table row {0} with value {1}", lookupColumn, lookupValue));
                    }
                }
            }

        }

        public void SelectTableRow(string lookupColumn, string lookupValue)//, string returnColumn)
        {
           
            int lookupColumnIndex = this.GetColumnIndex(lookupColumn);
            //int returnColumnIndex = this.GetColumnIndex(returnColumn);

            IReadOnlyCollection<IWebElement> rows = this.element.FindElements(By.CssSelector("tbody tr"));
            int rowCount = rows.Count;
            int currentCount = 1;
            foreach (IWebElement row in rows)
            {
                if (row.Text != "")
                {
                    IReadOnlyCollection<IWebElement> cells = row.FindElements(By.TagName("td"));
                    if (cells.ElementAt(lookupColumnIndex).Text == lookupValue)
                    {
                        cells.ElementAt(lookupColumnIndex).Click();
                        cells.ElementAt(lookupColumnIndex).SendKeys(Keys.Enter);
                        break;
                    }
                    else if (currentCount == rowCount)
                    {
                        throw new Exception(String.Format("Unable to find table row {0} with value {1}", lookupColumn, lookupValue));
                    }
                }
            }
            
           
        }

        public void SelectContainsTableRow(string lookupColumn, string lookupValue)//, string returnColumn)
        {

            int lookupColumnIndex = this.GetColumnIndex(lookupColumn);
            if (lookupValue.Length > 30)
            {
                lookupValue = lookupValue.Substring(0,30);
            }
            //int returnColumnIndex = this.GetColumnIndex(returnColumn);

            IReadOnlyCollection<IWebElement> rows = this.element.FindElements(By.CssSelector("tbody tr"));
            int rowCount = rows.Count;
            int currentCount = 1;
            foreach (IWebElement row in rows)
            {
                if (row.Text != "")
                {
                    IReadOnlyCollection<IWebElement> cells = row.FindElements(By.TagName("td"));

                    if (cells.ElementAt(lookupColumnIndex).Text.Contains(lookupValue))
                    {
                        cells.ElementAt(lookupColumnIndex).Click();
                        cells.ElementAt(lookupColumnIndex).SendKeys(Keys.Enter);
                        break;
                    }
                    else if (currentCount == rowCount)
                    {
                        throw new Exception(String.Format("Unable to find table row {0} with value {1}", lookupColumn, lookupValue));
                    }
                }
            }


        }

        private int GetColumnIndex(string lookupColumn)
        {
            IReadOnlyCollection<IWebElement> headerCells = this.element.FindElements(By.CssSelector("thead th"));

            for (int i = 0; i < headerCells.Count; i++)
            {
                    string text = headerCells.ElementAt(i).GetAttribute("innerHTML");

                if (text.Equals(lookupColumn))
                {
                    return i;
                }
            }
            throw new Exception("Unable to find column " + lookupColumn);
        }

        public int GetRowCount()
        {
            IReadOnlyCollection<IWebElement> rows = this.element.FindElements(By.CssSelector("tbody tr"));

            foreach(IWebElement row in rows)
            {
                if(row.Text!="")
                {
                IWebElement data = row.FindElement(By.CssSelector("td"));
                if (data.GetAttribute("innerHTML").Contains("No") && data.GetAttribute("innerHTML").Contains("records are available"))
                {
                    return 0;
                }
                }
            
            }              
            return rows.Count;
        }

        public void ClickTableColumnHeader(string lookupColumn)//, string returnColumn) 
        {

            IReadOnlyCollection<IWebElement> columnHeaders = this.element.FindElements(By.CssSelector("tbody tr th"));
            int headerCount = columnHeaders.Count;
            int currentCount = 1;
            foreach (IWebElement column in columnHeaders)
            {
                
                    if (column.GetAttribute("displaylabel") == lookupColumn)
                    {
                        column.Click();
                        Thread.Sleep(2000);
                        break;
                    }
                    else if (currentCount == headerCount)
                    {
                        throw new Exception(String.Format("Unable to find table row {0}", lookupColumn));
                    }
            } 
        }

        public bool GetNoRecordsInTable()
        {

            IWebElement noRecords = this.element.FindElement(By.CssSelector("tbody tr td"));

            if (noRecords.GetAttribute("innerHTML").Contains("No"))
                {
                    return true;
                    
                }
                else 
                {
                    return false;
                }

        }

        public int GetColumnHeaderIndex(string ColumnName)
        {
            IReadOnlyCollection<IWebElement> headers = this.element.FindElements(By.CssSelector("tbody tr th"));
            int headerCount = headers.Count;
            int currentCount = 0;

            foreach (IWebElement header in headers)
            {
                currentCount += 1;
                if (header.GetAttribute("displaylabel") == ColumnName)
                {
                    return currentCount;
                }               
            }
            return 0; // If specified column not found return 0
        }

        public string GetMultiplePageTableCellValue(string lookupColumn, string lookupValue, string returnColumn)
        {
            if (!this.element.Text.Contains("records are available in this view.")) //No Records are available
            {
                int lookupColumnIndex = this.GetColumnIndex(lookupColumn);
                int returnColumnIndex = this.GetColumnIndex(returnColumn);
                IReadOnlyCollection<IWebElement> rows = this.element.FindElements(By.CssSelector("tbody tr"));
                foreach (IWebElement row in rows)
                {
                    if (row.Text != "")
                    {
                        IReadOnlyCollection<IWebElement> cells = row.FindElements(By.TagName("td"));
                        if (cells.ElementAt(lookupColumnIndex).Text == lookupValue)
                        {
                            return cells.ElementAt(returnColumnIndex).Text;
                        }
                    }
                }
            }
            else
            {
                throw new Exception(String.Format("Unable to find record"));
            }
            throw new Exception(String.Format("Unable to find table row {0} with value {1}", lookupColumn, lookupValue));
        }

        public bool MatchingCellFound(string lookupColumn, string lookupValue)
        {
            if (!this.element.Text.Contains("records are available in this view.")) //No Records are available
            {
                int lookupColumnIndex = this.GetColumnIndex(lookupColumn);

                IReadOnlyCollection<IWebElement> rows = this.element.FindElements(By.CssSelector("tbody tr"));
                foreach (IWebElement row in rows)
                {
                    if (row.Text != "")
                    {
                        IReadOnlyCollection<IWebElement> cells = row.FindElements(By.TagName("td"));
                        if (cells.ElementAt(lookupColumnIndex).Text == lookupValue)
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }
    }
}
