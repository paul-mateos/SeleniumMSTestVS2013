using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RTA.Automation.AX.UI
{
    public class Table
    {
        private WinTable element;

        public Table(WinTable element)
        {
            this.element = element;
        }


        public static WinTable GetTable(string name, string itemWindowID, WinWindow parent)
        {
            WinWindow ItemWindow = new WinWindow(parent);
            ItemWindow.SearchProperties[WinWindow.PropertyNames.ControlId] = itemWindowID;
            ItemWindow.WaitForControlReady();
            WinTable mUITable = new WinTable(ItemWindow);
            mUITable.SearchProperties[WinTable.PropertyNames.Name] = name;
            mUITable.WaitForControlReady();
            return mUITable;
        }


      

       


        public string GetCellValue(string lookupColumn, string lookupValue, string returnColumn)
        {
              int lookupColumnIndex = this.GetColumnIndex(lookupColumn);
              int returnColumnIndex = this.GetColumnIndex(returnColumn);
              UITestControlCollection rows = this.element.Rows;
             foreach (WinRow row in rows)
            {
                int rowcount = rows.Count(); 
                 if(row.Value != "")
                  {

                      UITestControlCollection cells = row.Cells;
                      string text = cells.ElementAt(lookupColumnIndex).GetProperty("Value").ToString();
                      if (text.Equals(lookupValue))
                      {
                          return cells.ElementAt(returnColumnIndex).GetProperty("Value").ToString();
                      }           
                  }
                  
            }
            throw new Exception(String.Format("Unable to find table row {0} with value {1}", lookupColumn, lookupValue));
        }

        public bool GetCellValueExists(string lookupColumn, string lookupValue)
        {
            int lookupColumnIndex = this.GetColumnIndex(lookupColumn);
            UITestControlCollection rows = this.element.Rows;
            foreach (WinRow row in rows)
            {
                int rowcount = rows.Count();
                if (row.Value != "")
                {

                    UITestControlCollection cells = row.Cells;
                    string text = cells.ElementAt(lookupColumnIndex).GetProperty("Value").ToString();
                    if (text.Equals(lookupValue))
                    {
                        return true;
                    }
                }

            }
            return false;
        }

        public string GetCellContainsValue(string lookupColumn, string lookupValue, string returnColumn)
        {
            int lookupColumnIndex = this.GetColumnIndex(lookupColumn);
            int returnColumnIndex = this.GetColumnIndex(returnColumn);
            UITestControlCollection rows = this.element.Rows;
            foreach (WinRow row in rows)
            {
                int rowcount = rows.Count();
                if (row.Value != "")
                {

                    UITestControlCollection cells = row.Cells;
                    string text = cells.ElementAt(lookupColumnIndex).GetProperty("Value").ToString();
                    if (text.Contains(lookupValue))
                    {
                        return cells.ElementAt(returnColumnIndex).GetProperty("Value").ToString();
                    }
                }

            }
            throw new Exception(String.Format("Unable to find table row {0} with value {1}", lookupColumn, lookupValue));
        }

        public void ClickCellValue(string lookupColumn, string lookupValue, string returnColumn)
        {
            int lookupColumnIndex = this.GetColumnIndex(lookupColumn);
            int returnColumnIndex = this.GetColumnIndex(returnColumn);
            UITestControlCollection rows = this.element.Rows;
            int rowCount = rows.Count;
            int currentCount = 1;
            foreach (WinRow row in rows)
            {
                if (row.Value != "")
                {
                    UITestControlCollection cells = row.Cells;
                    string text = cells.ElementAt(lookupColumnIndex).GetProperty("Value").ToString();
                    if (text.Contains(lookupValue))
                    {
                        Mouse.Click(cells.ElementAt(returnColumnIndex));
                        break;
                    }
                    else if(currentCount == rowCount)
                    {
                        throw new Exception(String.Format("Unable to find table row {0} with value {1}", lookupColumn, lookupValue));
                    }
                }
                currentCount++;
            } 
        }    

        private int GetColumnIndex(string lookupColumn)
        {

            UITestControlCollection headerColumnCollection = this.element.ColumnHeaders;
 
            for (int i = 0; i < headerColumnCollection.Count; i++)
            {
                string text = headerColumnCollection.ElementAt(i).Name;

                if (text.Equals(lookupColumn))
                {
                    return i;
                }
            }
            throw new Exception("Unable to find column " + lookupColumn);
        }


         public void FilterCellValue(string lookupColumn)
        {
            int lookupColumnIndex = this.GetColumnIndex(lookupColumn);
            UITestControlCollection rows = this.element.Rows;
            int rowCount = rows.Count;
            if (rowCount > 0)
            {
                WinRow row = new WinRow(rows.First());
                UITestControlCollection cells = row.Cells;
                Mouse.Click(cells.ElementAt(lookupColumnIndex), MouseButtons.Left);
                Mouse.DoubleClick(cells.ElementAt(lookupColumnIndex), MouseButtons.Right);

 
            } else
            {
                throw new Exception(String.Format("Table is empty"));
            }
        }

         public bool SetCellValue(string lookupColumn, int lookupRow, string cellValue)
         {
             int lookupColumnIndex = this.GetColumnIndex(lookupColumn);

             UITestControlCollection rows = this.element.Rows;

             WinRow row = this.element.GetRow(lookupRow);
             UITestControlCollection cell = row.Cells;
             try
             {
                 cell.ElementAt(lookupColumnIndex).SetProperty("Value", cellValue);
                 return true;
             }
             catch 
             {
                 return false;
             }                
         }  
    }    
}
