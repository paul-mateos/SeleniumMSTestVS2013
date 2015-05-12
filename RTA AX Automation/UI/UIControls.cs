using ActionWordsLib.Attributes;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.VisualStudio.TestTools.UITesting.WpfControls;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RTA.Automation.AX.UI
{
    public class UIControls
    {

        [ActionMethod]
        public static WpfControl GetUIAControl(string name, string type, WinWindow parent)
        {
            WpfControl uIControl = new WpfControl(parent);
            uIControl.TechnologyName = "UIA";
            uIControl.SearchProperties.Add("ControlType", type);
            uIControl.SearchProperties.Add("AutomationId", name);
            uIControl.WaitForControlReady();
            return uIControl;

        }

        public static void SetControlValue(string name, string type, string value, WinWindow parent)
        {
            WinControl uIControl;
            uIControl = new WinControl(parent);
            uIControl.TechnologyName = "MSAA";
            uIControl.SearchProperties.Add("ControlType", type);
            uIControl.SearchProperties.Add("Name", name, PropertyExpressionOperator.Contains);
            uIControl.WaitForControlReady();

            if (type == "CheckBox")
            {
                WinCheckBox mUICheckBox = new WinCheckBox(uIControl);
                mUICheckBox.Checked = Convert.ToBoolean(value);
            }
            else if (type == "Edit")
            {
                WinEdit mUIEdit = new WinEdit(uIControl);
                mUIEdit.Text = value;
            }
        }


        public static void SetItemControlValue(string name, string type, string value, WinWindow parent)
        {

            WinWindow ItemWindow = new WinWindow(parent);
            ItemWindow.SearchProperties.Add("AccessibleName", name);
            ItemWindow.WaitForControlReady();
            WinControl uIControl = new WinControl(ItemWindow);
            uIControl.TechnologyName = "MSAA";
            uIControl.SearchProperties.Add("ControlType", type);
            uIControl.SearchProperties.Add("Name", name);
            uIControl.WaitForControlReady();

            if (type == "CheckBox")
            {
                WinCheckBox mUICheckBox = new WinCheckBox(uIControl);
                mUICheckBox.Checked = Convert.ToBoolean(value);
            }
            else if (type == "Edit")
            {
                WinEdit mUIEdit = new WinEdit(uIControl);
                mUIEdit.Text = value;
            }
        }



        public static void SetUIAControlValue(string name, string type, string value, WinWindow parent)
        {

            WpfControl uIControl = new WpfControl(parent);
            uIControl.TechnologyName = "UIA";
            uIControl.SearchProperties.Add("ControlType", type);
            uIControl.SearchProperties.Add("AutomationId", name);
            uIControl.WaitForControlReady();

            if (type == "CheckBox")
            {
                WpfCheckBox mUICheckBox = new WpfCheckBox(uIControl);
                mUICheckBox.Checked = Convert.ToBoolean(value);
            }
            else if (type == "Edit")
            {
                WpfEdit mUIEdit = new WpfEdit(uIControl);
                mUIEdit.Text = value;
            }
        }


        public static WinControl GetControl(string name, string type, WinWindow parent)
        {

            WinControl uIControl = new WinControl(parent);
            uIControl.SearchProperties.Add("ControlType", type);
            uIControl.SearchProperties.Add("Name", name);
            uIControl.WaitForControlReady();
            return uIControl;

        }

        public static WinControl GetChildControl(string name, WinControl parentControl, string type)
        {
            WinControl uIControl = new WinControl(parentControl);
            uIControl.SearchProperties.Add("ControlType", type);
            uIControl.SearchProperties.Add("Name", name);
            uIControl.WaitForControlReady();
            return uIControl;

        }

        public static WinControl GetButtonGroupControl(string name, string type, WinWindow parent)
        {
            WinGroup buttonGroup = new WinGroup(parent);
            buttonGroup.SearchProperties.Add("Name", "ButtonGroup");
            buttonGroup.SearchProperties.Add("ControlType", "Group");
            buttonGroup.WaitForControlReady();
            
            WinControl uIControl = new WinControl(buttonGroup);
            uIControl.SearchProperties.Add("ControlType", type);
            uIControl.SearchProperties.Add("Name", name);
            uIControl.WaitForControlReady();
            
            return uIControl;
        }

        public static WinControl GetDropDownControl(string name, string type, WinWindow parent)
        {
            WinWindow dropDownWindow = new WinWindow();
            dropDownWindow.SearchProperties[WinWindow.PropertyNames.AccessibleName] = "DropDown";
            dropDownWindow.SearchProperties.Add(new PropertyExpression(WinWindow.PropertyNames.ClassName, "WindowsForms10.Window", PropertyExpressionOperator.Contains));
            dropDownWindow.WaitForControlReady();
            WinMenu menu = new WinMenu(dropDownWindow);
            menu.SearchProperties[WinMenu.PropertyNames.Name] = "DropDown";
            menu.WaitForControlReady();
            WinControl uIControl = new WinControl(menu);
            uIControl.SearchProperties.Add("ControlType", type);
            uIControl.SearchProperties.Add("Name", name);
            uIControl.WaitForControlReady();
            return uIControl;

        }




        public static string GetControlValue(string name, string type, WinWindow parent)
        {
            WinControl uIControl = new WinControl(parent);
            uIControl.TechnologyName = "MSAA";
            uIControl.SearchProperties.Add("ControlType", type);
            uIControl.SearchProperties.Add("Name", name);
            uIControl.WaitForControlReady();

            if (type == "CheckBox")
            {
                WinCheckBox mUICheckBox = new WinCheckBox(uIControl);
                return mUICheckBox.Checked.ToString();
            }
            else if (type == "Edit")
            {
                WinEdit mUIEdit = new WinEdit(uIControl);
                return mUIEdit.Text;
            }
            else
            {
                throw new Exception(String.Format("Unknown control. contact automation administrator"));
            }

        }


        public static WinControl GetItemControl(string name, string type, WinWindow parent)
        {


            WinWindow ItemWindow = new WinWindow(parent);
            ItemWindow.SearchProperties.Add("AccessibleName", name);
            ItemWindow.WaitForControlReady();
            WinControl uIControl = new WinControl(ItemWindow);
            uIControl.SearchProperties.Add("ControlType", type);
            uIControl.SearchProperties.Add("Name", name);
            uIControl.WaitForControlReady();
            return uIControl;

        }

        public static WinControl GetItemWindowControl(string name, string type, WinWindow parent)
        {

            WinWindow itemWindow = new WinWindow(parent);
            itemWindow.SearchProperties.Add(new PropertyExpression(WinWindow.PropertyNames.ClassName, "WindowsForms10.Window", PropertyExpressionOperator.Contains));
            itemWindow.SearchProperties[WinWindow.PropertyNames.Instance] = "11";
            itemWindow.WaitForControlReady();
            WinControl uIControl = new WinControl(itemWindow);
            uIControl.SearchProperties.Add("ControlType", type);
            uIControl.SearchProperties.Add("Name", name);
            uIControl.WaitForControlReady();
            return uIControl;

        }


        public static WinControl GetTreeItemControl(string name, WinWindow parent)
        {

            WinWindow treeWindow = new WinWindow(parent);
            treeWindow.SearchProperties.Add(new PropertyExpression(WinWindow.PropertyNames.Name, "Tree", PropertyExpressionOperator.Contains));
            treeWindow.WaitForControlReady();
            WinControl uIControl = new WinControl(treeWindow);
            uIControl.SearchProperties.Add("ControlType", "TreeItem");
            uIControl.SearchProperties.Add("Name", name);
            uIControl.WaitForControlReady();
            return uIControl;

        }

        public static WinControl GetChildTreeItemControl(string name, WinControl parent)
        {

            WinControl uIControl = new WinControl(parent);
            uIControl.SearchProperties.Add("ControlType", "TreeItem");
            uIControl.SearchProperties.Add("Name", name);
            uIControl.WaitForControlReady();
            return uIControl;

        }


        public static WinControl GetGroupedControl(string name, string type, string group, WinWindow parent)
        {


            WinGroup uIGroup = new WinGroup(parent);
            uIGroup.SearchProperties.Add("Name", group);
            uIGroup.WaitForControlReady();
            WinControl uIControl = new WinControl(uIGroup);
            uIControl.SearchProperties.Add("ControlType", type);
            uIControl.SearchProperties.Add("Name", name);
            uIControl.WaitForControlReady();
            return uIControl;

        }

        public static WinGroup GetGroupControl(string groupName, WinWindow parent)
        {
            WinWindow groupWindow = new WinWindow(parent);
            groupWindow.SearchProperties.Add("AccessibleName", groupName);
            WinGroup uIGroup = new WinGroup(groupWindow);
            uIGroup.SearchProperties.Add("Name", groupName);
            uIGroup.WaitForControlReady();
            return uIGroup;
        }

        public static WinHyperlink GetGroupedHyperlinkControl(string hyperlinkName, string groupName, WinWindow parent)
        {
            WinGroup group = GetGroupControl(groupName, parent);
            WinHyperlink hyperlinkItem = new WinHyperlink(group);
            hyperlinkItem.SearchProperties.Add("Name", hyperlinkName);
            hyperlinkItem.WaitForControlReady();
            return hyperlinkItem;
        }



        public static WinControl GetTabGroupedControl(string name, string type, string tabName, string group, WinWindow parent)
        {

            WinWindow uITab = new WinWindow(parent);
            uITab.SearchProperties.Add("ControlName", tabName);
            uITab.SearchProperties.Add("ControlType", "Window");
            uITab.WaitForControlReady();
            WinGroup uIGroup = new WinGroup(uITab);
            uIGroup.SearchProperties.Add("Name", group);
            uIGroup.WaitForControlReady();
            WinControl uIControl = new WinControl(uIGroup);
            uIControl.SearchProperties.Add("ControlType", type);
            uIControl.SearchProperties.Add("Name", name);
            uIControl.WaitForControlReady();
            return uIControl;

        }

        [ActionMethod]
        public static WinControl GetRightClickControl(string name, string type)
        {
            WinWindow uIWindow = new WinWindow();
            uIWindow.SearchProperties[WinWindow.PropertyNames.AccessibleName] = "Context";
            uIWindow.SearchProperties[WinWindow.PropertyNames.ClassName] = "#32768";
            uIWindow.WaitForControlReady();
            WinMenu uIMenu = new WinMenu(uIWindow);
            uIMenu.SearchProperties[WinMenu.PropertyNames.Name] = "Context";
            uIMenu.WaitForControlReady();
            WinControl uIControl = new WinControl(uIMenu);
            uIControl.SearchProperties.Add("ControlType", type);
            uIControl.SearchProperties.Add("Name", name);
            uIControl.WaitForControlReady();
            return uIControl;

        }

        [ActionMethod]
        public static void ClickContextMenuItem(string item)
        {

            WinMenuItem uIMenuItem = new WinMenuItem(GetRightClickControl(item, "MenuItem"));
            Mouse.Click(uIMenuItem, new Point(uIMenuItem.Width / 2, uIMenuItem.Height / 2));

        }



       

        public static void ClickMaximizeButton(WinWindow parentWindow)
        {
            
            if (parentWindow.Maximized == false)
            {
                WinTitleBar titleBar = new WinTitleBar(parentWindow);
                WinButton maxButton = new WinButton(titleBar);
                maxButton.SearchProperties.Add("Name", "Maximize");
                maxButton.WaitForControlReady();
                Mouse.Click(maxButton, new Point(maxButton.Width / 2, maxButton.Height / 2));
            }

        }

        public static void SetUIAItemControlValue(string name, string type, string value, WinWindow parent)
        {
            WinWindow itemWindow = new WinWindow(parent);
            itemWindow.SearchProperties.Add(new PropertyExpression(WinWindow.PropertyNames.ClassName, "HwndWrapper", PropertyExpressionOperator.Contains));
            itemWindow.WaitForControlReady();
            WpfPane pane = new WpfPane(itemWindow);
            WpfCustom customControl = new WpfCustom(pane);
            customControl.SearchProperties[WpfControl.PropertyNames.ClassName] = "Uia.SegmentedEntry";
            customControl.WaitForControlReady();
            WpfControl uIControl = new WpfControl(customControl);
            uIControl.TechnologyName = "UIA";
            uIControl.SearchProperties.Add("ControlType", type);
            uIControl.SearchProperties.Add("AutomationId", name);
            uIControl.WaitForControlReady();

            if (type == "CheckBox")
            {
                WpfCheckBox mUICheckBox = new WpfCheckBox(uIControl);
                mUICheckBox.Checked = Convert.ToBoolean(value);
            }
            else if (type == "Edit")
            {
                WpfEdit mUIEdit = new WpfEdit(uIControl);
                mUIEdit.Text = value;
            }
        }

        public static void ClickMinimizeMenuGroup(string groupName, WinWindow parent)
        {
            //Minimise all group controls
            WinGroup group = UIControls.GetGroupControl(groupName, parent);

            if (group.Height > 29)
            {
                WinButton arrowButton = new WinButton(UIControls.GetGroupedControl("CollapseChevron", "Button", groupName, parent));
                Mouse.Click(arrowButton, new Point(arrowButton.Width / 2, arrowButton.Height / 2));

            }
        }

        public static void ClickMaximizeMenuGroup(string groupName, WinWindow parent)
        {
            //Minimise all group controls
            WinGroup group = UIControls.GetGroupControl(groupName, parent);

            if (group.Height <= 29)
            {
                WinButton arrowButton = new WinButton(UIControls.GetGroupedControl("CollapseChevron", "Button", groupName, parent));
                Mouse.Click(arrowButton, new Point(arrowButton.Width / 2, arrowButton.Height / 2));

            }
        }
    }
     
}
