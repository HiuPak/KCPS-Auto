using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Converters;
using System.Dynamic;
using System.Configuration;


namespace KCPS_Auto
{
    internal class Program
    {
        private static readonly string appName = ConfigurationManager.AppSettings["appName"];
        private static readonly string unitedUnitWindow  = ConfigurationManager.AppSettings["unitedUnitWindow"];
        private static readonly string jsonPath = ConfigurationManager.AppSettings["jsonPath"];
        private static readonly dynamic config = JsonConvert.DeserializeObject<ExpandoObject>(File.ReadAllText(jsonPath), new ExpandoObjectConverter());
        private static AutomationElement appElement;
        private static void Main(string[] args)
        {
            var rootElement = AutomationElement.RootElement;
            if (rootElement != null)
            {
                Condition mainWindowCondition = new PropertyCondition(AutomationElement.NameProperty, appName);
                appElement = rootElement.FindFirst(TreeScope.Children, mainWindowCondition);
                
                if (appElement != null)
                {
                    appElement.MaximizeWindow();

                    KCPS_Auto(appElement, config.workflows[0]);
                }
            }
            
        }

        private static void KCPS_Auto(AutomationElement currEle, dynamic parent)
        {
            while (true)
            {
                try
                {
                    System.Threading.Thread.Sleep(50);
                    string name, type = "";
                    int idx;
                    //bool test = ((IDictionary<String, object>)parent).ContainsKey("name");

                    if (parent.action == "AttachUnitedUnitWindow")
                    {
                        Condition unitedUnitWindowCondition = new PropertyCondition(AutomationElement.NameProperty, unitedUnitWindow);
                        currEle = AutomationElement.RootElement.FindFirst(TreeScope.Children, unitedUnitWindowCondition);
                        while (currEle == null)
                        {
                            System.Threading.Thread.Sleep(50);
                            unitedUnitWindowCondition = new PropertyCondition(AutomationElement.NameProperty, unitedUnitWindow);
                            currEle = AutomationElement.RootElement.FindFirst(TreeScope.Children, unitedUnitWindowCondition);
                        }
                        currEle.MaximizeWindow();
                    }
                    else if (parent.action == "AttachAddUnitWindow")
                    {
                        currEle = appElement.GetElementByControlType("window", 0);
                        while (currEle == null)
                        {
                            System.Threading.Thread.Sleep(50);
                            currEle = appElement.GetElementByControlType("window", 0);
                        }
                        currEle.MaximizeWindow();
                    }
                    else if (parent.type == "")
                    {
                        name = parent.name;
                        currEle = currEle.GetElementByName(name);
                        while (currEle == null)
                        {
                            System.Threading.Thread.Sleep(50);
                            currEle = currEle.GetElementByName(name);
                        }
                    }
                    else
                    {
                        bool next = parent.type.Contains("/next");
                        type = parent.type.Replace("/next", "");
                        idx = (int)parent.idx;
                        currEle = currEle.GetElementByControlType(type, idx);
                        while (currEle == null)
                        {
                            System.Threading.Thread.Sleep(50);
                            currEle = currEle.GetElementByControlType(type, idx);
                        }
                        if (next) currEle = TreeWalker.ControlViewWalker.GetNextSibling(currEle);
                    }
                    if (currEle != null)
                    {
                        try
                        {
                            switch (parent.action)
                            {
                                case "SelectionItemCombo":
                                    {
                                        var Pattern = currEle.GetCurrentPattern(ExpandCollapsePattern.Pattern) as ExpandCollapsePattern;
                                        if (Pattern != null)
                                        {
                                            Pattern.Expand();
                                            /*var selectEle = currEle.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, parent.selectName));
                                            var selectPattern = selectEle.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                                            if (selectPattern != null) selectPattern.Select();
                                            Pattern.Collapse();*/
                                        }
                                        break;
                                    }
                                case "SelectionItem":
                                    {
                                        var Pattern = currEle.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                                        if (Pattern != null) Pattern.Select();
                                        break;
                                    }
                                case "ExpandCollapse":
                                    {
                                        var Pattern = currEle.GetCurrentPattern(ExpandCollapsePattern.Pattern) as ExpandCollapsePattern;
                                        if (Pattern != null) Pattern.Expand();
                                        break;
                                    }
                                case "Invoke":
                                    {
                                        var Pattern = currEle.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                                        if (Pattern != null) Pattern.Invoke();
                                        break;
                                    }
                                case "MultiSelectionItem":
                                    {
                                        var Pattern = currEle.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                                        if (Pattern != null) Pattern.AddToSelection();
                                        break;
                                    }
                                case "Toggle":
                                    {
                                        var Pattern = currEle.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                                        if (Pattern != null)
                                        {
                                            if (parent.check && Pattern.Current.ToggleState == ToggleState.Off)
                                            {
                                                Pattern.Toggle();
                                            }
                                            else if (!parent.check && Pattern.Current.ToggleState == ToggleState.On)
                                            {
                                                Pattern.Toggle();
                                            }
                                        }
                                        break;
                                    }
                                case "Value":
                                    {
                                        string value = parent.value;
                                        if (value == "random")
                                        {
                                            var r = new Random();
                                            int rnd = r.Next((int)parent.min, (int)parent.max);
                                            value = rnd.ToString();
                                        }
                                        var Pattern = currEle.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                                        if (Pattern != null) Pattern.SetValue(value);
                                        break;
                                    }
                                default:
                                    break;
                            }

                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                            Console.ReadKey();
                        }
                    }

                    foreach (var child in parent.children)
                    {
                        KCPS_Auto(currEle, child);

                        if (type == "combo box")
                        {
                            var Pattern = currEle.GetCurrentPattern(ExpandCollapsePattern.Pattern) as ExpandCollapsePattern;
                            if (Pattern != null) Pattern.Collapse();
                        }
                    }
                    break;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"type:{parent.type} idx:{parent.idx}");
                    Console.WriteLine(ex.Message);
                    System.Threading.Thread.Sleep(50);
                }
            }
        }

    }
    public static class ExtentionMethod
    {
        public static AutomationElement GetElementByControlType(this AutomationElement parentElement, string value,int idx)
        {
            Condition condition = new PropertyCondition(AutomationElement.LocalizedControlTypeProperty, value);
            var elem = parentElement.FindAll(TreeScope.Children, condition);
            return elem[idx];
        }
        public static AutomationElement GetElementByName(this AutomationElement parentElement, string value)
        {
            Condition condition = new PropertyCondition(AutomationElement.NameProperty, value);
            var elem = parentElement.FindFirst(TreeScope.Children, condition);
            return elem;
        }
        public static void MaximizeWindow(this AutomationElement window)
        {
            var Pattern = window.GetCurrentPattern(WindowPattern.Pattern) as WindowPattern;
            if (Pattern != null) Pattern.SetWindowVisualState(WindowVisualState.Maximized);
        }
    }

}
