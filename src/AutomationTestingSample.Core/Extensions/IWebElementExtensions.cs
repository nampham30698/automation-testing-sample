using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;

namespace AutomationTestingSample.Core.Extensions
{
    public static class IWebElementExtensions
    {
        public static void EnterText(this IWebElement element, string value)
        {
            element.Clear();
            element.SendKeys(value);
        }

        //public static void SelectDropdownByText(this IWebElement element, string text)
        //{
        //    var selectElement = new SelectElement(element);
        //    selectElement.SelectByText(text);
        //}

        //public static void SelectDropdownByValue(this IWebElement element, string value)
        //{
        //    var selectElement = new SelectElement(element);
        //    selectElement.SelectByValue(value);
        //}

        //public static void MultipleSelectElements(this IWebElement element, string[] values)
        //{
        //    var selectElement = new SelectElement(element);

        //    foreach (var value in values)
        //    {
        //        selectElement.SelectByValue(value);
        //    }
        //}

        public static IEnumerable<T> GetAllSelectedList<T>(this IWebElement element)
        {
            var selectElement = new SelectElement(element);

            return selectElement.AllSelectedOptions.Select(x => GetValue<T>(x.Text));
        }

        private static T GetValue<T>(string value)
        {
            return (T)Convert.ChangeType(value, typeof(T));
        }
    }
}
