using AutomationTestingSample.Core.Helpers;
using OpenQA.Selenium;

namespace AutomationTestingSample.Core.Common
{
    public interface IBrowser
    {
        public void GoToUrl(string url);

        public string SaveScreenshot();

        public void SaveScreenshot(string fileName);
    }

    public class Browser : IBrowser
    {
        private IWebDriver _driver;

        public Browser(IWebDriver driver)
        {
            _driver = driver;
        }

        public void GoToUrl(string url)
        {
            _driver.Navigate().GoToUrl(url);
        }

        public string SaveScreenshot()
        {
            string filePath = FileHelpers.Actuals_Screenshots + Guid.NewGuid().ToString() + ".png";

            SaveScreenshot(filePath);

            return filePath;
        }

        public void SaveScreenshot(string fileName)
        {
            var file = ((ITakesScreenshot)_driver).GetScreenshot();

            file.SaveAsFile(fileName);
        }
    }
}
