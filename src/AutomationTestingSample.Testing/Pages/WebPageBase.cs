using AutomationTestingSample.Core.Common;
using AutomationTestingSample.Core.Reports;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;

namespace AutomationTestingSample.Testing.Pages
{
    public abstract class WebPageBase
    {
        protected IWebDriver Driver { get; private set; }

        protected IBrowser Browser { get; private set; }

        private readonly TimeSpan _pollingIntervalDefault = TimeSpan.FromMilliseconds(100);
        private readonly TimeSpan _timeoutDefault = TimeSpan.FromSeconds(10);

        public WebPageBase(IWebDriver driver)
        {
            Driver = driver;
            Browser = new Browser(driver);
        }

        protected virtual IWebElement HighlightElement(IWebElement element) {
            IJavaScriptExecutor jse = (IJavaScriptExecutor)Driver;
            jse.ExecuteScript("arguments[0].style.border='2px solid red'", element);
            return element;
        }

        protected virtual IWebElement FindElement(By by, TimeSpan? timeout = null, bool highLight = true) {

            var explicitWait = new WebDriverWait(Driver, timeout ?? _timeoutDefault)
            {
                PollingInterval = _pollingIntervalDefault
            };

            try
            {
                var element = explicitWait.Until(ExpectedConditions.ElementExists(by));

                if (highLight) {
                    HighlightElement(element);
                }

                return element;
            }
            catch (Exception ex) {
                ExtentReporting.Instance.LogFail(ex.ToString(),Browser.SaveScreenshot());
                throw;
            }
        }

        protected virtual void WaitForElementVisible(By by, TimeSpan? timeout = null)
        {

            var explicitWait = new WebDriverWait(Driver, timeout ?? _timeoutDefault)
            {
                PollingInterval = _pollingIntervalDefault
            };

            try
            {
                explicitWait.Until(ExpectedConditions.ElementIsVisible(by));
            }
            catch (Exception ex)
            {
                ExtentReporting.Instance.LogScreenshot(ex.Message, Browser.SaveScreenshot());
                throw;
            }
        }

        protected virtual void WaitForElementInVisible(By by, TimeSpan? timeout = null)
        {

            var explicitWait = new WebDriverWait(Driver, timeout ?? _timeoutDefault)
            {
                PollingInterval = _pollingIntervalDefault
            };

            try
            {
                explicitWait.Until(ExpectedConditions.InvisibilityOfElementLocated(by));
            }
            catch (Exception ex)
            {
                ExtentReporting.Instance.LogScreenshot(ex.Message, Browser.SaveScreenshot());
                throw;
            }
        }

        protected virtual void Sleep(TimeSpan timeout)
        {
            Thread.Sleep(timeout);
        }

    }
}