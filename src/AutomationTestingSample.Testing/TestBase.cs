using AutomationTestingSample.Core.Configurations;
using AutomationTestingSample.Core.Constants;
using AutomationTestingSample.Core.WebDriver;
using AutomationTestingSample.Core.Helpers;
using AutomationTestingSample.Core.Reports;
using NUnit.Framework.Interfaces;
using OpenQA.Selenium;

namespace AutomationTestingSample.Testing
{
    [TestFixtureSource(typeof(TestFixtureDataSource), nameof(TestFixtureDataSource.Browsers))]
    public class TestBase
    {
        protected IWebDriver Driver { get; private set; }

        protected string BrowserType { get; private set; }

        protected int BrowserWidth { get; private set; }

        protected int BrowserHeight { get; private set; }

        public TestBase(string browserType, int browserWidth, int browserHeight)
        {
            BrowserType = browserType;
            BrowserWidth = browserWidth;
            BrowserHeight = browserHeight;
        }

        [OneTimeSetUp]
        public virtual void OneTimeSetup()
        {
            FileHelpers.Initialize();

            ConfigurationManager.Configure();
        }

        [SetUp]
        public virtual void Setup()
        {
            var testGroupName = TestContext.CurrentContext.Test.ClassName + $"({BrowserType}-{BrowserWidth}x{BrowserHeight})";

            ExtentReporting.Instance.CreateTest(testGroupName, TestContext.CurrentContext.Test.MethodName, TestContext.CurrentContext.Test.Name);

            CreateDriver();
        }

        [TearDown]
        public virtual void TearDown()
        {
            try
            {
                EndTest();
                ExtentReporting.Instance.EndReporting();

            }
            finally
            {
                Driver.Quit();
                Driver.Dispose();
            }
        }

        [OneTimeTearDown]
        public virtual void OneTimeTearDown()
        {
            
        }

        private void CreateDriver()
        {
            Driver = WebDriverFactory.CreateWebDriver(BrowserType, BrowserWidth, BrowserHeight);

            Driver.Manage().Timeouts().ImplicitWait = new TimeSpan(0, 0, 10);

            Driver.Navigate().GoToUrl(ConfigurationManager.GetValue<string>(AppSettingConstants.BaseUrl));
        }

        private static void EndTest()
        {
            var testStatus = TestContext.CurrentContext.Result.Outcome.Status;
            var message = TestContext.CurrentContext.Result.Message;
            var stackTrace = TestContext.CurrentContext.Result.StackTrace;

            switch (testStatus)
            {
                case TestStatus.Failed:
                    ExtentReporting.Instance.LogFail($"Test has failed {message}");
                    break;
                case TestStatus.Skipped:
                    ExtentReporting.Instance.LogInfo($"Test skipped {message}");
                    ExtentReporting.Instance.LogInfo($"StackTrace: {stackTrace}");
                    break;
                default:
                    break;
            }
        }
    }
}
