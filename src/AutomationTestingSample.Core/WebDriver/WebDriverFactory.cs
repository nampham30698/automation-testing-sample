using AutomationTestingSample.Core.Configurations;
using AutomationTestingSample.Core.Constants;
using AutomationTestingSample.Core.Helpers;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Safari;

namespace AutomationTestingSample.Core.WebDriver
{
    public class WebDriverFactory
    {
        static readonly bool _enableHeadless = ConfigurationManager.GetValue<bool>(AppSettingConstants.EnableHeadless);

        public static IWebDriver CreateWebDriver(string browser, int width, int height)
        {
            return browser switch
            {
                BrowserConstants.Chrome => CreateChromeDriver(width, height),
                BrowserConstants.Edge => CreateEdgeDriver(width, height),
                BrowserConstants.Firefox => CreateFirefoxDriver(width, height),
                _ => CreateSafariDriver(width, height),
            };
        }

        private static ChromeDriver CreateChromeDriver(int browserWidth, int browserHeight)
        {
            var chromeOptions = new ChromeOptions();
            chromeOptions.AddUserProfilePreference("download.default_directory", FileHelpers.Actuals_Downloads);
            chromeOptions.AddUserProfilePreference("download.prompt_for_download", false);
            chromeOptions.AddUserProfilePreference("plugins.always_open_pdf_externally", true);
            chromeOptions.AddUserProfilePreference("disable-popup-blocking", "true");
            chromeOptions.AddArgument($"--window-size={browserWidth},{browserHeight}");
            chromeOptions.AddExcludedArgument("enable-automation");
            if (_enableHeadless)
            {
                chromeOptions.AddArgument("--headless");
            }

            return new ChromeDriver(chromeOptions);
        }

        private static EdgeDriver CreateEdgeDriver(int browserWidth, int browserHeight)
        {
            var edgeoptions = new EdgeOptions();
            edgeoptions.AddUserProfilePreference("download.default_directory", FileHelpers.Actuals_Downloads);
            edgeoptions.AddArgument("--inprivate");
            edgeoptions.AddArgument($"--window-size={browserWidth},{browserHeight}");
            edgeoptions.AddExcludedArgument("enable-automation");
            if (_enableHeadless)
            {
                edgeoptions.AddArgument("--headless");
            }

            return new EdgeDriver(EdgeDriverService.CreateDefaultService(".", "msedgedriver.exe"), edgeoptions);
        }

        private static FirefoxDriver CreateFirefoxDriver(int browserWidth, int browserHeight)
        {
            var options = new FirefoxOptions();
            options.SetPreference("browser.download.folderList", 2);
            options.SetPreference("browser.download.dir", FileHelpers.Actuals_Downloads);
            options.SetPreference("pdfjs.disabled", true);
            options.SetPreference("app.update.enabled", false);
            options.AddArguments($"-width={browserWidth}");
            options.AddArguments($"-height={browserHeight}");
            if (_enableHeadless)
            {
                // in headless mode, need to set withd and height 
                //https://github.com/mozilla/geckodriver/issues/1354
                Environment.SetEnvironmentVariable("MOZ_HEADLESS_WIDTH", $"{browserWidth}");
                Environment.SetEnvironmentVariable("MOZ_HEADLESS_HEIGHT", $"{browserHeight}");
                options.AddArgument("-headless");
            }

            return new FirefoxDriver(options);
        }

        private static SafariDriver CreateSafariDriver(int browserWidth, int browserHeight)
        {
            return new SafariDriver();
        }
    }
}
