using AutomationTestingSample.Core.Configurations;
using AutomationTestingSample.Core.Constants;
using NUnit.Framework;
using System.Collections;

namespace AutomationTestingSample.Core
{
    public class TestFixtureDataSource
    {
        const int _hdWith = 1366;
        const int _hdHeight = 768;

        const int _fullhdWith = 1920;
        const int _fullhdHeight = 1080;

        const string _hd = "hd";
        const string _fullhd = "fullhd";

        public static IEnumerable Browsers
        {
            get
            {
                var browserSection = ConfigurationManager.GetSection<BrowserSection>(AppSettingConstants.Browser) ?? throw new Exception("You have to config browser in appsettings.");

                var browsers = browserSection.GetBrowsers();

                var resolutions = browserSection.GetResolutions();

                foreach (var browser in browsers)
                {
                    if (resolutions.Contains(_hd))
                    {
                        yield return new TestFixtureData(browser, _hdWith, _hdHeight);
                    }

                    if (resolutions.Contains(_fullhd))
                    {
                        yield return new TestFixtureData(browser, _fullhdWith, _fullhdHeight);
                    }
                }
            }
        }
    }
}
