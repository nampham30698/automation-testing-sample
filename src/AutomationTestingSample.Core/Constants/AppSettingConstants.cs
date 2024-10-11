using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomationTestingSample.Core.Constants
{
    public static class AppSettingConstants
    {
        public const string Environment = "Environment";

        public const string BaseUrl = "Application:BaseUrl";

        public const string Browser = "Browser";

        public const string EnableHeadless = "Browser:EnableHeadless";
    }

    public class BrowserSection
    {
        public string Type { get; set; }

        public string Resolutions { get; set; }

        public string EnableHeadless { get; set; }

        public IEnumerable<string> GetBrowsers()
        {
            if(string.IsNullOrEmpty(Type)) throw new ArgumentNullException("type");
            return Type.Split(";");
        }

        public IEnumerable<string> GetResolutions()
        {
            if (string.IsNullOrEmpty(Type)) throw new ArgumentNullException("Resolutions");
            return Resolutions.Split(";");
        }
    }
}
