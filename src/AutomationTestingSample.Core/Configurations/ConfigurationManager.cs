using AutomationTestingSample.Core.Constants;
using AutomationTestingSample.Core.Helpers;
using Microsoft.Extensions.Configuration;

namespace AutomationTestingSample.Core.Configurations
{
    public class ConfigurationManager
    {
        static IConfigurationRoot _configurationRoot;

        public static void Configure() {

            var appSettingConfig = new ConfigurationBuilder()
                                 .SetBasePath(FileHelpers.ProjectPath)
                                 .AddJsonFile("appsettings.json", true).Build();


            _configurationRoot = new ConfigurationBuilder()
                                .SetBasePath(FileHelpers.ProjectPath)
                               .AddJsonFile("appsettings.json", true)
                               .AddJsonFile($"appsettings.{appSettingConfig[AppSettingConstants.Environment]}.json", true)
                               .AddEnvironmentVariables().Build();
        }

        public static T GetSection<T>(string key) where T : class
        {
            if (_configurationRoot is null)
                Configure();

            return _configurationRoot.GetSection(key).Get<T>();
        }

        public static T GetValue<T>(string key)
        {
            if (_configurationRoot is null)
                Configure();

            return _configurationRoot.GetValue<T>(key);
        }
    }
}
