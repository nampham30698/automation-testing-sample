using AutomationTestingSample.Core.Reports;
using AutomationTestingSample.Testing.Pages;
using AutomationTestingSample.Testing.Pages.Login;
using System.Net.Http.Headers;

namespace AutomationTestingSample.Testing.Tests
{
    internal class LoginPageTest : TestBase
    {
        public LoginPageTest(string browserType, int browserWidth, int browserHeight) : base(browserType, browserWidth, browserHeight)
        {
        }

        [TestCase]
        public void GetProductUrls()
        {
            var product = new ProductCalatlog(Driver);
            product.GetProductUrls();
        }
    }
}
