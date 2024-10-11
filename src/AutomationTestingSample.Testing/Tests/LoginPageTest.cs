using AutomationTestingSample.Core.Reports;
using AutomationTestingSample.Testing.Pages.Login;

namespace AutomationTestingSample.Testing.Tests
{
    internal class LoginPageTest : TestBase
    {
        public LoginPageTest(string browserType, int browserWidth, int browserHeight) : base(browserType, browserWidth, browserHeight)
        {
        }

        [Test]
        public void TestLogin()
        {
            var loginPage = new LoginPage(Driver);

            loginPage.Login("", "");

            ExtentReporting.Instance.LogInfo("this is a test TestLogin");

            var a = false;

            Assert.IsTrue(a);
        }

        [TestCase("hooa", "hola")]
        public void LoginWithSSO(string h, string hfd)
        {
            var loginPage = new LoginPage(Driver);

            //loginPage.Login("", "");

            ExtentReporting.Instance.LogInfo("this is a test LoginWithSSO");

            var a = true;

            Assert.IsTrue(a);
        }
    }
}
