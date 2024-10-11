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
            var b = new List<int>() { 1, 2, 3, 5 };
            var bc = new List<int>() { 1, 2, 3, 5 };

            Assert.IsTrue(a);
            Assert.That(b, Is.SubsetOf(bc));
        }

        [TestCase("hooa", "hola")]
        public void LoginWithSSO(string h, string hfd)
        {
            var loginPage = new LoginPage(Driver);

            var b = new List<int>() { 1, 2 };
            var bc = new List<int>() { 1, 2, 3, 5 };

            Assert.That(b, Is.SubsetOf(bc));
            Is.AnyOf(bc);
            Assert.That(5, Is.AnyOf(1,5,4));
        }
    }
}
