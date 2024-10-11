using OpenQA.Selenium;

namespace AutomationTestingSample.Testing.Pages.Login
{
    public class LoginPage : WebPageBase
    {
        private IWebElement TxtEmail => FindElement(By.XPath(""));

        private IWebElement TxtPassword => FindElement(By.XPath(""));

        private IWebElement BtnLogin => FindElement(By.XPath(""));

        public LoginPage(IWebDriver driver): base(driver)
        {

        }

        public void Login(string email, string password)
        {
            TxtEmail.SendKeys(email);
            TxtPassword.SendKeys(password);

            BtnLogin.Submit();
        }
    }
}
