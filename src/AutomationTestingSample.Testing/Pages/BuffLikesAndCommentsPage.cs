using AutomationTestingSample.Core.Extensions;
using OpenQA.Selenium;

namespace AutomationTestingSample.Testing.Pages
{
    public class BuffLikesAndCommentsPage : WebPageBase
    {
        private const string _url = "https://leduyhiep.com/login";
        private const string _facebookServicesUrl = "https://leduyhiep.com/service/facebook/like-post";

        private const string _userNameActual = "";
        private const string _passwordActual = "";

        private IWebElement _userName => FindElement(By.XPath("//input[@id='username']"));

        private IWebElement _password => FindElement(By.XPath("//input[@id='password']"));

        private IWebElement _btnLogin => FindElement(By.XPath("//button[@class='btn btn-lg btn-primary btn-block']"));

        private IWebElement _inputPost => FindElement(By.XPath("//input[@id='object_id']"));

        private const string _seviceString = "//label[@class='custom-control-label server__name' and contains(@for,'{0}')]";

        private IWebElement _serviceSelected => FindElement(By.XPath("//label[@class='custom-control-label server__name' and contains(@for,'{}')]"));

        private IWebElement _inputQuality => FindElement(By.XPath("//input[@id='quantity']"));

        private IWebElement _btnCreateProcess => FindElement(By.XPath("//button[@class='btn btn-success btn-block']"));

        private readonly List<LikePostOption> _likePosts =
        [
            new(){PostId = "1559339094690261", Quality = 50, Service = ServicesConstants.Sv5 },
        ];


        public BuffLikesAndCommentsPage(IWebDriver driver) : base(driver)
        {
            
        }

        public void BuffLikes()
        {
            NavigateToUrl(_url);

            Login();

            NavigateToUrl(_facebookServicesUrl);

            EnterService();
        }


        public void Login()
        {
            _userName.EnterText(_userNameActual);
            _password.EnterText(_passwordActual);

            _btnLogin.Submit();
        }

        public void EnterService()
        {
            foreach (var item in _likePosts)
            {
                _inputPost.EnterText(item.PostId);
                var selectOption = FindElement(By.XPath(string.Format(_seviceString, item.Service)));
                selectOption.Click();
                _inputQuality.EnterText(item.Quality.ToString());
            }
        }

        private class LikePostOption
        {
            public string PostId { get; set; }

            public int Quality { get; set; }

            public string Service { get; set; } = ServicesConstants.Sv5;
        }

        public static class ServicesConstants
        {
            public const string Sv1 = "sv1";
            public const string Sv2 = "sv2";
            public const string Sv3 = "sv3";
            public const string Sv4 = "sv4";
            public const string Sv5 = "sv5";
            public const string Sv6 = "sv6";
            public const string Sv7 = "sv7";
            public const string Sv8 = "sv8";
            public const string Sv9 = "sv9";
            public const string Sv10 = "sv10";
            public const string Sv11 = "sv11";
            public const string Sv12 = "sv12";
            public const string Sv13 = "sv13";
            public const string Sv14 = "sv14";
            public const string Sv15 = "sv15";
            public const string Sv16 = "sv16";
            public const string Sv17 = "sv17";
            public const string Sv18 = "sv18";
            public const string Sv19 = "sv19";
            public const string Sv20 = "sv20";
            public const string Sv21 = "sv21";
            public const string Sv22 = "sv22";
            public const string Sv23 = "sv23";
            public const string Sv24 = "sv24";
            public const string Sv25 = "sv25";
        }
    }
}
