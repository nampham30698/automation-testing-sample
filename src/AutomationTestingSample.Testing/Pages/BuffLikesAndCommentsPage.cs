using AutomationTestingSample.Core.Extensions;
using OpenQA.Selenium;

namespace AutomationTestingSample.Testing.Pages
{
    public class BuffLikesAndCommentsPage : WebPageBase
    {
        private const string _url = "https://leduyhiep.com/login";
        private const string _facebookServicesUrl = "https://leduyhiep.com/service/facebook/like-post";
        private const string _facebookCommentsUrl = "https://leduyhiep.com/service/facebook/comment";

        private const string _userNameActual = "";
        private const string _passwordActual = "";

        private IWebElement _userName => FindElement(By.XPath("//input[@id='username']"));

        private IWebElement _password => FindElement(By.XPath("//input[@id='password']"));

        private IWebElement _btnLogin => FindElement(By.XPath("//button[@class='btn btn-lg btn-primary btn-block']"));

        private IWebElement _inputPost => FindElement(By.XPath("//input[@id='object_id']"));

        private const string _seviceString = "//label[@class='custom-control-label server__name' and contains(@for,'{0}')]";

        private IWebElement _serviceSelected => FindElement(By.XPath("//label[@class='custom-control-label server__name' and contains(@for,'{}')]"));

        private IWebElement _inputQuality => FindElement(By.XPath("//input[@id='quantity']"));

        private IWebElement _btnCreateProcess => FindElement(By.XPath("//button[@id='btn-create']"));

        private IWebElement _btnConfirmation => FindElement(By.XPath("//button[@class='swal2-confirm swal2-styled']"));

        private IWebElement _textArea => FindElement(By.XPath("//textarea[@id='comments']"));

        private readonly List<LikePostOption> _likePosts =
        [
            //new(){PostId = "1703717253774645"     , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "386232357824186"      , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "379316815191082"      , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "1152211896079429"     , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "519277260561472"      , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "1021095966341962"     , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "1496169134621176"     , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "1857515754732326"     , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "3778232422392017"     , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "493063779980982"      , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "1045815300400518"     , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "499172146189537"      , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "523880836944574"      , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "823236063256203"      , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "1644755969640562"     , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "823356903241928"      , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "2152621815138189"     , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "1018558053338689"     , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "513222431217270"      , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "841610330978242"      , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "482839351121799"      , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "368165326119254"      , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "2205358206515657"     , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "8538765146151501"     , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "466473426152484"      , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "842250167845969"      , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "858915015671082"      , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "3705344403114702"     , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "1186444782594639"     , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "1946131319159760"     , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "1866294283858695"     , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "529170736131597"      , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "514132211021035"      , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "511729334583194"      , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "1518304959124579"     , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "1034897591600710"     , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "2842358402607540"     , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "524983626759651"      , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "1017498763201436"     , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "1019270386592473"     , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "1247656296267527"     , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "961797952645939"      , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "1327768558632774"      , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "920811729896326"      , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "2440156319521722"     , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "1253288069185853"     , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "1648276512410105"     , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "1063874518482170"     , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "1233683087767479"     , Quality = 50, Service = ServicesConstants.Sv5 },
            //new(){PostId = "1332627191056626"     , Quality = 50, Service = ServicesConstants.Sv5 },

            //new(){PostId = "1559339094690261"      , Quality = 10, Service = ServicesConstants.Sv2 },
            //new(){PostId = "1532203654083982"     , Quality = 10, Service = ServicesConstants.Sv2},
            //new(){PostId = "2385922731749820"     , Quality = 10, Service = ServicesConstants.Sv2 },
            //new(){PostId = "407440992302718"     , Quality = 10, Service = ServicesConstants.Sv2 },
            //new(){PostId = "3862765460719268"     , Quality = 10, Service = ServicesConstants.Sv2 },
            new(){PostId = "1443002343770419"     , Quality = 10, Service = ServicesConstants.Sv2 },
            new(){PostId = "1594460751443962"     , Quality = 10, Service = ServicesConstants.Sv2 },
            new(){PostId = "523365447076919"     , Quality = 10, Service = ServicesConstants.Sv2},
        ];


        public BuffLikesAndCommentsPage(IWebDriver driver) : base(driver)
        {

        }

        public void BuffLikes()
        {
            NavigateToUrl(_url);

            Login();

            NavigateToUrl(_facebookServicesUrl);

            EnterLikesService();
        }

        public void BuffComments()
        {
            NavigateToUrl(_url);

            Login();

            NavigateToUrl(_facebookCommentsUrl);

            EnterCommentsService();
        }


        public void Login()
        {
            _userName.EnterText(_userNameActual);
            _password.EnterText(_passwordActual);

            _btnLogin.Submit();
        }

        public void EnterLikesService()
        {
            foreach (var item in _likePosts)
            {
                _inputPost.EnterText(item.PostId);
                var selectOption = FindElement(By.XPath(string.Format(_seviceString, item.Service)));
                selectOption.Click();
                _inputQuality.EnterText(item.Quality.ToString());

                MoveToElement(_btnCreateProcess);
                _btnCreateProcess.Click();

                _btnConfirmation.Click();

                Thread.Sleep(7000);
                _btnConfirmation.Click();

                Thread.Sleep(1000);
                NavigateToUrl(_facebookServicesUrl);
            }
        }

        public void EnterCommentsService()
        {
            foreach (var item in _likePosts)
            {
                _inputPost.EnterText(item.PostId);
                var selectOption = FindElement(By.XPath(string.Format(_seviceString, item.Service)));
                selectOption.Click();

                _textArea.EnterText(@"Tranh bao nhiêu vậy bạn
Báo giá kích thước tranh
..
bn vậy
Giá tranh sao shop
ib gia
Ngang 60 giá bao nhiêu
xin giá shop
Tranh này bao nhiêu
Các mẫu tranh đẹp nhất");

                _inputQuality.EnterText(item.Quality.ToString());

                MoveToElement(_btnCreateProcess);
                _btnCreateProcess.Click();

                _btnConfirmation.Click();

                Thread.Sleep(7000);
                _btnConfirmation.Click();

                Thread.Sleep(1000);
                NavigateToUrl(_facebookCommentsUrl);
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
