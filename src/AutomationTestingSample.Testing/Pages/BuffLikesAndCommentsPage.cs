using AutomationTestingSample.Core.Extensions;
using AutomationTestingSample.Core.Reports;
using OpenQA.Selenium;

namespace AutomationTestingSample.Testing.Pages
{
    public class BuffLikesAndCommentsPage : WebPageBase
    {
        private const string _url = "https://leduyhiep.com/login";
        private const string _facebookServicesUrl = "https://leduyhiep.com/service/facebook/like-post";
        private const string _facebookCommentsUrl = "https://leduyhiep.com/service/facebook/comment";

        private const string _userNameActual = "hoatran68";
        private const string _passwordActual = "Hoatran123$";

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

        private Random _random = new Random();

        private readonly List<LikePostOption> _likePosts =
        [
            new(){PostId = "1780989559338402"  , Quality = new Random().Next(50,55), Service = ServicesConstants.Sv5 },
            new(){PostId = "3410431132597098"  , Quality = new Random().Next(50,55), Service = ServicesConstants.Sv5 },
            new(){PostId = "549638111289623"   , Quality = new Random().Next(50,55), Service = ServicesConstants.Sv5 },
            new(){PostId = "612417311349789"   , Quality = new Random().Next(50,55), Service = ServicesConstants.Sv5 },
            new(){PostId = "983024497185820"   , Quality = new Random().Next(50,55), Service = ServicesConstants.Sv5 },
            new(){PostId = "597706012916558"   , Quality = new Random().Next(50,55), Service = ServicesConstants.Sv5 },
            new(){PostId = "1745787476209628"  , Quality = new Random().Next(50,55), Service = ServicesConstants.Sv5 },
            new(){PostId = "1131892861631268"  , Quality = new Random().Next(50,55), Service = ServicesConstants.Sv5 },
            new(){PostId = "9263390440359374"  , Quality = new Random().Next(50,55), Service = ServicesConstants.Sv5 },
            new(){PostId = "973969041432276"   , Quality = new Random().Next(50,55), Service = ServicesConstants.Sv5 },
            new(){PostId = "1205890790974545"  , Quality = new Random().Next(50,55), Service = ServicesConstants.Sv5 },
            new(){PostId = "974518431176855"   , Quality = new Random().Next(50,55), Service = ServicesConstants.Sv5 },
            new(){PostId = "1243052853654409"  , Quality = new Random().Next(50,55), Service = ServicesConstants.Sv5 },
            new(){PostId = "1326899824960546"  , Quality = new Random().Next(50,55), Service = ServicesConstants.Sv5 },
        ];

        private readonly List<LikePostOption> _commentsData = 
        [
             new(){PostId = "1105186230602198" , Quality = 10, Service = ServicesConstants.Sv2 },
             new(){PostId = "1279199653211943" , Quality = 10, Service = ServicesConstants.Sv2 },
             new(){PostId = "3466176840342628" , Quality = 10, Service = ServicesConstants.Sv2 },
             new(){PostId = "566130932680418"  , Quality = 10, Service = ServicesConstants.Sv2 },
             new(){PostId = "1223877022030702" , Quality = 10, Service = ServicesConstants.Sv2 },
             new(){PostId = "575361315084258"  , Quality = 10, Service = ServicesConstants.Sv2 },
             new(){PostId = "1237506710863240" , Quality = 10, Service = ServicesConstants.Sv2 },
             new(){PostId = "1098307835219198" , Quality = 10, Service = ServicesConstants.Sv2 },
             new(){PostId = "465901132771132"  , Quality = 10, Service = ServicesConstants.Sv2 },
             new(){PostId = "1153453742864493" , Quality = 10, Service = ServicesConstants.Sv2 },
             new(){PostId = "444782242009722"  , Quality = 10, Service = ServicesConstants.Sv2 },
             new(){PostId = "8782369408483459" , Quality = 10, Service = ServicesConstants.Sv2 },
             new(){PostId = "1102842228006189" , Quality = 10, Service = ServicesConstants.Sv2 },
             new(){PostId = "890686319507254"  , Quality = 10, Service = ServicesConstants.Sv2 },
             new(){PostId = "503014379433673"  , Quality = 10, Service = ServicesConstants.Sv2 },
             new(){PostId = "577808677967936"  , Quality = 10, Service = ServicesConstants.Sv2 },
             new(){PostId = "2252422141794134" , Quality = 10, Service = ServicesConstants.Sv2 },
             new(){PostId = "1280408406629216" , Quality = 10, Service = ServicesConstants.Sv2 },
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
                try
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
                catch
                {
                    ExtentReporting.Instance.LogInfo("Fail at id: " + item.PostId);
                }
            }
        }

        
        public void EnterCommentsService()
        {
            foreach (var item in _commentsData)
            {
                try
                {
                    _inputPost.EnterText(item.PostId);
                    var selectOption = FindElement(By.XPath(string.Format(_seviceString, item.Service)));
                    selectOption.Click();

                    _textArea.EnterText(@"xin giá sỉ
..
Bn vậy shop
15 chiếc giá sao
ib gia
Xin giá shop
ib ạ
bn ạ
check ib
mẫu này bao nhiêu");

                    _inputQuality.EnterText(item.Quality.ToString());

                    MoveToElement(_btnCreateProcess);
                    _btnCreateProcess.Click();

                    _btnConfirmation.Click();

                    Thread.Sleep(7000);
                    _btnConfirmation.Click();

                    Thread.Sleep(1000);
                    NavigateToUrl(_facebookCommentsUrl);
                }
                catch
                {
                    ExtentReporting.Instance.LogInfo("Fail at url: " + item.PostId);
                }
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
