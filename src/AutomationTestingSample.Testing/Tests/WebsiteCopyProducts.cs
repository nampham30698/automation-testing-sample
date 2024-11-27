using AutomationTestingSample.Testing.Pages;

namespace AutomationTestingSample.Testing.Tests
{
    public class WebsiteCopyProducts : TestBase
    {
        public WebsiteCopyProducts(string browserType, int browserWidth, int browserHeight) : base(browserType, browserWidth, browserHeight)
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
