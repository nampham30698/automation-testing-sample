using AutomationTestingSample.Testing.Pages;

namespace AutomationTestingSample.Testing.Tests
{
    public class SingleProductPageTest : TestBase
    {
        public SingleProductPageTest(string browserType, int browserWidth, int browserHeight) : base(browserType, browserWidth, browserHeight)
        {
        }

        [TestCase]
        public void GetProductUrls()
        {
            var product = new SingleProductPage(Driver);
            product.GetProductUrls();
        }
    }
}
