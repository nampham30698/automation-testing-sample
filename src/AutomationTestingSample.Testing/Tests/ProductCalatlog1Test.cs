using AutomationTestingSample.Testing.Pages;

namespace AutomationTestingSample.Testing.Tests
{
    public class ProductCalatlog1Test : TestBase
    {
        public ProductCalatlog1Test(string browserType, int browserWidth, int browserHeight) : base(browserType, browserWidth, browserHeight)
        {
        }

        [TestCase]
        public void GetProductUrls()
        {
            var product = new ProductCalatlog1(Driver);
            product.GetProductUrls();
        }
    }
}
