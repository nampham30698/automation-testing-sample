using AutomationTestingSample.Testing.Pages;

namespace AutomationTestingSample.Testing.Tests
{
    public class ProductCalatlog2Test : TestBase
    {
        public ProductCalatlog2Test(string browserType, int browserWidth, int browserHeight) : base(browserType, browserWidth, browserHeight)
        {
        }

        [TestCase]
        public void GetProductUrls()
        {
            var product = new ProductCalatlog2(Driver);
            product.GetProductUrls();
        }
    }
}
