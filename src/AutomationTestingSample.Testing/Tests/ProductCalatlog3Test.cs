using AutomationTestingSample.Testing.Pages;

namespace AutomationTestingSample.Testing.Tests
{
    public class ProductCalatlog3Test : TestBase
    {
        public ProductCalatlog3Test(string browserType, int browserWidth, int browserHeight) : base(browserType, browserWidth, browserHeight)
        {
        }

        [TestCase]
        public void GetProductUrls()
        {
            var product = new ProductCalatlog3(Driver);
            product.GetProductUrls();
        }
    }
}
