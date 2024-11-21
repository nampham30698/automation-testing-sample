using OfficeOpenXml;
using OpenQA.Selenium;
using System;
using System.Collections.ObjectModel;

namespace AutomationTestingSample.Testing.Pages
{
    public class ProductCalatlog : WebPageBase
    {
        private ReadOnlyCollection<IWebElement> _singleProductLinks => FindElements(By.XPath("//div[contains(@class,'woocommerce-image__wrapper')]//a"));
        private IWebElement _btnClosePopup => FindElement(By.XPath("//button[contains(@class,'klaviyo-close-form')]"));

        private IWebElement _productTitle => FindElement(By.XPath("//div[contains(@class,'product-details-wrapper')]//h1[contains(@class,'product_title')]"));
        private ReadOnlyCollection<IWebElement> _productPriceRange => FindElements(By.XPath("//div[contains(@class,'product-details-wrapper')]//p[contains(@class,'price')]//span[contains(@class,'woocommerce-Price-amount')]"));

        private IWebElement _productRegularPrice => FindElement(By.XPath("(//div[contains(@class,'product-details-wrapper')]//p[contains(@class,'price')]//span[contains(@class,'woocommerce-Price-amount')])[1]"));
        private IWebElement _productSalePrice => FindElement(By.XPath("(//div[contains(@class,'product-details-wrapper')]//p[contains(@class,'price')]//span[contains(@class,'woocommerce-Price-amount')])[2]"));

        private ReadOnlyCollection<IWebElement> _productStyles => FindElements(By.XPath("//div[contains(@class,'product-details-wrapper')]//ul[@data-attribute='attribute_style']//li[contains(@class,'cgkit-attribute-swatch')]"));
        private ReadOnlyCollection<IWebElement> _productSizes => FindElements(By.XPath("//div[contains(@class,'product-details-wrapper')]//ul[@data-attribute='attribute_size']//li[contains(@class,'cgkit-attribute-swatch')]"));
        private ReadOnlyCollection<IWebElement> _productImages => FindElements(By.XPath("//div[contains(@class,'product-details-wrapper')]//ol[contains(@class,'flex-control-thumbs')]//img"));
       

        public ProductCalatlog(IWebDriver driver) : base(driver)
        {

        }

        public void GetProductUrls()
        {
            Thread.Sleep(2000);

            _btnClosePopup.Click();

            Thread.Sleep(2000);

            //GetProductInfo("https://btfluxury.com/product/uconn-mens-basketball-six-time-ncaa-champion-%F0%9F%8F%86%F0%9F%8F%86%F0%9F%8F%86%F0%9F%8F%86%F0%9F%8F%86%F0%9F%8F%86-coachs-hoodie/");
            
            if (_singleProductLinks.Count != 0)
            {
                var links = _singleProductLinks.Select(x => x.GetAttribute("href")).ToList();

                var i = 0;
                foreach (var link in links)
                {
                    if (i == 3) return;
                    GetProductInfo(link);
                    i++;
                }
            }
        }

        private void GetProductInfo(string url)
        {
            NavigateToUrl(url);
            Thread.Sleep(2000);
            Console.WriteLine(_productTitle.Text);
            Console.WriteLine(_productRegularPrice.Text);
            Console.WriteLine(_productSalePrice.Text);

            foreach (var item in _productStyles)
            {
                Console.WriteLine(item.Text);
            }

            foreach (var item in _productSizes)
            {
                Console.WriteLine(item.Text);
            }

            foreach (var item in _productImages)
            {
                Console.WriteLine(item.GetAttribute("src"));
            }
        }

        private void ExportExel()
        {
            ExcelPackage excel = new ExcelPackage("");

            // name of the sheet 
            var workSheet = excel.Workbook.Worksheets.Add("Sheet1");
        }
    }
}
