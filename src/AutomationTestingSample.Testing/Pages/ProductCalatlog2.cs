using AutomationTestingSample.Core.Extensions;
using AutomationTestingSample.Core.Helpers;
using AutomationTestingSample.Core.Reports;
using AutomationTestingSample.Core.Utilities;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Security.Policy;
using static AutomationTestingSample.Testing.Pages.ProductCalatlog;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace AutomationTestingSample.Testing.Pages
{
    //https://ciaohaha.com/shop/NFL-shoes-ver-40.1/
    public class ProductCalatlog2 : WebPageBase
    {
        private ReadOnlyCollection<IWebElement> _singleProductLinks => FindElements(By.XPath("//div[contains(@class,'image-none')]//a"));
        private IWebElement _btnClosePopup => FindElement(By.XPath("//button[contains(@class,'klaviyo-close-form')]"));

        private IWebElement _productTitle => FindElement(By.XPath("//div[contains(@class,'product-main')]//h1"));
        //private IWebElement _productDescription => FindElement(By.XPath("//div[contains(@class,'woocommerce-Tabs-panel')]//div[contains(@class,'vtab_container')]"));
        private IWebElement _productDescription => FindElement(By.XPath(@"//*[@id='tab-description']"));
        private ReadOnlyCollection<IWebElement> _productPriceRange => FindElements(By.XPath("//div[contains(@class,'product-details-wrapper')]//p[contains(@class,'price')]//span[contains(@class,'woocommerce-Price-amount')]"));

        private IWebElement _productRegularPrice => FindElement(By.XPath("(//div[contains(@class,'product-details-wrapper')]//p[contains(@class,'price')]//span[contains(@class,'woocommerce-Price-amount')])[1]"));
        private IWebElement _productSalePrice => FindElement(By.XPath("(//div[contains(@class,'product-details-wrapper')]//p[contains(@class,'price')]//span[contains(@class,'woocommerce-Price-amount')])[2]"));

        private By _productVariationRegularPriceBy => By.XPath("(//p[contains(@class,'price product-page-price')]//bdi)[1]");

        private By _productVariationSalePriceBy => By.XPath("(//p[contains(@class,'price product-page-price')]//bdi)[2]");

        // buttom type
        private IWebElement _productButtomAttribute1Text => FindElement(By.XPath("(//div[contains(@class,'product__variant-label')])[1]//label"));
        private ReadOnlyCollection<IWebElement> _productButtomAttribute1 => FindElements(By.XPath("(//div[contains(@class,'product__variant-button')])[1]//button"));
        private ReadOnlyCollection<IWebElement> _productButtomAttribute2 => FindElements(By.XPath("(//div[contains(@class,'product__variant-button')])[2]//button"));

        
        private SelectElement _productSelectAttribute1 => FindSelectElement(By.XPath("//select[@data-attribute_name='attribute_size']"));
        private SelectElement _productSelectAttribute2 => FindSelectElement(By.XPath("//div[contains(@class,'product-details-wrapper')]//select[@data-attribute_name='attribute_pa_size']"));

        private By _productMainImageBy => By.XPath("(//div[contains(@class,'product-details-wrapper')]//div[contains(@class,'woocommerce-product-gallery__wrapper')]//img)[1]");
        private By _productImagesBy => By.XPath("//div[contains(@class,'flickity-slider')]//div[contains(@class,'woocommerce-product-gallery__image')]//img");

        public ProductCalatlog2(IWebDriver driver) : base(driver)
        {

        }

        public void GetProductUrls()
        {
            //_btnClosePopup.Click();

            Thread.Sleep(5000);

            //MoveToElement(FindElement(By.XPath("//*[@id='86H']/div/div/div[2]/div/div[2]/div/ul/li[5]/a")));

            //Thread.Sleep(3000);

            //MoveToElement(FindElement(By.XPath("//*[@id='86H']/div/div/div[2]/div/div[2]/div/ul/li[5]/a")));

            if (_singleProductLinks.Count == 0) return;

            var productMetadata = new List<ProductMetadata>();

            //productMetadata.Add(GetProductInfo("https://championssport.store/product/limited-edition-pittsburgh-penguins-grateful-dead-night-hoodie-set-v3nt17112416id10ds11/"));

            var links = _singleProductLinks.Select(x => x.GetAttribute("href")).ToList();

            //var links = new List<string>()
            //{
            //    "https://sportsboutiques.com/collections/nhl-men-s-thickened-corduroy-jacket/products/vitpl130?variant=1000015623838733"
            //};

            foreach (var link in links)
            {
                var product = GetProductInfo(link);
                if (product is not null)
                {
                    productMetadata.Add(product);
                }
            }

            ExportExel(productMetadata);
        }

        private ProductMetadata GetProductInfo(string url)
        {
            try
            {
                NavigateToUrl(url);
                Thread.Sleep(1000);

                var attribute1 = GetAttribute1Values();

                var data = new ProductMetadata()
                {
                    Url = url,
                    Title = _productTitle.Text,
                    //Description = _productDescription.GetAttribute("innerHTML").Trim(),
                    //RegularPrice = Parser.ParseDoube(_productRegularPrice.Text),
                    //SalePrice = Parser.ParseDoube(_productSalePrice.Text),
                    Attribute1 = attribute1,
                    Attribute2 = GetAttribute2Values(),
                    Images = GetImages(),
                    PriceMapping = GetPriceMapping(attribute1),
                };

                
                return data;
            }
            catch (Exception ex)
            {
                ExtentReporting.Instance.LogInfo("Fail at url: " + url);
                ExtentReporting.Instance.LogFail(ex);
                return null;
            }
        }

        private Dictionary<string, Tuple<double, double>> GetPriceMapping(List<string> attribute1 = null)
        {
            var priceMapping = new Dictionary<string, Tuple<double, double>>();

            foreach (var item in attribute1)
            {
                priceMapping.Add(item.Trim().ToLower(), new Tuple<double, double>(129.99, 69.99));
            }

            return priceMapping;
        }

        private List<string> GetAttribute1Values()
        {
            return _productSelectAttribute1.Options.Select(x => x.Text.Trim()).Where(x => !x.Contains("Choose an option", StringComparison.OrdinalIgnoreCase)).ToList();
        }

        private List<string> GetAttribute2Values()
        {
            return _productButtomAttribute2.Select(x => x.Text).ToList();
        }

        private List<string> GetImages()
        {
            var galleryImages = FindElements(_productImagesBy);
            if (galleryImages.Count > 0)
            {
                return galleryImages.Select(x => x.GetAttribute("data-src")).ToList();
            }
            return [FindElement(_productMainImageBy).GetAttribute("data-src")];
        }

        private void ExportExel(List<ProductMetadata> data)
        {
            var filePath = FileHelpers.ProjectPath + "woo_product_template.xlsx";

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using ExcelPackage excel = new(new FileInfo(filePath));

            var workSheet = excel.Workbook.Worksheets[0];

            int rowIndex = 2;

            int totalItemImported = 0;

            int skuParentNumber = 0;

            foreach (var item in data)
            {
                try
                {
                    skuParentNumber++;

                    var sku = $"{ProductConstants.SKU}-{skuParentNumber}";

                    #region main product

                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Id].Value = item.Url;
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Type].Value = ProductConstants.Variable;
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.SKU].Value = sku;
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Name].Value = item.Title;
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Published].Value = 1;
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.IsFeatured].Value = 0;
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.VisibilityINCatalog].Value = "visible";
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.ShortDescription].Value = "";
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Description].Value = ProductConstants.Description;
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.DateSalePriceStarts].Value = "";
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.DateSalePriceEnds].Value = "";
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.TaxStatus].Value = "taxable";
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.TaxClass].Value = "";
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.InStock].Value = 1;
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Stock].Value = "";
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.LowStock].Value = "";
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.BackOrdersAllowed].Value = 0;
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.SoldIndividually].Value = 0;
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Weight].Value = "";
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Lenght].Value = "";
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Width].Value = "";
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Height].Value = "";
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.AllowCustomerReviews].Value = 0;
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.PurchaseNote].Value = "";

                    workSheet.Cells[rowIndex, (int)WooExcelColumn.SalePrice].Value = "";
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.RegularPrice].Value = "";

                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Categories].Value = ProductConstants.Categories;
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Tags].Value = "";
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.ShippingClass].Value = "";
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Images].Value = string.Join(',', item.Images);
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.DownloadLimit].Value = "";
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.DownloadExpiryDays].Value = "";
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Parent].Value = "";
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.GroupedProduct].Value = "";
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.UpSells].Value = "";
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.CrossSells].Value = "";
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.ExternalUrl].Value = "";
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.ButtonText].Value = "";
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Position].Value = 0;

                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Attribute1Name].Value = ProductConstants.Attribute1Name;
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Attribute1Values].Value = string.Join(',', item.Attribute1);
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Attribute1Visible].Value = 1;
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Attribute1Global].Value = 0;

                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Attribute2Name].Value = ProductConstants.Attribute2Name;
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Attribute2Values].Value = string.Join(',', item.Attribute2);
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Attribute2Visible].Value = 1;
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Attribute2Global].Value = 0;

                    #endregion

                    var rowChild = rowIndex + 1;
                    int variationPosition = 0;

                    foreach (var atribute1 in item.Attribute1)
                    {
                        var (regularPrice, salePrice) = item.PriceMapping[atribute1.Trim().ToLower()];

                        variationPosition++;

                        #region variation

                        workSheet.Cells[rowChild, (int)WooExcelColumn.Id].Value = item.Url;
                        workSheet.Cells[rowChild, (int)WooExcelColumn.Type].Value = ProductConstants.variation;
                        workSheet.Cells[rowChild, (int)WooExcelColumn.SKU].Value = "";
                        workSheet.Cells[rowChild, (int)WooExcelColumn.Name].Value = item.Title;
                        workSheet.Cells[rowChild, (int)WooExcelColumn.Published].Value = 1;
                        workSheet.Cells[rowChild, (int)WooExcelColumn.IsFeatured].Value = 0;
                        workSheet.Cells[rowChild, (int)WooExcelColumn.VisibilityINCatalog].Value = "visible";
                        workSheet.Cells[rowChild, (int)WooExcelColumn.ShortDescription].Value = "";
                        workSheet.Cells[rowChild, (int)WooExcelColumn.Description].Value = "";
                        workSheet.Cells[rowChild, (int)WooExcelColumn.DateSalePriceStarts].Value = "";
                        workSheet.Cells[rowChild, (int)WooExcelColumn.DateSalePriceEnds].Value = "";
                        workSheet.Cells[rowChild, (int)WooExcelColumn.TaxStatus].Value = "taxable";
                        workSheet.Cells[rowChild, (int)WooExcelColumn.TaxClass].Value = "";
                        workSheet.Cells[rowChild, (int)WooExcelColumn.InStock].Value = 1;
                        workSheet.Cells[rowChild, (int)WooExcelColumn.Stock].Value = "";
                        workSheet.Cells[rowChild, (int)WooExcelColumn.LowStock].Value = "";
                        workSheet.Cells[rowChild, (int)WooExcelColumn.BackOrdersAllowed].Value = 0;
                        workSheet.Cells[rowChild, (int)WooExcelColumn.SoldIndividually].Value = 0;
                        workSheet.Cells[rowChild, (int)WooExcelColumn.Weight].Value = "";
                        workSheet.Cells[rowChild, (int)WooExcelColumn.Lenght].Value = "";
                        workSheet.Cells[rowChild, (int)WooExcelColumn.Width].Value = "";
                        workSheet.Cells[rowChild, (int)WooExcelColumn.Height].Value = "";
                        workSheet.Cells[rowChild, (int)WooExcelColumn.AllowCustomerReviews].Value = 0;
                        workSheet.Cells[rowChild, (int)WooExcelColumn.PurchaseNote].Value = "";

                        workSheet.Cells[rowChild, (int)WooExcelColumn.SalePrice].Value = salePrice;
                        workSheet.Cells[rowChild, (int)WooExcelColumn.RegularPrice].Value = regularPrice;

                        workSheet.Cells[rowChild, (int)WooExcelColumn.Categories].Value = "";
                        workSheet.Cells[rowChild, (int)WooExcelColumn.Tags].Value = "";
                        workSheet.Cells[rowChild, (int)WooExcelColumn.ShippingClass].Value = "";
                        workSheet.Cells[rowChild, (int)WooExcelColumn.Images].Value = "";
                        workSheet.Cells[rowChild, (int)WooExcelColumn.DownloadLimit].Value = "";
                        workSheet.Cells[rowChild, (int)WooExcelColumn.DownloadExpiryDays].Value = "";
                        workSheet.Cells[rowChild, (int)WooExcelColumn.Parent].Value = sku;
                        workSheet.Cells[rowChild, (int)WooExcelColumn.GroupedProduct].Value = "";
                        workSheet.Cells[rowChild, (int)WooExcelColumn.UpSells].Value = "";
                        workSheet.Cells[rowChild, (int)WooExcelColumn.CrossSells].Value = "";
                        workSheet.Cells[rowChild, (int)WooExcelColumn.ExternalUrl].Value = "";
                        workSheet.Cells[rowChild, (int)WooExcelColumn.ButtonText].Value = "";
                        workSheet.Cells[rowChild, (int)WooExcelColumn.Position].Value = variationPosition;

                        workSheet.Cells[rowChild, (int)WooExcelColumn.Attribute1Name].Value = ProductConstants.Attribute1Name;
                        workSheet.Cells[rowChild, (int)WooExcelColumn.Attribute1Values].Value = atribute1;
                        workSheet.Cells[rowChild, (int)WooExcelColumn.Attribute1Visible].Value = "";
                        workSheet.Cells[rowChild, (int)WooExcelColumn.Attribute1Global].Value = 0;

                        #endregion

                        rowChild++;
                    }

                    rowIndex = rowChild;

                    totalItemImported++;
                }
                catch (Exception ex)
                {
                    ExtentReporting.Instance.LogInfo("Fail at url: " + item.Url);
                    ExtentReporting.Instance.LogFail(ex);
                }
            }

            ExtentReporting.Instance.LogInfo("Total imported products: " + totalItemImported);

            SaveWorksheetAsCsv(workSheet, FileHelpers.ProjectPath + "woo_products_import_" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".csv");
        }

        static void SaveWorksheetAsCsv(ExcelWorksheet worksheet, string csvFilePath)
        {
            using (var writer = new StreamWriter(csvFilePath))
            {
                // Iterate through rows and columns
                for (int row = worksheet.Dimension.Start.Row; row <= worksheet.Dimension.End.Row; row++)
                {
                    var rowValues = new List<string>();
                    for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++)
                    {
                        var cellValue = worksheet.Cells[row, col].Text; // Get the cell value as text
                        rowValues.Add(cellValue.Replace("\"", "\"\"")); // Escape quotes for CSV
                    }

                    // Write the row to the CSV
                    writer.WriteLine(string.Join(",", rowValues.Select(value => $"\"{value}\"")));
                }
            }
        }

        public class ProductMetadata
        {
            public string Url { get; set; }
            public string Title { get; set; }
            public string Description { get; set; }
            public double RegularPrice { get; set; }
            public double SalePrice { get; set; }
            public List<string> Attribute1 { get; set; } = new();
            public List<string> Attribute2 { get; set; } = new();
            public List<string> Images { get; set; } = new();
            public Dictionary<string, Tuple<double, double>> PriceMapping { get; set; } = new();
        }

        public enum WooExcelColumn
        {
            Id = 1,
            Type,
            SKU,
            Name,
            Published,
            IsFeatured,
            VisibilityINCatalog,
            ShortDescription,
            Description,
            DateSalePriceStarts,
            DateSalePriceEnds,
            TaxStatus,
            TaxClass,
            InStock,
            Stock,
            LowStock,
            BackOrdersAllowed,
            SoldIndividually,
            Weight,
            Lenght,
            Width,
            Height,
            AllowCustomerReviews,
            PurchaseNote,
            SalePrice,
            RegularPrice,
            Categories,
            Tags,
            ShippingClass,
            Images,
            DownloadLimit,
            DownloadExpiryDays,
            Parent,
            GroupedProduct,
            UpSells,
            CrossSells,
            ExternalUrl,
            ButtonText,
            Position,
            Attribute1Name,
            Attribute1Values,
            Attribute1Visible,
            Attribute1Global,
            Attribute2Name,
            Attribute2Values,
            Attribute2Visible,
            Attribute2Global,
        }

        public static class ProductConstants
        {
            public const string Simple = "simple"; 
            public const string Variable = "variable";
            public const string variation = "variation";

            public const string Description = @"Customize Your Name With  Arizona Cardinals Ver 40.1 Sport Shoes

Hi there!! Are you looking at our  Arizona Cardinals shoes/sneakers. These sneakers are unique, made with love, and they are perfect accessories for your outfit. The Arizona Cardinals  Emblem on the shoes makes it a perfect item for any Arizona Cardinals Lovers.

With the launch of the exclusive Made-to-Order service, you can make your sneakers/shoes unique by not only having your special Arizona Cardinals Emblem but also your Name printed on the shoes.";

            public const string SKU = "nfl-shoes-ver-40-1";
            public const string Categories = "NFL shoes ver 40.1";

            public const string Attribute1Name = "Size";
            public const string Attribute2Name = "Size";

            public static VariationType VariationType = VariationType.Buttom;
        }

        public enum VariationType
        {
            None,
            Buttom,
            Select
        }
    }
}
