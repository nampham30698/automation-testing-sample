using AutomationTestingSample.Core.Helpers;
using AutomationTestingSample.Core.Reports;
using OfficeOpenXml;
using OpenQA.Selenium;
using System;
using System.Collections.ObjectModel;

namespace AutomationTestingSample.Testing.Pages
{
    // url https://gotyourstyle.com/product-category/clothings/hoodie/nfl-fishing-collection-2024
    public class ProductCalatlogRadioOneVariation : WebPageBase
    {
        private ReadOnlyCollection<IWebElement> _singleProductLinks => FindElements(By.XPath("//div[contains(@class,'woocommerce-image__wrapper')]//a"));

        private IWebElement _btnClosePopup => FindElement(By.XPath("//button[contains(@class,'klaviyo-close-form')]"));

        private IWebElement _productTitle => FindElement(By.XPath("//section[contains(@class,'container-page')]//div[@id='detail-contents']//h1[contains(@class,'product__name-product')]"));

        private IWebElement _productDescription => FindElement(By.XPath(@"//*[@id='tab-description']"));
        private ReadOnlyCollection<IWebElement> _productPriceRange => FindElements(By.XPath("//main[contains(@class,'site-main')]//p[contains(@class,'price')]//span[contains(@class,'woocommerce-Price-amount')]"));

        private IWebElement _productRegularPrice => FindElement(By.XPath("(//main[contains(@class,'site-main')]//p[contains(@class,'price')]//span[contains(@class,'woocommerce-Price-amount')])[1]"));
        private IWebElement _productSalePrice => FindElement(By.XPath("(//main[contains(@class,'site-main')]//p[contains(@class,'price')]//span[contains(@class,'woocommerce-Price-amount')])[2]"));

        // label type
        private ReadOnlyCollection<IWebElement> _productButtomAttribute1 => FindElements(By.XPath("//main[contains(@class,'site-main')]//div[contains(@data-validation,'Choose Hoodie Type')]//label"));
        private ReadOnlyCollection<IWebElement> _productButtomAttribute2 => FindElements(By.XPath("//main[contains(@class,'site-main')]//div[contains(@data-validation,'Choose Your Size')]//label"));

        private By _productMainImageBy => By.XPath("//section[contains(@class,'container-page')]//div[@id='product-image-carousel']//img");
        private By _productImagesBy => By.XPath("//section[contains(@class,'container-page')]//ul[contains(@class,'media-gallery-carousel__thumbs')]//img");

        public ProductCalatlogRadioOneVariation(IWebDriver driver) : base(driver)
        {

        }

        public void GetProductUrls()
        {
            //_btnClosePopup.Click();

            //if (_singleProductLinks.Count == 0) return;

            var productMetadata = new List<ProductMetadata>();

            //productMetadata.Add(GetProductInfo("https://championssport.store/product/limited-edition-pittsburgh-penguins-grateful-dead-night-hoodie-set-v3nt17112416id10ds11/"));

            //var links = _singleProductLinks.Select(x => x.GetAttribute("href")).ToList();

            var links = new List<string>()
            {
                 "https://tvtcloset.com/collections/wsh-orm24/products/orm-wsh-tt-ha-24111401",
                "https://tvtcloset.com/collections/wsh-orm24/products/orm-wsh-tt-ha-24111402",
                "https://tvtcloset.com/collections/wsh-orm24/products/orm-wsh-tt-ha-24111403",
                "https://tvtcloset.com/collections/wsh-orm24/products/orm-wsh-tt-ha-24111404",
                "https://tvtcloset.com/collections/wsh-orm24/products/orm-wsh-tt-ha-24111405",
                "https://tvtcloset.com/collections/wsh-orm24/products/orm-wsh-tt-ha-24111406",
                "https://tvtcloset.com/collections/wsh-orm24/products/orm-wsh-tt-ha-24111407",
                "https://tvtcloset.com/collections/wsh-orm24/products/orm-wsh-tt-ha-24111408",
                "https://tvtcloset.com/collections/wsh-orm24/products/orm-wsh-tt-ha-24111409"

            };

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

                var data = new ProductMetadata()
                {
                    Url = url,
                    Title = _productTitle.Text,
                    //Description = _productDescription.GetAttribute("innerHTML").Trim(),
                    //RegularPrice = Parser.ParseDoube(_productRegularPrice.Text),
                    //SalePrice = Parser.ParseDoube(_productSalePrice.Text),
                    Attribute1 = GetAttribute1Values(),
                    Attribute2 = GetAttribute2Values(),
                    Images = GetImages(),
                    PriceMapping = GetPriceMapping(),
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

        private Dictionary<string, Tuple<double, double>> GetPriceMapping()
        {
            var priceMapping = new Dictionary<string, Tuple<double, double>>();

            var _productButtomAttribute1 = new Dictionary<string, Tuple<double, double>>()
            {
                {"Pack 1", new Tuple<double, double>(19.95,14.95) },
                {"Pack 3", new Tuple<double, double>(44.95,34.95) },
                {"Pack 5", new Tuple<double, double>(79.95,44.95) }
            };

            foreach (var item in _productButtomAttribute1)
            {
                priceMapping.Add(item.Key.Trim().ToLower(), item.Value);
            }

            return priceMapping;
        }

        private List<string> GetAttribute1Values()
        {
            return
            [
                "Pack 1",
                "Pack 3",
                "Pack 5"
            ];
        }

        private List<string> GetAttribute2Values()
        {
            return
            [
                "S",
                "M",
                "L",
                "XL",
                "2XL",
                "3XL",
                "4XL",
                "5XL",
            ];
        }

        private List<string> GetImages()
        {
            //var galleryImages = FindElements(_productImagesBy);
            //if (galleryImages.Count > 0)
            //{
            //    return galleryImages.Select(x => x.GetAttribute("src")).ToList();
            //}
            return [FindElement(_productMainImageBy).GetAttribute("src")];
        }

        private void ExportExel(List<ProductMetadata> data)
        {
            var filePath = FileHelpers.ProjectPath + "woo_product_template_one_variation.xlsx";

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
            Attribute3Name,
            Attribute3Values,
            Attribute3Visible,
            Attribute3Global,
        }

        public static class ProductConstants
        {
            public const string Simple = "simple";
            public const string Variable = "variable";
            public const string variation = "variation";

            public const string Description = @"This ornament is perfect for your Christmas tree, car rear view mirror, house, birthday gift, or gift for any holiday. Try something new for your Christmas Decoration concept.

<strong>Product Details</strong>
<ul>
 	<li>DOUBLE sided dye-sublimation printing.
Size: 3.3 inches tall</li>
 	<li>Made of lightweight mica acrylic, easy for hanging.</li>
 	<li>Nicely sized for easy coordination with other small or large-sized ornaments.</li>
 	<li>Comes ready with an attached loop for immediate hanging on your rear view mirror.</li>
 	<li>Perfect for Christmas tree decoration,car rear view mirror, wreaths, doorways, and more. Due to the difference monitor and light effect, the actual color and size of the item may be slightly difference from the visual image.</li>
</ul>";

            public const string SKU = "wsh-orm24";
            public const string Categories = "WSH ORM24";

            public const string Attribute1Name = "Pack";

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
