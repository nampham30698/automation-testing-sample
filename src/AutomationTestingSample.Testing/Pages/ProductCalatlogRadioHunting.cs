using AutomationTestingSample.Core.Helpers;
using AutomationTestingSample.Core.Reports;
using OfficeOpenXml;
using OpenQA.Selenium;
using System.Collections.ObjectModel;

namespace AutomationTestingSample.Testing.Pages
{
    // url https://temose.com/category/recv5tvbvc
    public class ProductCalatlogRadioHunting : WebPageBase
    {
        private ReadOnlyCollection<IWebElement> _singleProductLinks => FindElements(By.XPath("//div[@class='text-center rounded mb-7']//a"));

        private IWebElement _btnClosePopup => FindElement(By.XPath("//button[contains(@class,'klaviyo-close-form')]"));

        private IWebElement _productTitle => FindElement(By.XPath("//*[@id='container']/div/div/div[3]/div/div/div/div[2]/h2[1]"));

        private IWebElement _productDescription => FindElement(By.XPath(@"//*[@id='tab-description']"));
        private ReadOnlyCollection<IWebElement> _productPriceRange => FindElements(By.XPath("//main[contains(@class,'site-main')]//p[contains(@class,'price')]//span[contains(@class,'woocommerce-Price-amount')]"));

        private IWebElement _productRegularPrice => FindElement(By.XPath("(//main[contains(@class,'site-main')]//p[contains(@class,'price')]//span[contains(@class,'woocommerce-Price-amount')])[1]"));
        private IWebElement _productSalePrice => FindElement(By.XPath("(//main[contains(@class,'site-main')]//p[contains(@class,'price')]//span[contains(@class,'woocommerce-Price-amount')])[2]"));

        // label type
        private ReadOnlyCollection<IWebElement> _productButtomAttribute1 => FindElements(By.XPath("//main[contains(@class,'site-main')]//div[contains(@data-validation,'Choose Hoodie Type')]//label"));
        private ReadOnlyCollection<IWebElement> _productButtomAttribute2 => FindElements(By.XPath("//main[contains(@class,'site-main')]//div[contains(@data-validation,'Choose Your Size')]//label"));

        private By _productMainImageBy => By.XPath("//*[@id='container']/div/div/div[3]/div/div/div/div[2]/div[2]/a[1]/div/img");
        private By _productImagesBy => By.XPath("//*[@id='container']/div/div/div[3]/div/div/div/div[2]/div[2]/a[2]/div/img");

        public ProductCalatlogRadioHunting(IWebDriver driver) : base(driver)
        {

        }

        public void GetProductUrls()
        {
            //_btnClosePopup.Click();

            if (_singleProductLinks.Count == 0) return;

            foreach (var item in _singleProductLinks)
            {
                Console.WriteLine(item.GetAttribute("href"));
            }

            var productMetadata = new List<ProductMetadata>();

            //productMetadata.Add(GetProductInfo("https://championssport.store/product/limited-edition-pittsburgh-penguins-grateful-dead-night-hoodie-set-v3nt17112416id10ds11/"));

            //var links = _singleProductLinks.Select(x => x.GetAttribute("href")).ToList();

            var links = new List<string>()
            {
                "https://temose.com/campaign/p546bv",
                "https://temose.com/campaign/phil35g6v",
                "https://temose.com/campaign/4rfgf5g",
                "https://temose.com/campaign/4gfg56g",
                "https://temose.com/campaign/buff45fg5",
                "https://temose.com/campaign/cle45ct6fbg",
                "https://temose.com/campaign/la4fgcvyt",
                "https://temose.com/campaign/ka564vbvb",
                "https://temose.com/campaign/ne45gvcg",
                "https://temose.com/campaign/435cgcgbb",
                "https://temose.com/campaign/te4gfgfg",
                "https://temose.com/campaign/lo45gfg",
                "https://temose.com/campaign/ne3vcvc",
                "https://temose.com/campaign/4ggfgfg",
                "https://temose.com/campaign/spo6c5v6",
                "https://temose.com/campaign/tenn4dfg",
                "https://temose.com/campaign/jac4fgh6",
                "https://temose.com/campaign/ind45ghgy",
                "https://temose.com/campaign/3cv45vv",
                "https://temose.com/campaign/atla345dffg",
                "https://temose.com/campaign/ar3vc4rtvb",
                "https://temose.com/campaign/cin56gfvg",
                "https://temose.com/campaign/mi456bvb",
                "https://temose.com/campaign/chi4vcbvbv",
                "https://temose.com/campaign/new4fgyb",
                "https://temose.com/campaign/ta435cgv65t",
                "https://temose.com/campaign/sa56gyb",
                "https://temose.com/campaign/w456fg67",
                "https://temose.com/campaign/ho6fghy",
                "https://temose.com/campaign/ba4gf6",
                "https://temose.com/campaign/mi56vbvh",
                "https://temose.com/campaign/de4gfg6",
                "https://temose.com/campaign/dvrebvvbv"
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
                {"3D Unisex Hoodie", new Tuple<double, double>(54.04,46.99) },
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
                "3D Unisex Hoodie",
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
            try
            {
                var galleryImages = FindElement(_productImagesBy);
                //if (galleryImages.Count > 0)
                //{
                //    return galleryImages.Select(x => x.GetAttribute("src")).ToList();
                //}
                return [FindElement(_productMainImageBy).GetAttribute("src"), galleryImages.GetAttribute("src")];
            }
            catch 
            {
                return [FindElement(_productMainImageBy).GetAttribute("src")];
            }
            
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

                    var sku = $"{ProductConstants.SKU}{skuParentNumber}";

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
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Attribute1Values].Value = string.Join(',', item.Attribute1) + ",Cap";
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Attribute1Visible].Value = 1;
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Attribute1Global].Value = 0;

                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Attribute2Name].Value = ProductConstants.Attribute2Name;
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Attribute2Values].Value = string.Join(',', item.Attribute2) + ",Cap one size";
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Attribute2Visible].Value = 1;
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Attribute2Global].Value = 0;

                    #endregion

                    var rowChild = rowIndex + 1;
                    int variationPosition = 0;

                    foreach (var atribute1 in item.Attribute1)
                    {
                        var (regularPrice, salePrice) = item.PriceMapping[atribute1.Trim().ToLower()];

                        foreach (var atribute2 in item.Attribute2)
                        {
                            variationPosition++;

                            var skuChild = $"{sku}-child-{variationPosition}";

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

                            workSheet.Cells[rowChild, (int)WooExcelColumn.Attribute2Name].Value = ProductConstants.Attribute2Name;
                            workSheet.Cells[rowChild, (int)WooExcelColumn.Attribute2Values].Value = atribute2;
                            workSheet.Cells[rowChild, (int)WooExcelColumn.Attribute2Visible].Value = "";
                            workSheet.Cells[rowChild, (int)WooExcelColumn.Attribute2Global].Value = 0;

                            #endregion

                            rowChild++;
                        }
                    }

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

                    workSheet.Cells[rowChild, (int)WooExcelColumn.SalePrice].Value = "35.00";
                    workSheet.Cells[rowChild, (int)WooExcelColumn.RegularPrice].Value = "40.25";

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
                    workSheet.Cells[rowChild, (int)WooExcelColumn.Position].Value = variationPosition + 1;

                    workSheet.Cells[rowChild, (int)WooExcelColumn.Attribute1Name].Value = ProductConstants.Attribute1Name;
                    workSheet.Cells[rowChild, (int)WooExcelColumn.Attribute1Values].Value = "Cap";
                    workSheet.Cells[rowChild, (int)WooExcelColumn.Attribute1Visible].Value = "";
                    workSheet.Cells[rowChild, (int)WooExcelColumn.Attribute1Global].Value = 0;

                    workSheet.Cells[rowChild, (int)WooExcelColumn.Attribute2Name].Value = ProductConstants.Attribute2Name;
                    workSheet.Cells[rowChild, (int)WooExcelColumn.Attribute2Values].Value = "Cap one size";
                    workSheet.Cells[rowChild, (int)WooExcelColumn.Attribute2Visible].Value = "";
                    workSheet.Cells[rowChild, (int)WooExcelColumn.Attribute2Global].Value = 0;

                    rowIndex = rowChild + 1;

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

            public const string Description = @"<div class=""d-flex justify-content-between mt-7"">
<h2 class=""font-weight-bolder text-dark"">Product detail</h2>
</div>
<div class=""line-height-xl"">
<div>
<ul class=""mb-0"">
 	<li>Materials: Polyester</li>
 	<li>Printing technique: Sublimination</li>
 	<li>Our hats are available in a variety of colors and are suitable for both men and women. They provide excellent sun protection and are ideal for any outdoor activity! With adjustable snap back construction (one size fits all), everyone in your family can now wear these hats!</li>
 	<li>Designed by Fans</li>
</ul>
</div>
</div>
<div class=""line-height-xl"">* Caution: The design on the apparel can be printed smaller compared to what is shown on this screen.</div>
<div class=""line-height-xl"">* Caution: Please allow 1-3cm( 0.39-1.18 inch) differences due to manual error measurement, thank you for your understanding.</div>";

            public const string SKU = "hunting-collection-2024";
            public const string Categories = "Hunting Collection 2024";

            public const string Attribute1Name = "Style";
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
