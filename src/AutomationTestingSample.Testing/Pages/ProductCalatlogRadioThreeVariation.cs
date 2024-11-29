using AutomationTestingSample.Core.Helpers;
using AutomationTestingSample.Core.Reports;
using OfficeOpenXml;
using OpenQA.Selenium;
using System.Collections.ObjectModel;

namespace AutomationTestingSample.Testing.Pages
{
    // url https://gotyourstyle.com/product-category/clothings/hoodie/nfl-fishing-collection-2024
    public class ProductCalatlogRadioThreeVariation : WebPageBase
    {
        private ReadOnlyCollection<IWebElement> _singleProductLinks => FindElements(By.XPath("//div[contains(@class,'woocommerce-image__wrapper')]//a"));

        private IWebElement _btnClosePopup => FindElement(By.XPath("//button[contains(@class,'klaviyo-close-form')]"));

        private IWebElement _productTitle => FindElement(By.XPath("//main[contains(@class,'site-main')]//h1[contains(@class,'product_title')]"));

        private IWebElement _productDescription => FindElement(By.XPath(@"//*[@id='tab-description']"));
        private ReadOnlyCollection<IWebElement> _productPriceRange => FindElements(By.XPath("//main[contains(@class,'site-main')]//p[contains(@class,'price')]//span[contains(@class,'woocommerce-Price-amount')]"));

        private IWebElement _productRegularPrice => FindElement(By.XPath("(//main[contains(@class,'site-main')]//p[contains(@class,'price')]//span[contains(@class,'woocommerce-Price-amount')])[1]"));
        private IWebElement _productSalePrice => FindElement(By.XPath("(//main[contains(@class,'site-main')]//p[contains(@class,'price')]//span[contains(@class,'woocommerce-Price-amount')])[2]"));

        // label type
        private ReadOnlyCollection<IWebElement> _productButtomAttribute1 => FindElements(By.XPath("//main[contains(@class,'site-main')]//div[contains(@data-validation,'Choose Hoodie Type')]//label"));
        private ReadOnlyCollection<IWebElement> _productButtomAttribute2 => FindElements(By.XPath("//main[contains(@class,'site-main')]//div[contains(@data-validation,'Choose Your Size')]//label"));

        private By _productMainImageBy => By.XPath("(//main[contains(@class,'site-main')]//div[contains(@class,'woocommerce-product-gallery__wrapper')]//img)[1]");
        private By _productImagesBy => By.XPath("//main[contains(@class,'site-main')]//ol[contains(@class,'flex-control-thumbs')]//img");

        public ProductCalatlogRadioThreeVariation(IWebDriver driver) : base(driver)
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
                 "https://gotyourstyle.com/product/washington-commanders-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/tennessee-titans-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/tampa-bay-buccaneers-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/seattle-seahawks-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/san-francisco-49ers-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/pittsburgh-steelers-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/new-york-jets-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/new-york-giants-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/new-orleans-saints-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/new-england-patriots-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/minnesota-vikings-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/miami-dolphins-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/los-angeles-rams-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/los-angeles-chargers-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/las-vegas-raiders-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/kansas-city-chiefs-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/jacksonville-jaguars-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/indianapolis-colts-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/denver-broncos-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/dallas-cowboys-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/cleveland-browns-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/cincinnati-bengals-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/chicago-bears-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/carolina-panthers-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/atlanta-falcons-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/arizona-cardinals-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/philadelphia-eagles-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/houston-texans-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/green-bay-packers-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/detroit-lions-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/buffalo-bills-nfl-x-fishing-2024-limited-edition-hoodie/",
                "https://gotyourstyle.com/product/baltimore-ravens-nfl-x-fishing-2024-limited-edition-hoodie/"

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
                {"Regular Hoodie", new Tuple<double, double>(51.95,40.95) },
                {"Zipper Hoodie", new Tuple<double, double>(54.95,43.95) },
                {"Fleece Hoodie", new Tuple<double, double>(61.95,50.95) },
                {"Zipper Fleece Hoodie", new Tuple<double, double>(64.95,55.95) },
                {"Mask Hoodie", new Tuple<double, double>(53.95,42.95) },
                {"Jogger", new Tuple<double, double>(41.95,30.95) },
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
                "Regular Hoodie",
                "Zipper Hoodie",
                "Fleece Hoodie",
                "Zipper Fleece Hoodie",
                "Mask Hoodie",
                "Jogger"
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
            var galleryImages = FindElements(_productImagesBy);
            if (galleryImages.Count > 0)
            {
                return galleryImages.Select(x => x.GetAttribute("src").Replace("-100x100", "")).ToList();
            }
            return [FindElement(_productMainImageBy).GetAttribute("src")];
        }

        private void ExportExel(List<ProductMetadata> data)
        {
            var filePath = FileHelpers.ProjectPath + "woo_product_template_three_variation.xlsx";

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
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Attribute1Values].Value = string.Join(',', item.Attribute1) + ",Cap";
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Attribute1Visible].Value = 1;
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Attribute1Global].Value = 0;

                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Attribute2Name].Value = ProductConstants.Attribute2Name;
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Attribute2Values].Value = string.Join(',', item.Attribute2) + ",Cap one size";
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Attribute2Visible].Value = 1;
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Attribute2Global].Value = 0;

                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Attribute3Name].Value = ProductConstants.Attribute3Name;
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Attribute3Values].Value = "Black,Green,Beige";
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Attribute3Visible].Value = 1;
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Attribute3Global].Value = 0;

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

                            workSheet.Cells[rowChild, (int)WooExcelColumn.Attribute3Name].Value = ProductConstants.Attribute3Name;

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

                    workSheet.Cells[rowChild, (int)WooExcelColumn.SalePrice].Value = "22.95";
                    workSheet.Cells[rowChild, (int)WooExcelColumn.RegularPrice].Value = "33.95";

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

                    workSheet.Cells[rowChild, (int)WooExcelColumn.Attribute3Name].Value = ProductConstants.Attribute3Name;

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

            public const string Description = @"<b>KEY FEATURES:</b>
<ul>
 	<li aria-level=""1"">Occasion: casual outfit for everyday use. Suit for school, home, parties, uniforms, weddings, sports, outdoor activities, and camping, to name a few.</li>
 	<li aria-level=""1"">It is suitable for special occasions such as Christmas, birthdays, festivals, and housewarming presents.</li>
</ul>
<b>INFORMATION:</b>
<ul>
 	<li aria-level=""1"">Skin-friendly fabrics include white Arctic Velvet and Polyester, a soft, toasty cloth that may keep you warm on frigid days.</li>
 	<li aria-level=""1"">To keep you warm, the shirt has a strap and a hood.</li>
 	<li aria-level=""1"">The hoodie has a pocket on the front that may be utilized for extra storage or warmth.</li>
</ul>
<strong>PRINTS:</strong> Dye-sublimation printing

<strong>WASHABLE:</strong> Machine wash.

<strong>Baltimore Ravens NFL x Fishing 2024 Limited Edition Hoodie: A Catch for Ravens Fans Who Love Fishing</strong>

The <strong>Baltimore Ravens NFL x Fishing 2024 Limited Edition Hoodie</strong> is designed for fans who enjoy both football and fishing. This hoodie combines the Baltimore Ravens branding with fishing-themed graphics and a camouflage pattern, creating a unique and appealing piece of apparel for Ravens fans who want to show their love for both their team and their favorite pastime.

<strong>Design Details that Blend Football and Fishing</strong>
<ul>
 	<li><strong>Camouflage Pattern:</strong> The hoodie features a camouflage pattern, making it ideal for outdoor activities like fishing. This pattern also adds a rugged and outdoorsy aesthetic. The specific camouflage pattern appears to vary slightly between images, possibly representing different available options.</li>
 	<li><strong>Ravens Logo and “Baltimore” Text:</strong> The front of the hoodie displays the Baltimore Ravens logo and the word “BALTIMORE,” clearly identifying the team. The option for custom name personalization above the location adds a personal touch.</li>
 	<li><strong>Fishing-Themed Back Design:</strong> The back of the hoodie features a stylized American flag graphic with a largemouth bass superimposed over it. A fishing rod and reel are also incorporated into the design, further emphasizing the fishing theme. The Ravens logo appears in the upper left corner of the flag, connecting the team to the overall design.</li>
 	<li><strong>NFL Shield Logo:</strong> A small NFL shield logo on the left sleeve adds an official touch.</li>
 	<li><strong>American Flag Patch:</strong> A subdued American flag patch on the left sleeve adds a patriotic element.</li>
 	<li><strong>Matching Cap:</strong> The included cap features the Ravens logo on a camouflage background, complementing the hoodie’s design.</li>
</ul>
<strong>A Hoodie for Ravens Fans Who Enjoy the Outdoors</strong>

The <strong>Baltimore Ravens NFL x Fishing 2024 Limited Edition Hoodie</strong> is more than just a piece of fan apparel. It’s a way for Ravens fans to express their passion for both their team and the outdoors, specifically fishing. The combination of team branding, fishing graphics, and the camouflage pattern creates a unique and appealing design that is perfect for anglers and football enthusiasts alike.";

            public const string SKU = "nfl-fishing-collection-2024";
            public const string Categories = "NFL Fishing Collection 2024";

            public const string Attribute1Name = "Hoodie Type";

            public const string Attribute2Name = "Size";

            public const string Attribute3Name = "Hoodie Color";

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
