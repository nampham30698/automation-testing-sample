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
    public class ProductCalatlog : WebPageBase
    {
        private ReadOnlyCollection<IWebElement> _singleProductLinks => FindElements(By.XPath("//div[contains(@class,'woocommerce-image__wrapper')]//a"));
        private IWebElement _btnClosePopup => FindElement(By.XPath("//button[contains(@class,'klaviyo-close-form')]"));

        private IWebElement _productTitle => FindElement(By.XPath("//div[contains(@class,'product-details-wrapper')]//h1[contains(@class,'product_title')]"));
        //private IWebElement _productDescription => FindElement(By.XPath("//div[contains(@class,'woocommerce-Tabs-panel')]//div[contains(@class,'vtab_container')]"));
        private IWebElement _productDescription => FindElement(By.XPath(@"//*[@id='tab-description']"));
        private ReadOnlyCollection<IWebElement> _productPriceRange => FindElements(By.XPath("//div[contains(@class,'product-details-wrapper')]//p[contains(@class,'price')]//span[contains(@class,'woocommerce-Price-amount')]"));

        private IWebElement _productRegularPrice => FindElement(By.XPath("(//div[contains(@class,'product-details-wrapper')]//p[contains(@class,'price')]//span[contains(@class,'woocommerce-Price-amount')])[1]"));
        private IWebElement _productSalePrice => FindElement(By.XPath("(//div[contains(@class,'product-details-wrapper')]//p[contains(@class,'price')]//span[contains(@class,'woocommerce-Price-amount')])[2]"));

        private By _productVariationRegularPriceBy => By.XPath("(//div[contains(@class,'product-details-wrapper')]//span[contains(@class,'price')]//span[contains(@class,'woocommerce-Price-amount')])[1]");

        private By _productVariationSalePriceBy => By.XPath("(//div[contains(@class,'product-details-wrapper')]//span[contains(@class,'price')]//span[contains(@class,'woocommerce-Price-amount')])[2]");

        // buttom type
        private ReadOnlyCollection<IWebElement> _productButtomAttribute1 => FindElements(By.XPath("//div[contains(@class,'product-details-wrapper')]//ul[@data-attribute='attribute_style']//button"));
        private ReadOnlyCollection<IWebElement> _productButtomAttribute2 => FindElements(By.XPath("//div[contains(@class,'product-details-wrapper')]//ul[@data-attribute='attribute_size']//button"));

        // select options
        //private SelectElement _productSelectAttribute1 => FindSelectElement(By.XPath("//div[contains(@class,'product-details-wrapper')]//select[@data-attribute_name='attribute_your-style']"));
        //private SelectElement _productSelectAttribute2 => FindSelectElement(By.XPath("//div[contains(@class,'product-details-wrapper')]//select[@data-attribute_name='attribute_size']"));
        
        private SelectElement _productSelectAttribute1 => FindSelectElement(By.XPath("//div[contains(@class,'product-details-wrapper')]//select[@data-attribute_name='attribute_pa_style']"));
        private SelectElement _productSelectAttribute2 => FindSelectElement(By.XPath("//div[contains(@class,'product-details-wrapper')]//select[@data-attribute_name='attribute_pa_size']"));

        private By _productMainImageBy => By.XPath("(//div[contains(@class,'product-details-wrapper')]//div[contains(@class,'woocommerce-product-gallery__wrapper')]//img)[1]");
        private By _productImagesBy => By.XPath("//div[contains(@class,'product-details-wrapper')]//ol[contains(@class,'flex-control-thumbs')]//img");

        public ProductCalatlog(IWebDriver driver) : base(driver)
        {

        }

        public void GetProductUrls()
        {
            _btnClosePopup.Click();

            if (_singleProductLinks.Count == 0) return;

            var productMetadata = new List<ProductMetadata>();

            //productMetadata.Add(GetProductInfo("https://championssport.store/product/limited-edition-pittsburgh-penguins-grateful-dead-night-hoodie-set-v3nt17112416id10ds11/"));

            var links = _singleProductLinks.Select(x => x.GetAttribute("href")).ToList();

            //var links = new List<string>()
            //{
            //    "https://merchsport.net/product/_miami-dolphins-2024-salute-to-service-club-hoodie/",
            //    "https://merchsport.net/product/_-pittsburgh-steelers-2024-salute-to-service-hoodie/",
            //    "https://merchsport.net/product/_dallas-cowboys-2024-salute-to-service-club-hoodie/",
            //    "https://merchsport.net/product/__kansas-city-chiefs-2024-salute-to-service-club-hoodie/",
            //    "https://merchsport.net/product/__new-england-patriots-hoodie-camo-2024-salute-to-service-club/",
            //    "https://merchsport.net/product/__philadelphia-eagles-2024-salute-to-service-club-hoodie/",
            //    "https://merchsport.net/product/_arizona-cardinals-2024-salute-to-service-club-hoodie/",
            //    "https://merchsport.net/product/_atlanta-falcons-2024-salute-to-service-club-hoodie/",
            //    "https://merchsport.net/product/_baltimore-ravens-2024-salute-to-service-club-hoodie/",
            //    "https://merchsport.net/product/_carolina-panthers-2024-salute-to-service-club-hoodie/",
            //    "https://merchsport.net/product/_chicago-bears-2024-salute-to-service-club-hoodie/",
            //    "https://merchsport.net/product/_cincinnati-bengals-2024-salute-to-service-club-hoodie/",
            //    "https://merchsport.net/product/_cleveland-browns-2024-salute-to-service-club-hoodie/",
            //    "https://merchsport.net/product/_denver-broncos-2024-salute-to-service-club-hoodie/",
            //    "https://merchsport.net/product/_detroit-lions-2024-salute-to-service-club-hoodie/",
            //    "https://merchsport.net/product/detroit-lions-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/_green-bay-packers-2024-salute-to-service-club-hoodie/",
            //    "https://merchsport.net/product/_-houston-texans-2024-salute-to-service-club-hoodie/",
            //    "https://merchsport.net/product/_indianapolis-colts-2024-salute-to-service-club-hoodie/",
            //    "https://merchsport.net/product/_jacksonville-jaguars-2024-salute-to-service-club-hoodie/",
            //    "https://merchsport.net/product/kansas-city-chiefs-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/_las-vegas-raiders-2024-salute-to-service-club-hoodie/",
            //    "https://merchsport.net/product/_los-angeles-chargers-2024-salute-to-service-club-hoodie/",
            //    "https://merchsport.net/product/_los-angeles-rams-2024-salute-to-service-hoodie/",
            //    "https://merchsport.net/product/_minnesota-vikings-2024-salute-to-service-club-hoodie/",
            //    "https://merchsport.net/product/_new-orleans-saints-2024-salute-to-service-club-hoodie/",
            //    "https://merchsport.net/product/_new-york-giants-2024-salute-to-service-club-hoodie/",
            //    "https://merchsport.net/product/_new-york-jets-2024-salute-to-service-club-hoodie/",
            //    "https://merchsport.net/product/pittsburgh-steelers-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/_san-francisco-49ers-hoodie-camo-2024-salute-to-service-club/",
            //    "https://merchsport.net/product/seattle-seahawks-nike-arctic-camo-2024-salute-to-service-club-hoodie-2024-limited-edition-2/",
            //    "https://merchsport.net/product/_tampa-bay-buccaneers-hoodie-camo-2024-salute-to-service-club-limited-edition/",
            //    "https://merchsport.net/product/washington-commanders-hoodie-camo-2024-salute-to-service-club-2/",
            //    "https://merchsport.net/product/_arizona-cardinals-jersey-camo-2024-salute-to-service-club-limited-edition-4/",
            //    "https://merchsport.net/product/_los-angeles-chargers-jersey-camo-2024-salute-to-service-club-limited-edition-4/",
            //    "https://merchsport.net/product/_miami-dolphins-nike-arctic-camo-2024-salute-to-service-club-flash-%e2%9a%a1%ef%b8%8f-sale-jersey-limited-editions-5/",
            //    "https://merchsport.net/product/_new-orleans-saints-jersey-camo-2024-salute-to-service-club-limited-edition-3/",
            //    "https://merchsport.net/product/_san-francisco-49ers-jersey-camo-2024-salute-to-service-club-5/",
            //    "https://merchsport.net/product/dallas-cowboys-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/los-angeles-rams-jersey-camo-2024-salute-to-service-club/",
            //    "https://merchsport.net/product/new-england-patriots-jersey-camo-2024-salute-to-service-club-5/",
            //    "https://merchsport.net/product/_-kansas-city-chiefs-nike-arctic-camo-2024-salute-to-service-club-jersey-9/",
            //    "https://merchsport.net/product/arizona-cardinals-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/atlanta-falcons-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/baltimore-ravens-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/buffalo-bills-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/carolina-panthers-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/chicago-bears-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/cincinnati-bengals-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/cleveland-browns-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/denver-broncos-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/green-bay-packers-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/houston-texans-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/indianapolis-colts-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/jacksonville-jaguars-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/las-vegas-raiders-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/los-angeles-chargers-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/los-angeles-rams-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/miami-dolphins-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/minnesota-vikings-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/new-england-patriots-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/new-orleans-saints-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/new-york-giants-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/new-york-jets-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/philadelphia-eagles-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/san-francisco-49ers-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/seattle-seahawks-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/tampa-bay-buccaneers-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/tennessee-titans-2024-salute-to-service-club-hoodie/",
            //    "https://merchsport.net/product/tennessee-titans-x-duck-hunting-hoodie-2024-limited-edition/",
            //    "https://merchsport.net/product/washington-commanders-x-duck-hunting-hoodie-2024-limited-edition/"
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

                var data = new ProductMetadata()
                {
                    Url = url,
                    Title = _productTitle.Text,
                    Description = _productDescription.GetAttribute("innerHTML").Trim(),
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

            if (ProductConstants.VariationType == VariationType.Buttom)
            {
                var firstSize = _productButtomAttribute2.Skip(1).First();

                var sizeClasses = firstSize.GetAttribute("class");

                if (!sizeClasses.Contains("selected", StringComparison.OrdinalIgnoreCase))
                {
                    firstSize.Click();
                }

                foreach (var style in _productButtomAttribute1)
                {
                    if (style.Text.Trim().Equals("Cap", StringComparison.OrdinalIgnoreCase)) continue;
                    Thread.Sleep(200);

                    var styleClasses = style.GetAttribute("class");
                    if (!styleClasses.Contains("selected", StringComparison.OrdinalIgnoreCase))
                    {
                        style.Click();
                    }

                    Thread.Sleep(500);
                    var regularPrice = Parser.ParseDoube(FindElement(_productVariationRegularPriceBy).Text.Replace("$", ""));
                    var salePrice = Parser.ParseDoube(FindElement(_productVariationSalePriceBy).Text.Replace("$", ""));

                    priceMapping.Add(style.Text.Trim().ToLower(), new Tuple<double, double>(regularPrice, salePrice));
                }

            }
            else
            {
                _productSelectAttribute2.SelectByIndex(2);

                var selectAttribute2Indexes = _productSelectAttribute1.Options.Skip(1).Select((element, index) => index + 1).ToList();

                foreach (var index in selectAttribute2Indexes)
                {
                    Thread.Sleep(200);
                    _productSelectAttribute1.SelectByIndex(index);
                    Thread.Sleep(500);

                    var regularPrice = Parser.ParseDoube(FindElement(_productVariationRegularPriceBy).Text.Replace("$", ""));
                    var salePrice = Parser.ParseDoube(FindElement(_productVariationSalePriceBy).Text.Replace("$", ""));

                    priceMapping.Add(_productSelectAttribute1.SelectedOption.Text.Trim().ToLower(), new Tuple<double, double>(regularPrice, salePrice));
                }
            }

            return priceMapping;
        }

        private List<string> GetAttribute1Values()
        {
            if(ProductConstants.VariationType == VariationType.Buttom)
            {
                return _productButtomAttribute1.Select(x => x.Text.Trim()).Where(x => !x.Equals("Cap", StringComparison.OrdinalIgnoreCase)).ToList();
            }
            else
            {
                return _productSelectAttribute1.Options.Skip(1).Select(x => x.Text.Trim()).Where(x => !x.Equals("cap", StringComparison.OrdinalIgnoreCase)).ToList();
            }
        }

        private List<string> GetAttribute2Values()
        {
            if (ProductConstants.VariationType == VariationType.Buttom)
            {
                return _productButtomAttribute2.Select(x => x.Text.Trim()).Where(x => !x.Contains("Cap one size", StringComparison.OrdinalIgnoreCase) && 
                !x.Contains("One Size Cap", StringComparison.OrdinalIgnoreCase) && !x.Contains("Caps one size", StringComparison.OrdinalIgnoreCase) && 
                !x.Contains("Kid", StringComparison.OrdinalIgnoreCase)).ToList();
            }
            else
            {
                return _productSelectAttribute2.Options.Skip(1).Select(x => x.Text).Where(x => !x.Contains("Universal fit", StringComparison.OrdinalIgnoreCase)).ToList();
            }
        }

        private List<string> GetImages()
        {
            var galleryImages = FindElements(_productImagesBy);
            if (galleryImages.Count > 0)
            {
                return galleryImages.Select(x => x.GetAttribute("src")).ToList();
            }
            return [FindElement(_productMainImageBy).GetAttribute("src")];
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
                    workSheet.Cells[rowIndex, (int)WooExcelColumn.Description].Value = item.Description; //ProductConstants.Description
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

                    workSheet.Cells[rowChild, (int)WooExcelColumn.SalePrice].Value = "39.99";
                    workSheet.Cells[rowChild, (int)WooExcelColumn.RegularPrice].Value = "35.99";

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

                    rowIndex = rowChild +  1;

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

            public const string Description = @"";

            public const string SKU = "minnesota-football";
            public const string Categories = "Minnesota Football";

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
