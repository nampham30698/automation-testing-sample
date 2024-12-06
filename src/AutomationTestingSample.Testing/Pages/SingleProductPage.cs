using AutomationTestingSample.Core.Extensions;
using AutomationTestingSample.Core.Helpers;
using AutomationTestingSample.Core.Reports;
using AutomationTestingSample.Core.Utilities;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.ObjectModel;

namespace AutomationTestingSample.Testing.Pages
{
    public class SingleProductPage : WebPageBase
    {
        private ReadOnlyCollection<IWebElement> _singleProductLinks => FindElements(By.XPath("//div[contains(@class,'woocommerce-image__wrapper')]//a"));
        private IWebElement _btnClosePopup => FindElement(By.XPath("//button[contains(@class,'klaviyo-close-form')]"));

        private IWebElement _productTitle => FindElement(By.XPath("//div[contains(@class,'product-main')]//h1[contains(@class,'product_title')]"));
        //private IWebElement _productDescription => FindElement(By.XPath("//div[contains(@class,'woocommerce-Tabs-panel')]//div[contains(@class,'vtab_container')]"));
        private IWebElement _productDescription => FindElement(By.XPath(@"//*[@id='tab-description']"));
        private ReadOnlyCollection<IWebElement> _productPriceRange => FindElements(By.XPath("//div[contains(@class,'product-main')]//p[contains(@class,'price')]//span[contains(@class,'woocommerce-Price-amount')]"));

        private IWebElement _productRegularPrice => FindElement(By.XPath("(//div[contains(@class,'product-main')]//p[contains(@class,'price')]//span[contains(@class,'woocommerce-Price-amount')])[1]"));
        private IWebElement _productSalePrice => FindElement(By.XPath("(//div[contains(@class,'product-main')]//p[contains(@class,'price')]//span[contains(@class,'woocommerce-Price-amount')])[2]"));

        private By _productVariationRegularPriceBy => By.XPath("(//div[contains(@class,'product-main')]//span[contains(@class,'price')]//span[contains(@class,'woocommerce-Price-amount')])[1]");

        private By _productVariationSalePriceBy => By.XPath("(//div[contains(@class,'product-main')]//span[contains(@class,'price')]//span[contains(@class,'woocommerce-Price-amount')])[2]");

        // buttom type
        private ReadOnlyCollection<IWebElement> _productButtomAttribute1 => FindElements(By.XPath("//div[contains(@class,'product-main')]//ul[@data-attribute_name='attribute_pa_color']//li"));
        private ReadOnlyCollection<IWebElement> _productButtomAttribute2 => FindElements(By.XPath("//div[contains(@class,'product-main')]//ul[@data-attribute='attribute_size']//button"));

        // select options
        //private SelectElement _productSelectAttribute1 => FindSelectElement(By.XPath("//div[contains(@class,'product-main')]//select[@data-attribute_name='attribute_your-style']"));
        //private SelectElement _productSelectAttribute2 => FindSelectElement(By.XPath("//div[contains(@class,'product-main')]//select[@data-attribute_name='attribute_size']"));
        
        private SelectElement _productSelectAttribute1 => FindSelectElement(By.XPath("//div[contains(@class,'product-main')]//select[@data-attribute_name='attribute_pa_style']"));
        private SelectElement _productSelectAttribute2 => FindSelectElement(By.XPath("//div[contains(@class,'product-main')]//select[@data-attribute_name='attribute_pa_size']"));

        private By _productMainImageBy => By.XPath("(//div[contains(@class,'product-main')]//div[contains(@class,'woocommerce-product-gallery__wrapper')]//img)[1]");
        private By _productImagesBy => By.XPath("//div[contains(@class,'product-main')]//div[contains(@class,'flickity-slider')]//img[contains(@class,'attachment-woocommerce_thumbnail')]");

        public SingleProductPage(IWebDriver driver) : base(driver)
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
                 "https://southernlifeus.net/product/alabama-crimson-tide-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/notre-dame-fighting-irish-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/auburn-tigers-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/ohio-state-buckeyes-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/clemson-tigers-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/florida-gators-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/florida-state-seminoles-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/georgia-bulldogs-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/tennessee-volunteers-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/oklahoma-sooners-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/iowa-state-cyclones-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/kentucky-wildcats-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/lsu-tigers-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/west-virginia-mountaineers-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/south-carolina-gamecocks-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/mississippi-state-bulldogs-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/oklahoma-state-cowboys-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/nc-state-wolfpack-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/nebraska-cornhuskers-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/north-carolina-tar-heels-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/northern-illinois-huskies-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/arkansas-razorbacks-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/boise-state-broncos-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/iowa-hawkeyes-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/ole-miss-rebels-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/oregon-ducks-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/penn-state-nittany-lions-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/pittsburgh-panthers-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/michigan-wolverines-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/indiana-hoosiers-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/texas-am-aggies-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/texas-longhorns-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/ucf-knights-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/usc-trojans-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/stanford-cardinal-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/miami-hurricanes-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/wisconsin-badgers-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/85-michigan-state-spartans-1831/",
"https://southernlifeus.net/product/100-wayne-state-warriors-8084/",
"https://southernlifeus.net/product/arizona-state-sun-devils-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/arkansas-state-red-wolves-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/army-black-knights-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/baylor-bears-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/byu-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/cincinnati-bearcats-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/coastal-carolina-chanticleers-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/colorado-buffaloes-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/colorado-state-rams-credit-card-holder-leather-small-wallet-custom-name/",
"https://southernlifeus.net/product/49-cornell-big-red-5193/",
"https://southernlifeus.net/product/50-davidson-wildcats-7689/",
"https://southernlifeus.net/product/51-harvard-crimson-7610/",
"https://southernlifeus.net/product/52-hawaii-rainbow-warriors-3068/",
"https://southernlifeus.net/product/53-houston-cougars-3960/",
"https://southernlifeus.net/product/54-illinois-fighting-illini-3254/",
"https://southernlifeus.net/product/55-louisville-cardinals-2853/",
"https://southernlifeus.net/product/56-maryland-terrapins-3682/",
"https://southernlifeus.net/product/57-minnesota-golden-gophers-8041/",
"https://southernlifeus.net/product/58-montana-grizzlies-9663/",
"https://southernlifeus.net/product/59-navy-midshipmen-8668/",
"https://southernlifeus.net/product/60-northern-colorado-bears-7460/",
"https://southernlifeus.net/product/61-purdue-boilermakers-5303/",
"https://southernlifeus.net/product/62-rutgers-scarlet-knights-3197/",
"https://southernlifeus.net/product/63-san-jose-state-spartans-7696/",
"https://southernlifeus.net/product/64-smu-mustangs-3125/",
"https://southernlifeus.net/product/65-south-florida-bulls-7583/",
"https://southernlifeus.net/product/67-tcu-horned-frogs-7167/",
"https://southernlifeus.net/product/68-temple-owls-3070/",
"https://southernlifeus.net/product/69-texas-tech-red-raiders-5118/",
"https://southernlifeus.net/product/70-tulane-green-wave-4554/",
"https://southernlifeus.net/product/71-ucla-2473/",
"https://southernlifeus.net/product/72-utah-state-aggies-1306/",
"https://southernlifeus.net/product/73-utah-utes-6789/",
"https://southernlifeus.net/product/74-utsa-roadrunners-8812/",
"https://southernlifeus.net/product/75-vanderbilt-commodores-5874/",
"https://southernlifeus.net/product/76-virginia-tech-hokies-8697/",
"https://southernlifeus.net/product/77-wake-forest-demon-deacons-2140/",
"https://southernlifeus.net/product/78-washington-huskies-3211/",
"https://southernlifeus.net/product/79-washington-state-cougars-6594/",
"https://southernlifeus.net/product/80-colgate-raiders-8311/",
"https://southernlifeus.net/product/81-drake-bulldogs-5796/",
"https://southernlifeus.net/product/82-marshall-thundering-herd-8014/",
"https://southernlifeus.net/product/83-missouri-tigers-2369/",
"https://southernlifeus.net/product/86-syracuse-orange-7002/",
"https://southernlifeus.net/product/87-northwestern-wildcats-5178/",
"https://southernlifeus.net/product/88-kansas-state-wildcats-2878/",
"https://southernlifeus.net/product/89-kansas-jayhawks-5117/",
"https://southernlifeus.net/product/90-duke-blue-devils-5700/",
"https://southernlifeus.net/product/91-virginia-cavaliers-3725/",
"https://southernlifeus.net/product/92-boston-college-eagles-4913/",
"https://southernlifeus.net/product/93-georgia-tech-yellow-jackets-5520/",
"https://southernlifeus.net/product/94-california-golden-bears-6560/",
"https://southernlifeus.net/product/95-arizona-wildcats-5551/",
"https://southernlifeus.net/product/96-air-force-falcons-3815/",
"https://southernlifeus.net/product/97-idaho-state-bengals-1928/",
"https://southernlifeus.net/product/98-morehouse-maroon-tigers-7909/",
"https://southernlifeus.net/product/99-southern-miss-golden-eagles-2720/"
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
                    //Attribute2 = GetAttribute2Values(),
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

            //if (ProductConstants.VariationType == VariationType.Buttom)
            //{
            //    var firstSize = _productButtomAttribute2.Skip(1).First();

            //    var sizeClasses = firstSize.GetAttribute("class");

            //    if (!sizeClasses.Contains("selected", StringComparison.OrdinalIgnoreCase))
            //    {
            //        firstSize.Click();
            //    }

            //    foreach (var style in _productButtomAttribute1)
            //    {
            //        //if (style.Text.Trim().Equals("Cap", StringComparison.OrdinalIgnoreCase)) continue;
            //        Thread.Sleep(200);

            //        var styleClasses = style.GetAttribute("class");
            //        if (!styleClasses.Contains("selected", StringComparison.OrdinalIgnoreCase))
            //        {
            //            style.Click();
            //        }

            //        Thread.Sleep(500);
            //        var regularPrice = Parser.ParseDoube(FindElement(_productVariationRegularPriceBy).Text.Replace("$", ""));
            //        var salePrice = Parser.ParseDoube(FindElement(_productVariationSalePriceBy).Text.Replace("$", ""));

            //        priceMapping.Add(style.Text.Trim().ToLower(), new Tuple<double, double>(regularPrice, salePrice));
            //    }

            //}
            //else
            //{
            //    _productSelectAttribute2.SelectByIndex(2);

            //    var selectAttribute2Indexes = _productSelectAttribute1.Options.Skip(1).Select((element, index) => index + 1).ToList();

            //    foreach (var index in selectAttribute2Indexes)
            //    {
            //        Thread.Sleep(200);
            //        _productSelectAttribute1.SelectByIndex(index);
            //        Thread.Sleep(500);

            //        var regularPrice = Parser.ParseDoube(FindElement(_productVariationRegularPriceBy).Text.Replace("$", ""));
            //        var salePrice = Parser.ParseDoube(FindElement(_productVariationSalePriceBy).Text.Replace("$", ""));

            //        priceMapping.Add(_productSelectAttribute1.SelectedOption.Text.Trim().ToLower(), new Tuple<double, double>(regularPrice, salePrice));
            //    }
            //}

            //foreach (var style in _productButtomAttribute1)
            //{
            //    Thread.Sleep(200);

            //    style.Click();

            //    Thread.Sleep(500);
            //    var regularPrice = Parser.ParseDoube(FindElement(_productVariationRegularPriceBy).Text.Replace("$", ""));
            //    var salePrice = Parser.ParseDoube(FindElement(_productVariationSalePriceBy).Text.Replace("$", ""));

            //    priceMapping.Add(style.Text.Trim().ToLower(), new Tuple<double, double>(regularPrice, salePrice));
            //}

            priceMapping.Add("black", new Tuple<double, double>(39.99, 25.95));
            priceMapping.Add("brown", new Tuple<double, double>(39.99, 25.95));


            return priceMapping;
        }

        private List<string> GetAttribute1Values()
        {
            //if(ProductConstants.VariationType == VariationType.Buttom)
            //{
            //    return _productButtomAttribute1.Select(x => x.Text.Trim()).ToList();
            //    //return _productButtomAttribute1.Select(x => x.Text.Trim()).Where(x => !x.Equals("Cap", StringComparison.OrdinalIgnoreCase)).ToList();
            //}
            //else
            //{
            //    return _productSelectAttribute1.Options.Skip(1).Select(x => x.Text.Trim()).Where(x => !x.Equals("cap", StringComparison.OrdinalIgnoreCase) && !x.Contains("Snapback Cap", StringComparison.OrdinalIgnoreCase)).ToList();
            //}
            return ["Black", "Brown"];
        }

        private List<string> GetAttribute2Values()
        {
            if (ProductConstants.VariationType == VariationType.Buttom)
            {
                return _productButtomAttribute2.Select(x => x.Text.Trim()).ToList();

                //return _productButtomAttribute2.Select(x => x.Text.Trim()).Where(x => !x.Contains("Cap one size", StringComparison.OrdinalIgnoreCase) &&
                //!x.Contains("One Size Cap", StringComparison.OrdinalIgnoreCase) && !x.Contains("Caps one size", StringComparison.OrdinalIgnoreCase) &&
                //!x.Contains("Kid", StringComparison.OrdinalIgnoreCase)).ToList();
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
                return galleryImages.Select(x => x.GetAttribute("src").Replace("-247x296","")).ToList();
            }
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
        }

        public static class ProductConstants
        {
            public const string Simple = "simple"; 
            public const string Variable = "variable";
            public const string variation = "variation";

            public const string Description = @"";

            public const string SKU = "ncaa-card-holder-leather-small-wallet";
            public const string Categories = "NCAA Card Holder Leather Small Wallet";

            public const string Attribute1Name = "Color";
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
