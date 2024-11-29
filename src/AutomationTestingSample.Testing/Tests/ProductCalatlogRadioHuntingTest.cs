﻿using AutomationTestingSample.Testing.Pages;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomationTestingSample.Testing.Tests
{
    public class ProductCalatlogRadioHuntingTest : TestBase
    {
        public ProductCalatlogRadioHuntingTest(string browserType, int browserWidth, int browserHeight) : base(browserType, browserWidth, browserHeight)
        {
        }

        [TestCase]
        public void GetProductUrls()
        {
            var product = new ProductCalatlogRadioHunting(Driver);
            product.GetProductUrls();
        }
    }
}