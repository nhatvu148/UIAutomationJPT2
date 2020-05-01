using System;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;

namespace AutoJPT2
{
    [TestClass]
    [TestCategory("Jupiter Pre V5.0")]
    public class JupiterTest
    {
        static WindowsDriver<WindowsElement> Driver;
        static TestContext objTestContext;
        static Actions action;

        [ClassInitialize]
        public static void PrepareForTestingJupiter(TestContext testContext)
        {
            Debug.WriteLine("Hello ClassInitialize");

            AppiumOptions appOptions = new AppiumOptions();
            appOptions.AddAdditionalCapability("app", @"C:\Program Files\TechnoStar\Jupiter-Pre_5.0\DCAD_main.exe");

            Driver = new WindowsDriver<WindowsElement>(
                new Uri("http://127.0.0.1:4723"),
                appOptions
                );

            objTestContext = testContext;
        }

        [ClassCleanup]
        public static void CleanupAfterAllTests()
        {
            Debug.WriteLine("Hello ClassCleanup");
            
            Driver.Quit();
            /*Driver.FindElementByName("Application menu").Click();
            action = new Actions(Driver);
            action.SendKeys(Keys.Right);
            action.SendKeys(Keys.Enter);
            action.Perform();*/
        }

        private WebDriverWait wait;

        [TestInitialize]
        public void SetupBeforeEveryTestMethod()
        {
            Debug.WriteLine("Hello Jupiter");
            wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(20));
            Driver.Manage().Window.Maximize();
        }

        [TestMethod]
        public void CreateCube()
        {
            Driver.FindElementByXPath("//TabItem[@Name='Geometry']").Click();

            Driver.FindElementByXPath("//SplitButton[@Name='Part']").Click();
            action = new Actions(Driver);
            action.SendKeys(Keys.Down);
            action.SendKeys(Keys.Enter);
            action.Perform();

            //Driver.FindElementByXPath("//MenuItem[@Name='Cube']").Click();

            var partName = Driver.FindElementByAccessibilityId("1004");
            partName.SendKeys("My First Cube");

            var btnApply = Driver.FindElementByName("Apply");
            var btnOk = Driver.FindElementByName("OK");

            Assert.IsTrue(btnApply.Enabled);

            var originX = Driver.FindElementByAccessibilityId("1007");
            originX.SendKeys("aaa");
            Assert.IsFalse(btnApply.Enabled);
            Assert.IsFalse(btnOk.Enabled);
            originX.Clear();
            originX.SendKeys("0");

            var originY = Driver.FindElementByAccessibilityId("1008");
            originY.SendKeys("1234");
            Assert.IsTrue(btnApply.Enabled);
            Assert.IsTrue(btnOk.Enabled);
            originY.Clear();
            originY.SendKeys("0");

            var originZ = Driver.FindElementByAccessibilityId("1009");
            originZ.Clear();
            originZ.SendKeys("0");
            var lengthX = Driver.FindElementByAccessibilityId("1010");
            lengthX.Clear();
            lengthX.SendKeys("12");
            var lengthY = Driver.FindElementByAccessibilityId("1011");
            lengthY.Clear();
            lengthY.SendKeys("12");
            var lengthZ = Driver.FindElementByAccessibilityId("1012");
            lengthZ.Clear();
            lengthZ.SendKeys("12");
            var numX = Driver.FindElementByAccessibilityId("1013");
            numX.Clear();
            numX.SendKeys("24");
            var numY = Driver.FindElementByAccessibilityId("1014");
            numY.Clear();
            numY.SendKeys("24");
            var numZ = Driver.FindElementByAccessibilityId("1015");
            numZ.Clear();
            numZ.SendKeys("24");
            Assert.IsTrue(btnApply.Enabled);
            Assert.IsTrue(btnOk.Enabled);

            if (btnOk.Enabled)
            {
                btnOk.Click();
            }

            Driver.FindElementByXPath("//TreeItem[@Name='All Parts']").Click();

            action = new Actions(Driver);
            var mainWindow = Driver.FindElementByAccessibilityId("59648");
            action.MoveToElement(mainWindow);
            // Right Click
            action.ContextClick();
            action.SendKeys(Keys.Down);
            action.SendKeys(Keys.Enter);
            action.Build();
            action.Perform();
            action.Perform();
            Driver.FindElementByName("Associated Pick").Click();

            action = new Actions(Driver);
            action.SendKeys(Keys.Up);
            action.SendKeys(Keys.Up);
            action.SendKeys(Keys.Up);
            action.SendKeys(Keys.Up);
            action.SendKeys(Keys.Enter);
            action.Build();
            action.Perform();

            //Compare number of nodes
            Assert.AreEqual("3176", Driver.FindElementByAccessibilityId("5105")
                .Text.Substring(9).Trim());

            //.GetAttribute("LegacyIAccessible.Value")
            //Assert.AreEqual("0", originX.Text);

        }

        [TestCleanup]
        public void CleanupAfterEveryTestMethod()
        {

        }
    }
}
