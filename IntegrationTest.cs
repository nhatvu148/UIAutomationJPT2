using System;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;

namespace JupiterTestAPI
{
    [TestClass]
    [TestCategory("Integration Test")]
    public class IntegrationTest
    {
        protected static WindowsDriver<WindowsElement> Driver;
        protected static TestContext objTestContext;
        static Actions action;
        private WebDriverWait wait;
        static WindowsElement allParts;
        static WindowsElement toolBar;
        static WindowsElement jupiter;

        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            if (Driver == null)
            {
                AppiumOptions appOptions = new AppiumOptions();
                Assert.IsNotNull(appOptions);
                appOptions.AddAdditionalCapability("app", @"C:\Program Files\TechnoStar\Jupiter-Pre_5.0\DCAD_main.exe");

                Driver = new WindowsDriver<WindowsElement>(
                    new Uri("http://127.0.0.1:4723"),
                    appOptions,
                    TimeSpan.FromMinutes(5)
                    );
                Assert.IsNotNull(Driver);
                Assert.IsNotNull(Driver.SessionId);

                objTestContext = testContext;
            }
        }

        [ClassCleanup]
        public static void ClassCleanup()
        {
            if (Driver != null)
            {
                Driver.Quit();
                Driver = null;
            }

        }

        [TestInitialize]
        public void SetupBeforeEveryTestMethod()
        {
            wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(20));
            Assert.IsNotNull(wait);
            Driver.Manage().Window.Maximize();
            jupiter = Driver.FindElementByXPath("//Window[starts-with(@Name,'Jupiter-Pre 5.0')]");
            toolBar = Driver.FindElementByName("Ribbon Tabs");
        }

        [TestMethod]
        public void CreateMesh()
        {
            Driver.FindElementByName("Open...").Click();

            System.Threading.Thread.Sleep(1000);
            action = new Actions(Driver);
            Assert.IsNotNull(action);
            action.SendKeys(@"V:\TechnoStar\Mazda\R2-IDI3_TS.jtdb");
            action.SendKeys(Keys.Enter);
            action.Perform();

            Driver.FindElementByName("Assembly").Click();
            allParts = Driver.FindElementByName("All Parts");
            var boltOld = allParts.FindElementByXPath("//TreeItem[@Name='bolt_old']");
            Assert.IsNotNull(boltOld);
            action = new Actions(Driver);
            Assert.IsNotNull(action);
            action.MoveToElement(boltOld);
            action.ContextClick();
            action.SendKeys(Keys.Up);
            action.SendKeys(Keys.Enter);
            action.Build().Perform();

            var headCyl = allParts.FindElementByXPath("//TreeItem[@Name ='HEAD-CYL']");
            Assert.IsNotNull(headCyl);
            action = new Actions(Driver);
            Assert.IsNotNull(action);
            action.MoveToElement(headCyl);
            action.ContextClick();
            action.Build().Perform();
            Driver.FindElementByXPath("//MenuItem[@Name ='Show Only']").Click();
       
            //////////////////////////////////////////////
            var geometry = toolBar.FindElementByXPath("//TabItem[@Name ='Geometry']");
            geometry.Click();
            Driver.FindElementByXPath("//SplitButton[@Name ='Show Adjacent']").Click();

            action = new Actions(Driver);
            Assert.IsNotNull(action);
            action.SendKeys(Keys.Down);
            action.SendKeys(Keys.Enter);
            action.Build().Perform();

            toolBar.FindElementByName("Home").Click();
            var find = Driver.FindElementByName("Find");
            Assert.IsNotNull(find);
            action = new Actions(Driver);
            Assert.IsNotNull(action);
            action.MoveToElement(find);
            action.MoveToElement(find, find.Size.Width / 2, find.Size.Height / 3 + 20).Click();
            action.SendKeys(Keys.Down);
            action.SendKeys(Keys.Enter);
            action.Perform();

            WindowsElement idBox = Driver.FindElementByAccessibilityId("1582");
            Assert.IsNotNull(idBox);
            InputId(15254, idBox, action, Driver, find);

            var showAdj = jupiter.FindElementByName("Show Adjacent | Faces");
            showAdj.FindElementByName("Stop Face").Click();
            InputId(19885, idBox, action, Driver, find);
            InputId(16323, idBox, action, Driver, find);
            InputId(20597, idBox, action, Driver, find);
            InputId(20487, idBox, action, Driver, find);
            InputId(20596, idBox, action, Driver, find);
            InputId(20698, idBox, action, Driver, find);
            InputId(20694, idBox, action, Driver, find);
            Assert.AreEqual("8", Driver.FindElementByAccessibilityId("5105")
                .Text.Substring(9).Trim());

            var numLayers = showAdj.FindElementByXPath("//Edit[@Name='Number of Layers']");
            Assert.IsNotNull(numLayers);
            numLayers.Clear();
            numLayers.SendKeys("100");
            showAdj.FindElementByName("Apply").Click();
            // Screenshot to compare ?

            //////////////////////////////////////////////
            

            var tools = toolBar.FindElementByName("Tools");
            tools.Click();
            var modelFilter = Driver.FindElementByName("Model Filter");
            Assert.IsNotNull(modelFilter);
            action = new Actions(Driver);
            Assert.IsNotNull(action);
            action.MoveToElement(modelFilter);
            action.MoveToElement(modelFilter, -modelFilter.Size.Width / 2,
                modelFilter.Size.Height / 2).Click().Perform();

            Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);

            var group = jupiter.FindElementByName("Group");
            var groupName = group.FindElementByAccessibilityId("1001");
            Assert.IsNotNull(groupName);
            groupName.Clear();
            groupName.SendKeys("WaterJacket");
            Driver.FindElementByName("OK").Click();

            Driver.FindElementByXPath("//Button[@Name='Application menu']").Click();
            Driver.FindElementByName("Save As...").Click();

            System.Threading.Thread.Sleep(1000);
            action = new Actions(Driver);
            action.SendKeys(@"V:\TechnoStar\Mazda\01_Groups.jtdb");
            action.SendKeys(Keys.Enter);
            action.Perform();

            //////////////////////////////////////////////
            

            var meshing = toolBar.FindElementByName("Meshing");
            meshing.Click();
            var localSettingIcon = Driver.FindElementByName("Local Settings");
            action = new Actions(Driver);
            action.MoveToElement(localSettingIcon, localSettingIcon.Size.Width / 4,
                localSettingIcon.Size.Height / 2).Click().Perform();
            Driver.FindElementByName("Face").Click();

            var localSettingName = Driver.FindElementByAccessibilityId("1003");
            Assert.IsNotNull(localSettingName);
            localSettingName.Clear();
            localSettingName.SendKeys("WaterJacket");

            Driver.FindElementByName("Group").Click();
            var waterJacketGroup = Driver.FindElementByName("WaterJacket");

            var meshSizeCB = Driver.FindElementByAccessibilityId("1005");
            if (!meshSizeCB.Selected) meshSizeCB.Click();

            Driver.FindElementByAccessibilityId("1006").Clear();
            Driver.FindElementByAccessibilityId("1006").SendKeys("1.5");

            Driver.FindElementByAccessibilityId("1007").Clear();
            Driver.FindElementByAccessibilityId("1007").SendKeys("0.5");

            Driver.FindElementByAccessibilityId("1008").Clear();
            Driver.FindElementByAccessibilityId("1008").SendKeys("3");

            action = new Actions(Driver);
            Assert.IsNotNull(action);
            action.MoveToElement(waterJacketGroup);
            action.ContextClick();
            action.SendKeys(Keys.Down);
            action.SendKeys(Keys.Enter).Perform();
            Driver.FindElementByName("Apply").Click();

            //////////////////////////////////////////////
           

            Driver.FindElementByName("Meshing").Click();
           
            action = new Actions(Driver);
            action.MoveToElement(localSettingIcon, localSettingIcon.Size.Width / 4,
                localSettingIcon.Size.Height / 2).Click().Perform();
            Driver.FindElementByName("Part").Click();

            Driver.FindElementByName("Assembly").Click();

            var localSettingParts = Driver.FindElementByAccessibilityId("1003");
            Assert.IsNotNull(localSettingParts);
            localSettingParts.Clear();
            localSettingParts.SendKeys("Guide-Seat-Valve");

            meshSizeCB = Driver.FindElementByAccessibilityId("1005");
            if (!meshSizeCB.Selected) meshSizeCB.Click();

            Driver.FindElementByAccessibilityId("1006").Clear();
            Driver.FindElementByAccessibilityId("1006").SendKeys("2");

            Driver.FindElementByAccessibilityId("1007").Clear();
            Driver.FindElementByAccessibilityId("1007").SendKeys("0.5");

            Driver.FindElementByAccessibilityId("1008").Clear();
            Driver.FindElementByAccessibilityId("1008").SendKeys("3");

            allParts = Driver.FindElementByName("All Parts");
            var guide = allParts.FindElementByName("GUIDE-VALVE");
            guide.Click();
            action = new Actions(Driver);
            action.KeyDown(Keys.Control).Perform();
            allParts.FindElementByName("valve-exh").Click();
            allParts.FindElementByName("valve-int").Click();
            allParts.FindElementByName("valve-seat-exh").Click();
            allParts.FindElementByName("valve-seat-int").Click();
            action.MoveToElement(guide);
            action.ContextClick().Perform();
            action.SendKeys(Keys.Down);
            action.SendKeys(Keys.Enter).Perform();

            action = new Actions(Driver);
            action.KeyUp(Keys.Control).Perform();

            Driver.FindElementByName("Apply").Click();

            Driver.FindElementByXPath("//Button[@Name='Application menu']").Click();
            Driver.FindElementByName("Save As...").Click();

            System.Threading.Thread.Sleep(1000);
            var saveText = Driver.FindElementByAccessibilityId("1001");
            saveText.SendKeys(@"V:\TechnoStar\Mazda\02_Local_Settings.jtdb");
            saveText.SendKeys(Keys.Enter);
        }

        public static void InputId(int faceId, WindowsElement idBox, Actions action,
                            WindowsDriver<WindowsElement> sessionJpt, WindowsElement find)
        {
            idBox.SendKeys(Keys.Control + "a" + Keys.Control);
            idBox.SendKeys(Convert.ToString(faceId));
            action = new Actions(sessionJpt);
            action.MoveToElement(find);
            action.MoveToElement(find, find.Size.Width / 2 - 30, find.Size.Height / 3 - 4).Click().Perform();
        }

        [TestCleanup]
        public void TestCleanup()
        {

        }
    }
}
