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
    [TestCategory("Jupiter Test")]
    public class JupiterTest
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
            wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(50));
            Assert.IsNotNull(wait);
            Driver.Manage().Window.Maximize();
            jupiter = Driver.FindElementByXPath("//Window[starts-with(@Name,'Jupiter-Pre 5.0')]");
            toolBar = Driver.FindElementByName("Ribbon Tabs");
        }

        [TestMethod]
        public void CreateMesh1()
        {
            Driver.FindElementByName("Open...").Click();

            System.Threading.Thread.Sleep(1000);
            action = new Actions(Driver);
            Assert.IsNotNull(action);
            action.SendKeys(@"D:\TechnoStar\Mazda\R2-IDI3_TS.jtdb");
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
            action.SendKeys(@"D:\TechnoStar\Mazda\01_Groups.jtdb");
            action.SendKeys(Keys.Enter);
            action.Perform();
        }

        [TestMethod]
        public void CreateMesh2()
        {
            Driver.FindElementByName("Open...").Click();

            System.Threading.Thread.Sleep(1000);
            action = new Actions(Driver);
            Assert.IsNotNull(action);
            action.SendKeys(@"D:\TechnoStar\Mazda\01_Groups.jtdb");
            action.SendKeys(Keys.Enter);
            action.Perform();

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


            meshing.Click();

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
            saveText.SendKeys(@"D:\TechnoStar\Mazda\02_Local_Settings.jtdb");
            saveText.SendKeys(Keys.Enter);
        }

        [TestMethod]
        public void CreateMesh3()
        {
            Driver.FindElementByName("Open...").Click();

            System.Threading.Thread.Sleep(1000);
            action = new Actions(Driver);
            Assert.IsNotNull(action);
            action.SendKeys(@"D:\TechnoStar\Mazda\02_Local_Settings.jtdb");
            action.SendKeys(Keys.Enter);
            action.Perform();

            toolBar.FindElementByName("Meshing").Click();
            Driver.FindElementByName("Surface Meshing").Click();

            allParts = Driver.FindElementByName("All Parts");
            var block = allParts.FindElementByName("BLOCK");
            block.Click();
            action = new Actions(Driver);
            action.KeyDown(Keys.Control).Perform();
            allParts.FindElementByName("GUIDE-VALVE").Click();
            allParts.FindElementByName("HEAD-CYL").Click();
            allParts.FindElementByName("liner").Click();
            allParts.FindElementByName("valve-exh").Click();
            allParts.FindElementByName("valve-int").Click();
            allParts.FindElementByName("valve-seat-exh").Click();
            allParts.FindElementByName("valve-seat-int").Click();
            action.MoveToElement(block);
            action.ContextClick().Perform();
            action.SendKeys(Keys.Down);
            action.SendKeys(Keys.Enter).Perform();

            action = new Actions(Driver);
            action.KeyUp(Keys.Control).Perform();

            var meshSurf = jupiter.FindElementByName("Meshing Surf Meshing");
            meshSurf.FindElementByAccessibilityId("1001").Clear();
            meshSurf.FindElementByAccessibilityId("1001").SendKeys("5");
            meshSurf.FindElementByAccessibilityId("1002").Clear();
            meshSurf.FindElementByAccessibilityId("1002").SendKeys("1");
            meshSurf.FindElementByAccessibilityId("1003").Clear();
            meshSurf.FindElementByAccessibilityId("1003").SendKeys("10");
            meshSurf.FindElementByAccessibilityId("1004").Clear();
            meshSurf.FindElementByAccessibilityId("1004").SendKeys("1");
            meshSurf.FindElementByName("OK").Click();

            System.Threading.Thread.Sleep(180000);
            action = new Actions(Driver);
            action.MoveToElement(meshSurf).Click();
            action.SendKeys(Keys.Escape);
            action.Perform();

            Driver.FindElementByXPath("//Button[@Name='Application menu']").Click();
            Driver.FindElementByName("Save As...").Click();

            System.Threading.Thread.Sleep(1000);
            var saveText = Driver.FindElementByAccessibilityId("1001");
            saveText.SendKeys(@"D:\TechnoStar\Mazda\03_Surface_meshing.jtdb");
            saveText.SendKeys(Keys.Enter);
        }

        [TestMethod]
        public void CreateMesh4()
        {
            Driver.FindElementByName("Open...").Click();

            System.Threading.Thread.Sleep(1000);
            action = new Actions(Driver);
            Assert.IsNotNull(action);
            action.SendKeys(@"D:\TechnoStar\Mazda\03_Surface_meshing.jtdb");
            action.SendKeys(Keys.Enter);
            action.Perform();

            toolBar.FindElementByName("Mesh Cleanup").Click();
            var freeEdgeIcon = Driver.FindElementByName("Free Edges");
            freeEdgeIcon.Click();

            allParts = Driver.FindElementByName("All Parts");
            var block = allParts.FindElementByName("BLOCK");
            block.Click();
            action = new Actions(Driver);
            action.KeyDown(Keys.Control).Perform();
            allParts.FindElementByName("GUIDE-VALVE").Click();
            allParts.FindElementByName("HEAD-CYL").Click();
            allParts.FindElementByName("liner").Click();
            allParts.FindElementByName("valve-exh").Click();
            allParts.FindElementByName("valve-int").Click();
            allParts.FindElementByName("valve-seat-exh").Click();
            allParts.FindElementByName("valve-seat-int").Click();
            action.MoveToElement(block);
            action.ContextClick().Perform();
            action.SendKeys(Keys.Down);
            action.SendKeys(Keys.Enter).Perform();

            action = new Actions(Driver);
            action.KeyUp(Keys.Control).Perform();

            var freeEdges = jupiter.FindElementByName("Mesh Quality | Free Edges");
            var nonManifold = freeEdges.FindElementByAccessibilityId("1008");
            if (nonManifold.Selected) nonManifold.Click();
            freeEdges.FindElementByName("OK").Click();

            System.Threading.Thread.Sleep(5000);

            var manual2DIcon = Driver.FindElementByName("Manual 2D");

            action = new Actions(Driver);
            Assert.IsNotNull(action);
            action.MoveToElement(manual2DIcon);
            action.MoveToElement(manual2DIcon, manual2DIcon.Size.Width / 2, manual2DIcon.Size.Height / 2 + 20).Click();
            action.Perform();

            Driver.FindElementByName("Equivalence").Click();
            var manual2D = jupiter.FindElementByName("Manual Cleanup 2D | Equivalence");

            //var mergetowards = manual2D.FindElementByName("Merge Towards");
            //mergetowards.Click();
            // mergetowards.FindElementByName("First Node").Click();
            var multipleNode = manual2D.FindElementByName("Multiple Node");
            if (multipleNode.Selected) multipleNode.Click();

            manual2D.FindElementByName("Node").Click();
            toolBar.FindElementByName("Home").Click();
            var find = Driver.FindElementByName("Find");
            Assert.IsNotNull(find);
            action = new Actions(Driver);
            Assert.IsNotNull(action);
            action.MoveToElement(find);
            action.MoveToElement(find, find.Size.Width / 2, find.Size.Height / 3 + 20).Click();
            action.SendKeys("n");
            action.SendKeys(Keys.Enter);
            action.Perform();

            WindowsElement idBox = Driver.FindElementByAccessibilityId("1582");
            Assert.IsNotNull(idBox);
            InputId(977545, idBox, action, Driver, find);
            InputId(976130, idBox, action, Driver, find);
            manual2D.FindElementByName("Apply").Click();

            InputId(977548, idBox, action, Driver, find);
            InputId(977578, idBox, action, Driver, find);
            manual2D.FindElementByName("Apply").Click();

            InputId(977520, idBox, action, Driver, find);
            InputId(976091, idBox, action, Driver, find);
            manual2D.FindElementByName("Apply").Click();

            InputId(977523, idBox, action, Driver, find);
            InputId(977565, idBox, action, Driver, find);
            manual2D.FindElementByName("Apply").Click();

            InputId(977533, idBox, action, Driver, find);
            InputId(976119, idBox, action, Driver, find);
            manual2D.FindElementByName("Apply").Click();

            InputId(977536, idBox, action, Driver, find);
            InputId(977572, idBox, action, Driver, find);
            manual2D.FindElementByName("Apply").Click();

            InputId(977502, idBox, action, Driver, find);
            InputId(976093, idBox, action, Driver, find);
            manual2D.FindElementByName("Apply").Click();

            InputId(977505, idBox, action, Driver, find);
            InputId(977512, idBox, action, Driver, find);
            manual2D.FindElementByName("Apply").Click();


            toolBar.FindElementByName("Mesh Cleanup").Click();
            jupiter.FindElementByName("Manual Cleanup 2D").FindElementByName("Close").Click();
            action = new Actions(Driver);
            Assert.IsNotNull(action);
            action.MoveToElement(manual2DIcon);
            action.MoveToElement(manual2DIcon, manual2DIcon.Size.Width / 2, manual2DIcon.Size.Height / 2 + 20).Click();
            action.Perform();

            Driver.FindElementByName("Collapse").Click();
            var collapse = jupiter.FindElementByName("Manual Cleanup 2D | Collapse");

            collapse.FindElementByName("Node").Click();
            toolBar.FindElementByName("Home").Click();
            InputId(977545, idBox, action, Driver, find);
            collapse.FindElementByName("Minimum").Click();
            collapse.FindElementByName("Apply").Click();

            collapse.FindElementByName("Node").Click();
            InputId(977520, idBox, action, Driver, find);
            collapse.FindElementByName("Minimum").Click();
            collapse.FindElementByName("Apply").Click();

            collapse.FindElementByName("Node").Click();
            InputId(977533, idBox, action, Driver, find);
            collapse.FindElementByName("Minimum").Click();
            collapse.FindElementByName("Apply").Click();

            collapse.FindElementByName("Node").Click();
            InputId(977502, idBox, action, Driver, find);
            collapse.FindElementByName("Minimum").Click();
            collapse.FindElementByName("Apply").Click();

            collapse.FindElementByName("Node").Click();
            InputId(381270, idBox, action, Driver, find);
            action = new Actions(Driver);
            Assert.IsNotNull(action);
            action.MoveToElement(find);
            action.MoveToElement(find, find.Size.Width / 2, find.Size.Height / 3 + 20).Click();
            action.SendKeys("1");
            action.SendKeys(Keys.Enter);
            action.Perform();
            InputId("381270-976137", idBox, action, Driver, find);
            collapse.FindElementByName("Apply").Click();

            collapse.FindElementByName("Node").Click();
            action = new Actions(Driver);
            Assert.IsNotNull(action);
            action.MoveToElement(find);
            action.MoveToElement(find, find.Size.Width / 2, find.Size.Height / 3 + 20).Click();
            action.SendKeys("n");
            action.SendKeys(Keys.Enter);
            action.Perform();
            InputId(977585, idBox, action, Driver, find);
            action = new Actions(Driver);
            Assert.IsNotNull(action);
            action.MoveToElement(find);
            action.MoveToElement(find, find.Size.Width / 2, find.Size.Height / 3 + 20).Click();
            action.SendKeys("1");
            action.SendKeys(Keys.Enter);
            action.Perform();
            InputId("364714-977585", idBox, action, Driver, find);
            collapse.FindElementByName("Apply").Click();

            collapse.FindElementByName("Node").Click();
            action = new Actions(Driver);
            Assert.IsNotNull(action);
            action.MoveToElement(find);
            action.MoveToElement(find, find.Size.Width / 2, find.Size.Height / 3 + 20).Click();
            action.SendKeys("n");
            action.SendKeys(Keys.Enter);
            action.Perform();
            InputId(977585, idBox, action, Driver, find);
            action = new Actions(Driver);
            Assert.IsNotNull(action);
            action.MoveToElement(find);
            action.MoveToElement(find, find.Size.Width / 2, find.Size.Height / 3 + 20).Click();
            action.SendKeys("1");
            action.SendKeys(Keys.Enter);
            action.Perform();
            InputId("977558-977585", idBox, action, Driver, find);
            collapse.FindElementByName("Apply").Click();

            collapse.FindElementByName("Node").Click();
            action = new Actions(Driver);
            Assert.IsNotNull(action);
            action.MoveToElement(find);
            action.MoveToElement(find, find.Size.Width / 2, find.Size.Height / 3 + 20).Click();
            action.SendKeys("n");
            action.SendKeys(Keys.Enter);
            action.Perform();
            InputId(6016, idBox, action, Driver, find);
            action = new Actions(Driver);
            Assert.IsNotNull(action);
            action.MoveToElement(find);
            action.MoveToElement(find, find.Size.Width / 2, find.Size.Height / 3 + 20).Click();
            action.SendKeys("1");
            action.SendKeys(Keys.Enter);
            action.Perform();
            InputId("6016-555096", idBox, action, Driver, find);
            collapse.FindElementByName("Apply").Click();

            collapse.FindElementByName("Node").Click();
            InputId(555097, idBox, action, Driver, find);
            collapse.FindElementByName("Minimum").Click();
            collapse.FindElementByName("Apply").Click();
            collapse.FindElementByName("Next>").Click();
            collapse.FindElementByName("Apply").Click();

            collapse.FindElementByName("Node").Click();
            InputId(10030, idBox, action, Driver, find);
            collapse.FindElementByName("Minimum").Click();
            collapse.FindElementByName("Apply").Click();
            collapse.FindElementByName("Apply").Click();
            collapse.FindElementByName("Apply").Click();

            collapse.FindElementByName("Node").Click();
            InputId(386909, idBox, action, Driver, find);
            collapse.FindElementByName("Minimum").Click();
            collapse.FindElementByName("Apply").Click();
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

        public static void InputId(string faceId, WindowsElement idBox, Actions action,
                            WindowsDriver<WindowsElement> sessionJpt, WindowsElement find)
        {
            idBox.SendKeys(Keys.Control + "a" + Keys.Control);
            idBox.SendKeys(faceId);
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
