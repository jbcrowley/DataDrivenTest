using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using RetailChannels.ATG.Common;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Threading;

namespace DataDrivenTest
{
    [TestFixtureSource("Products")]
    [Parallelizable(ParallelScope.Children)]
    public class UnitTest1
    {
        // path settings
        static readonly string projectDirectory = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.Parent.Parent.FullName;

        // Excel info
        static readonly string sheetFileName = "SampleData.xlsx";
        static readonly string run = "Y";
        static readonly string sheetFullPath = Path.Combine(projectDirectory, sheetFileName);
        static readonly string sheetName = "Sheet1";

        // SauceLabs
        ThreadLocal<IWebDriver> Driver = new ThreadLocal<IWebDriver>();

        // misc
        private DataRowView row;

        public UnitTest1(DataRowView row)
        {
            this.row = row;
        }

        [SetUp]
        public void Setup()
        {
            Driver.Value = new ChromeDriver();
            Driver.Value.Manage().Window.Maximize();
            Driver.Value.Url = "https://www.ssfcu.org/";
        }

        [Test]
        public void TestMethod1()
        {
            string outputFolder = Path.Combine(projectDirectory, @"Results");
            string outputFullPath = Path.Combine(outputFolder, $"{row["ID"].ToString()}.txt");

            StringBuilder output = new StringBuilder();
            Assert.True(true, row["ID"].ToString());
            Driver.Value.FindElement(By.Id("Username")).SendKeys(row["ID"].ToString());
            Driver.Value.FindElement(By.Id("Password")).SendKeys("Passw0rd");
            // Driver.Value.FindElement(By.XPath("//button[contains(.,'Log In')]")).Click());

            // if something fails or needs to be written
            if (true)
            {
                output.Append(row["ID"].ToString());
                File.WriteAllLines(outputFullPath, new string[] { output.ToString() });
            }
        }

        [TearDown]
        public void Teardown()
        {
            Driver.Value.Quit();
        }

        public static IEnumerable<TestFixtureData> Products
        {
            get
            {
                ExcelOleDbReader reader = new ExcelOleDbReader();
                DataTable dataTable = reader.ImportExcelFile(sheetFullPath, sheetName);
                DataView dataView = new DataView(dataTable);
                dataView.RowFilter = $"RUN = '{run}'"; // filters out empty data rows and filters to only those rows that should be run
                foreach (DataRowView row in dataView)
                {
                    yield return new TestFixtureData(row);
                }
            }
        }
    }
}