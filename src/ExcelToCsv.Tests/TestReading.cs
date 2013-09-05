using System;
using System.IO;
using ExcelToCsv.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelToCsv.Tests
{
    [TestClass]
    public class TestReading
    {
        private void TestProcessing(string fileName, int expectedCorrectWorksheets)
        {
            var resultProcessing = ProcessExcel.ProcessWorkbook(fileName);
            Assert.AreEqual(expectedCorrectWorksheets, resultProcessing.SheetToCsvModels.Count); // it should be three csv files
        }

        [TestMethod]
        public void TestExcel2003()
        {
            var fileName = Path.Combine(Environment.CurrentDirectory, "TestData", "TestExcel.xls");
            TestProcessing(fileName, 3);
        }

        [TestMethod]
        public void TestExcel2007()
        {
            var fileName = Path.Combine(Environment.CurrentDirectory, "TestData", "TestExcel.xlsx");
            TestProcessing(fileName, 3);
        }

        [TestMethod]
        public void TestCreateCsvExcel2007()
        {
            var fileName = Path.Combine(Environment.CurrentDirectory, "TestData", "TestExcel.xlsx");
            var resultProcessing = ProcessExcel.ProcessWorkbook(fileName);
            Assert.AreEqual(3, resultProcessing.SheetToCsvModels.Count); // it should be three csv files
        }
    }
}
