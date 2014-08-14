namespace Orc.ExcelToCsv.Test
{
    using System;
    using System.IO;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Models;

    [TestClass]
    public class TestTryProcessWorkbook
    {
        #region private helpers
        /// <summary>
        /// Internal method to process excel file with expected worksheets
        /// </summary>
        /// <param name="fileName">fileName (note that this file should be placed in TestData folder</param>
        /// <param name="expectedCorrectWorksheets">the expected count of correct worksheets</param>
        /// <param name="isExpectedCorrectWorkbook">do we expect that workbook at filename is correct</param>
        private void TestTryProcessing(string fileName, int expectedCorrectWorksheets, bool isExpectedCorrectWorkbook)
        {
            IProcessWorkbookResult resultProcessing;
            bool wasFileProcessed =
                ProcessExcel.TryProcessWorkbook(Path.Combine(Environment.CurrentDirectory, "TestData", fileName),
                                                out resultProcessing);
            Assert.AreEqual(wasFileProcessed, isExpectedCorrectWorkbook);
            if (wasFileProcessed)
            {
                Assert.AreEqual(expectedCorrectWorksheets, resultProcessing.SheetToCsvModels.Count);
            }
        }
        #endregion

        #region Excel2003 test methods
        #region Correct test methods
        /// <summary>
        /// Test that we can open excel 2003 files
        /// </summary>
        ////[TestMethod]
        public void TestExcel2003()
        {
            TestTryProcessing("TestExcel2003.xls", 3, true);
        }
        #endregion

        #region Test methods with catching exception
        /// <summary>
        /// Test that file that actually 2007 but have extension like 2003 will be processed
        /// </summary>
        [TestMethod]
        public void TestExcel2003Invalid()
        {
            TestTryProcessing("TestExcel2003Invalid.xls", 3, false);
        }

        /// <summary>
        /// Test that we'll correctly process situation when file with extension 2003 Excel was broken 
        /// (delete some important part from excel file via notepad)
        /// </summary>
        [TestMethod]
        public void TestExcel2003Broken()
        {
            TestTryProcessing("TestExcel2003Broken.xls", 3, false);
        }
        #endregion
        #endregion

        #region Excel2007 and Excel2010 test methods
        #region Correct test methods
        /// <summary>
        /// Test that we can open excel 2007 files
        /// </summary>
        ////[TestMethod]
        public void TestExcel2007()
        {
            TestTryProcessing("TestExcel2007.xlsx", 3, true);
        }

        /// <summary>
        /// Test that we can open excel 2010 files
        /// </summary>
        ////[TestMethod]
        public void TestExcel2010()
        {
            TestTryProcessing("TestExcel2010_1.xlsx", 3, true);

            // not sure that this files are different from 2007 files
            // so I need to use second file for test purposing
            TestTryProcessing("TestExcel2010_2.xlsx", 3, true);
        }
        #endregion

        #region Test methods with catching exception
        /// <summary>
        /// Test that file that actually 2003 but have extension like 2007 will be processed
        /// </summary>
        [TestMethod]
        public void TestExcel2007Invalid()
        {
            TestTryProcessing("TestExcel2007Invalid.xlsx", 3, false);
        }

        /// <summary>
        /// Test that we'll correctly process situation when file with extension 2007 Excel was broken 
        /// (delete some important part from excel file via notepad)
        /// </summary>
        [TestMethod]
        public void TestExcel2007Broken()
        {
            TestTryProcessing("TestExcel2007Broken.xlsx", 3, false);
        }
        #endregion
        #endregion

        #region Common test methods
        /// <summary>
        /// Test that we'll return FileNotFound exception if there are no excel file with this name 
        /// </summary>
        [TestMethod]
        public void TestFileNotFound()
        {
            TestTryProcessing("BadDataBadDataBadData123421", 3, false);
        }

        /// <summary>
        /// Test that we'll correctly process situation when file name reference to not excel file
        /// </summary>
        [TestMethod]
        public void TestAnotherFileType()
        {
            TestTryProcessing("veryveryinvalid.txt", 3, false);
        }
        #endregion 
    }
}
