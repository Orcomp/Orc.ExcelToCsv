namespace ExcelToCsv.Tests
{
    using System;
    using System.IO;
    using Core;
    using Core.Model.Exceptions;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    public class TestProcessWorkbook
    {
        #region private helpers
        /// <summary>
        /// Internal method to process excel file with expected worksheets
        /// </summary>
        /// <param name="fileName">fileName (note that this file should be placed in TestData folder</param>
        /// <param name="expectedCorrectWorksheets">the expected count of correct worksheets</param>
        private void TestProcessing(string fileName, int expectedCorrectWorksheets)
        {
            var resultProcessing = ProcessExcel.ProcessWorkbook(Path.Combine(Environment.CurrentDirectory, "TestData", fileName));
            Assert.AreEqual(expectedCorrectWorksheets, resultProcessing.SheetToCsvModels.Count); // it should be three csv files
        }
        #endregion

        #region Excel2003 test methods
        #region Correct test methods 
        /// <summary>
        /// Test that we can open excel 2003 files
        /// </summary>
        [TestMethod]
        public void TestExcel2003()
        {
            TestProcessing("TestExcel2003.xls", 3);
        }
        #endregion

        #region Test methods with catching exception
        /// <summary>
        /// Test that file that actually 2007 but have extension like 2003 will be processed
        /// </summary>
        [TestMethod]
        public void TestExcel2003Invalid()
        {
            try
            {
                TestProcessing("TestExcel2003Invalid.xls", 3);
            }
            catch (ExcelInvalidException)
            {
                //good exception as we want
            }
        }

        /// <summary>
        /// Test that we'll correctly process situation when file with extension 2003 Excel was broken 
        /// (delete some important part from excel file via notepad)
        /// </summary>
        [TestMethod]
        public void TestExcel2003Broken()
        {
            try
            {
                TestProcessing("TestExcel2003Broken.xls", 3);
            }
            catch (ExcelInvalidException)
            {
                //good exception as we want
            }
        }
        #endregion
        #endregion

        #region Excel2007 and Excel2010 test methods

        #region Correct test methods
        /// <summary>
        /// Test that we can open excel 2007 files
        /// </summary>
        [TestMethod]
        public void TestExcel2007()
        {
            TestProcessing("TestExcel2007.xlsx", 3);
        }


        /// <summary>
        /// Test that we can open excel 2010 files
        /// </summary>
        [TestMethod]
        public void TestExcel2010()
        {
            TestProcessing("TestExcel2010_1.xlsx", 3);

            // not sure that this files are different from 2007 files
            // so I need to use second file for test purposing
            TestProcessing("TestExcel2010_2.xlsx", 3);
        }
        #endregion


        #region Test methods with catching exception
        /// <summary>
        /// Test that file that actually 2003 but have extension like 2007 will be processed
        /// </summary>
        [TestMethod]
        public void TestExcel2007Invalid()
        {
            try
            {
                TestProcessing("TestExcel2007Invalid.xlsx", 3);
            }
            catch (ExcelInvalidException)
            {
                //good exception as we want
            }
        }

        /// <summary>
        /// Test that we'll correctly process situation when file with extension 2007 Excel was broken 
        /// (delete some important part from excel file via notepad)
        /// </summary>
        [TestMethod]
        public void TestExcel2007Broken()
        {
            try
            {
                TestProcessing("TestExcel2007Broken.xlsx", 3);
            }
            catch (ExcelInvalidException)
            {
                //good exception as we want
            }
        }
        #endregion

        #endregion

        #region Common for all versions of excel tests
        /// <summary>
        /// Test that we'll return FileNotFound exception if there are no excel file with this name 
        /// </summary>
        [TestMethod]
        public void TestFileNotFound()
        {
            try
            {
                TestProcessing("BadDataBadDataBadData123421", 3);
            }
            catch (FileNotFoundException)
            {
                //good exception
            }
        }

        /// <summary>
        /// Test that we'll correctly process situation when file name reference to not excel file
        /// </summary>
        [TestMethod]
        public void TestAnotherFileType()
        {
            try
            {
                TestProcessing("veryveryinvalid.txt", 3);
            }
            catch (ExcelInvalidException)
            {
                //good exception
            }
        }
        #endregion
    }
}
