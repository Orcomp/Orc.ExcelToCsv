


namespace Orc.ExcelToCsv.Test
{
    using System;
    using System.IO;
    using Models;
    using NUnit.Framework;

    [TestFixture]
    public class TestTryCreateCsvWorkbook
    {
        #region private helpers
        /// <summary>
        /// Internal method to process excel file with expected worksheets
        /// </summary>
        /// <param name="fileName">fileName (note that this file should be placed in TestData folder</param>
        /// <param name="expectedCsvFiles">the expected count of correct worksheets</param>
        /// <param name="isExpectedCorrectWorkbook">do we expect that workbook at filename is correct</param>
        private void TestTryCreateCsv(string fileName, int expectedCsvFiles, bool isExpectedCorrectWorkbook)
        {
            string nameForOutputDirectory = Path.Combine(Environment.CurrentDirectory, "TestData", DateTime.Now.Ticks.ToString());
            Directory.CreateDirectory(nameForOutputDirectory); // create directory for output csv

            IProcessWorkbookResult resultProcessing;
            var wasFileProcessed = ProcessExcel.TryCreateCsv(Path.Combine(Environment.CurrentDirectory, "TestData", fileName), nameForOutputDirectory + "\\");

            var filesCollection = Directory.GetFiles(nameForOutputDirectory, "*.csv");

            Directory.Delete(nameForOutputDirectory, true); // delete temp data (NOTE, we do this before Assert to avoid dirty folders)

            Assert.AreEqual(wasFileProcessed, isExpectedCorrectWorkbook);
            if (wasFileProcessed)
            {
                Assert.AreEqual(expectedCsvFiles, filesCollection.Length);
            }


            
        }
        #endregion

        #region Excel2003 test methods
        #region Correct test methods
        /// <summary>
        /// Test that we can open excel 2003 files
        /// </summary>
        ////[TestCase]
        public void TestExcel2003()
        {
            TestTryCreateCsv("TestExcel2003.xls", 3, true);
        }
        #endregion

        #region Test methods with catching exception
        /// <summary>
        /// Test that file that actually 2007 but have extension like 2003 will be processed
        /// </summary>
        [TestCase]
        public void TestExcel2003Invalid()
        {
            TestTryCreateCsv("TestExcel2003Invalid.xls", 3, false);
        }

        /// <summary>
        /// Test that we'll correctly process situation when file with extension 2003 Excel was broken 
        /// (delete some important part from excel file via notepad)
        /// </summary>
        [TestCase]
        public void TestExcel2003Broken()
        {
                TestTryCreateCsv("TestExcel2003Broken.xls", 3, false);
        }
        #endregion
        #endregion

        #region Excel2007 and Excel2010 test methods

        #region Correct test methods
        /// <summary>
        /// Test that we can open excel 2007 files
        /// </summary>
        ////[TestCase]
        public void TestExcel2007()
        {
            TestTryCreateCsv("TestExcel2007.xlsx", 3, true);
        }


        /// <summary>
        /// Test that we can open excel 2010 files
        /// </summary>
        ////[TestCase]
        public void TestExcel2010()
        {
            TestTryCreateCsv("TestExcel2010_1.xlsx", 3, true);

            // not sure that this files are different from 2007 files
            // so I need to use second file for test purposing
            TestTryCreateCsv("TestExcel2010_2.xlsx", 3, true);
        }
        #endregion


        #region Test methods with catching exception
        /// <summary>
        /// Test that file that actually 2003 but have extension like 2007 will be processed
        /// </summary>
        [TestCase]
        public void TestExcel2007Invalid()
        {
            TestTryCreateCsv("TestExcel2007Invalid.xlsx", 3, false);
        }

        /// <summary>
        /// Test that we'll correctly process situation when file with extension 2007 Excel was broken 
        /// (delete some important part from excel file via notepad)
        /// </summary>
        [TestCase]
        public void TestExcel2007Broken()
        {
            TestTryCreateCsv("TestExcel2007Broken.xlsx", 3, false);
        }
        #endregion

        #endregion

        #region Common for all versions of excel tests
        /// <summary>
        /// Test that we'll return FileNotFound exception if there are no excel file with this name 
        /// </summary>
        [TestCase]
        public void TestFileNotFound()
        {
            TestTryCreateCsv("BadDataBadDataBadData123421", 3, false);
        }

        /// <summary>
        /// Test that we'll correctly process situation when file name reference to not excel file
        /// </summary>
        [TestCase]
        public void TestAnotherFileType()
        {
            TestTryCreateCsv("veryveryinvalid.txt", 3, false);
        }

        /// <summary>
        /// Test that we'll correctly process situation when we try to write in unexisting folder
        /// </summary>
        [TestCase]
        public void TestUnexistingFolder()
        {
            string unexistedFolder;
            do
            {
                unexistedFolder = "E:\\" + DateTime.Now.Ticks.ToString();
            } while (Directory.Exists(unexistedFolder));
            bool wasCreated = ProcessExcel.TryCreateCsv(Path.Combine(Environment.CurrentDirectory, "TestData", "TestExcel2003.xls"),
                                       unexistedFolder + "\\");
            Assert.AreEqual(wasCreated, false);

        }

        #endregion
    }
}
