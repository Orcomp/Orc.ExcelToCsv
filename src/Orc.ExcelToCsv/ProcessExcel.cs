// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ProcessExcel.cs" company="Orcomp development team">
//   Copyright (c) 2008 - 2014 Orcomp development team. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------


namespace Orc.ExcelToCsv.Models
{
    using System;
    using System.IO;
    using ICSharpCode.SharpZipLib.Zip;
    using NPOI.HSSF.UserModel;
    using NPOI.POIFS.FileSystem;
    using NPOI.SS.UserModel;
    using NPOI.XSSF.UserModel;

    /// <summary>
    /// Method with static methods for process excel files.
    /// </summary>
    public static class ProcessExcel
    {
        #region Methods
        /// <summary>
        /// Read file and create csv file in outputDirectory for each sheet
        /// </summary>
        /// <param name="excelFileName">the file name of excel file</param>
        /// <param name="outputDirectory">the working directory where new csv files will be placed</param>
        public static void CreateCsv(string excelFileName, string outputDirectory)
        {
            if (!((new FileInfo(excelFileName)).Exists))
            {
                throw new FileNotFoundException("There are no excel files");
            }
            if (!((new DirectoryInfo(outputDirectory)).Exists))
            {
                throw new DirectoryNotFoundException("There are no output directory");
            }
            ProcessWorkbook(excelFileName).ExportToCsv(outputDirectory);
        }

        /// <summary>
        /// Try create csv files in outputDirectory for each sheet
        /// </summary>
        /// <param name="excelFileName">the file name of excel file</param>
        /// <param name="outputDirectory">the working directory where new csv files will be placed</param>
        /// <returns>true if nothing was happened, otherwize false </returns>
        public static bool TryCreateCsv(string excelFileName, string outputDirectory)
        {
            try
            {
                CreateCsv(excelFileName, outputDirectory);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Create IProcessWorkbookResult
        /// </summary>
        /// <param name="excelFileName">the file name of excel</param>
        /// <returns>result of processing excelFileName</returns>
        public static IProcessWorkbookResult ProcessWorkbook(string excelFileName)
        {
            if (!((new FileInfo(excelFileName)).Exists))
            {
                throw new FileNotFoundException("There are no excel files");
            }

            // read workbook from file 
            using (var file = new FileStream(excelFileName, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook;
                var fileInfo = new FileInfo(excelFileName);
                try
                {
                    if (fileInfo.Extension == ".xls")
                    {
                        workbook = new HSSFWorkbook(file); // hssf workbook is using for Excel 2003
                    }
                    else if (fileInfo.Extension == ".xlsx")
                    {
                        workbook = new XSSFWorkbook(file); // xssf workbook is using for Excel 2007
                    }
                    else
                    {
                        throw new ExcelInvalidException("Unknown version of file");
                    }
                }
                catch (OfficeXmlFileException)
                {
                    throw new ExcelInvalidException("Invalid excel file");
                }
                catch (IOException)
                {
                    throw new ExcelInvalidException("Cannot read headers in excel file. Possibly broken");
                }
                catch (ZipException)
                {
                    throw new ExcelInvalidException("Cannot create zip archive from xlsx file. Possibly broken");
                }
                return ProcessWorkbook(workbook); // and then process this workbook 
            }
        }

        public static bool TryProcessWorkbook(string excelFileName, out IProcessWorkbookResult workbookResult)
        {
            try
            {
                workbookResult = ProcessWorkbook(excelFileName);
                return true;
            }
            catch (Exception)
            {
                workbookResult = null;
                return false;
            }
        }

        /// <summary>
        /// Process workbook (go through the all sheets and 
        /// </summary>
        /// <param name="workbook">the workbook</param>
        /// <param name="workingDirectory">directory</param>
        public static IProcessWorkbookResult ProcessWorkbook(IWorkbook workbook)
        {
            var result = new ProcessWorkbookResult();

            // go through the all sheets 
            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                var sheet = workbook.GetSheetAt(i);
                result.SheetToCsvModels.Add(ProcessSheet(sheet)); // and process this sheet (create csv file)
            }

            return result;
        }

        /// <summary>
        /// Method to process sheet (for every sheets we're creating new csv file with same name in working directory)
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="workingDirectory"></param>
        public static ISheetToCsvModel ProcessSheet(ISheet sheet)
        {
            return new SheetToCsvModel(sheet);
        }
        #endregion
    }
}