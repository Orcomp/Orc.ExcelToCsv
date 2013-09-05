

namespace ExcelToCsv.Core
{
    using System.IO;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using Model;
    using Interfaces;
    using NPOI.XSSF.UserModel;


    /// <summary>
    /// Method with static methods for process excel files.
    /// </summary>
    public static class ProcessExcel
    {
        /// <summary>
        /// Read file and create csv file in outputDirectory for each sheet
        /// </summary>
        /// <param name="excelFileName">the file name of excel file</param>
        /// <param name="outputDirectory">the working directory where new csv files will be placed</param>
        public static void CreateCsv(string excelFileName, string outputDirectory)
        {
            ProcessWorkbook(excelFileName).ExportToCsv(outputDirectory);
        }

        public static IProcessWorkbookResult ProcessWorkbook(string excelFileName)
        {
            // read workbook from file 
            using (var file = new FileStream(excelFileName, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook;
                var fileInfo = new FileInfo(excelFileName);
                if (fileInfo.Extension == ".xls")
                {
                    workbook = new HSSFWorkbook(file); // hssf workbook is using for Excel 2003
                }
                else
                {
                    workbook = new XSSFWorkbook(file); // xssf workbook is using for Excel 2007
                }
                return ProcessWorkbook(workbook); // and then process this workbook 
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
    }
}
