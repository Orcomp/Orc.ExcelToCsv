namespace ExcelToCsv.Core.Model
{
    using System.Collections.Generic;
    using Interfaces;

    /// <summary>
    /// Class that represent workbook processing result (list of sheets and ability to export this sheets to CSV)
    /// </summary>
    internal class ProcessWorkbookResult : IProcessWorkbookResult 
    {
        /// <summary>
        /// collection of sheet to csv models
        /// </summary>
        private readonly IList<ISheetToCsvModel> _sheetToCsvModels = new List<ISheetToCsvModel>(); 
        public IList<ISheetToCsvModel> SheetToCsvModels
        {
            get { return _sheetToCsvModels; }
        }

        /// <summary>
        /// Export existing result to csv in working directory
        /// </summary>
        /// <param name="workingDirectory">the working directory where csv files will be placed</param>
        public void ExportToCsv(string workingDirectory)
        {
            foreach (var csvModel in SheetToCsvModels)
            {
                csvModel.CsvModel.ExportToCsv(string.Format("{0}{1}.csv", workingDirectory, csvModel.SheetName));
            }
        }
    }
}
