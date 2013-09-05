namespace ExcelToCsv.Interfaces
{
    using System.Collections.Generic;

    public interface IProcessWorkbookResult
    {
        /// <summary>
        /// collection of sheet to csv models
        /// </summary>
        IList<ISheetToCsvModel> SheetToCsvModels { get; }


        /// <summary>
        /// Export existing result to csv in working directory
        /// </summary>
        /// <param name="workingDirectory">the working directory where csv files will be placed</param>
        void ExportToCsv(string workingDirectory);

    }
}
