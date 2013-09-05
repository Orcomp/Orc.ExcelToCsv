namespace ExcelToCsv.Interfaces
{
    using System.Collections.Generic;

    /// <summary>
    /// Interface that described csv model
    /// </summary>
    public interface ICsvModel
    {

        /// <summary>
        /// Count of columns
        /// </summary>
        int ColumnCount { get; }


        /// <summary>
        /// Collection of rows
        /// </summary>
        IList<ICsvRowModel> Rows { get; }


        /// <summary>
        /// Export csv model to file
        /// </summary>
        /// <param name="fileName">the file name of new file</param>
        void ExportToCsv(string fileName);
    }
}
