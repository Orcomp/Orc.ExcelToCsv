namespace ExcelToCsv.Interfaces
{
    using System.Collections.Generic;

    /// <summary>
    /// Interface that described the row model
    /// </summary>
    public interface ICsvRowModel
    {
        /// <summary>
        /// collection of cells 
        /// </summary>
        IList<ICsvCellModel> Cells { get; }

        /// <summary>
        /// Add cells to make cell count in row equals to columnCount
        /// </summary>
        /// <param name="columnCount">count of expected columns</param>
        void NormalizeRow(int columnCount);
    }
}
