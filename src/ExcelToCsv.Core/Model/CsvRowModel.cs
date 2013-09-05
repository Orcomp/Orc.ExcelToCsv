namespace ExcelToCsv.Core.Model
{
    using System.Collections.Generic;
    using Interfaces;

    /// <summary>
    /// Class that represent the row model
    /// </summary>
    internal class CsvRowModel : ICsvRowModel
    {
        /// <summary>
        /// collection of cells 
        /// </summary>
        private readonly IList<ICsvCellModel> _cells = new List<ICsvCellModel>();
        public IList<ICsvCellModel> Cells
        {
            get { return _cells; }
        }

        /// <summary>
        /// Initialize CSVRowModel from collection of CellModels
        /// </summary>
        /// <param name="cells">collection of cells</param>
        public CsvRowModel(IEnumerable<ICsvCellModel> cells)
        {
            foreach (var cellModel in cells)
            {
                Cells.Add(cellModel);
            }
        }

        /// <summary>
        /// Initialize CSVRowModel from collection of string (convert inside to CSVCellModel) 
        /// </summary>
        /// <param name="cellValues">collection of cell values</param>
        public CsvRowModel(IEnumerable<string> cellValues)
        {
            foreach (var value in cellValues)
            {
                Cells.Add(new CsvCellModel(value));
            }
        }

        /// <summary>
        /// Add cells to make cell count in row equals to columnCount
        /// </summary>
        /// <param name="columnCount">count of expected columns</param>
        public void NormalizeRow(int columnCount)
        {
            while (Cells.Count != columnCount) Cells.Add(new CsvCellModel(""));
        }

    }
}
