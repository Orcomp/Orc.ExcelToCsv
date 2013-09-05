namespace ExcelToCsv.Core.Model
{
    using Interfaces;

    /// <summary>
    /// Class that represent model of cell
    /// </summary>
    internal class CsvCellModel : ICsvCellModel 
    {
        /// <summary>
        /// Value of cell 
        /// </summary>
        public string _value;
        public string Value {
            get { return _value; }
        }

        /// <summary>
        /// Initialize the cell model
        /// </summary>
        /// <param name="value">value of cell</param>
        public CsvCellModel(string value)
        {
            _value = value;
        }
    }
}
