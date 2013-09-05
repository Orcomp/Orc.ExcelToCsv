namespace ExcelToCsv.Interfaces
{
    /// <summary>
    /// Interface that described model of cell
    /// </summary>
    public interface ICsvCellModel
    {
        /// <summary>
        /// Value of cell 
        /// </summary>
        string Value { get; }
    }
}
