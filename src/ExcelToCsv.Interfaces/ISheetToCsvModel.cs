namespace ExcelToCsv.Interfaces
{
    /// <summary>
    /// Interface that described sheet to csv model
    /// </summary>
    public interface ISheetToCsvModel
    {
        /// <summary>
        /// CsvModel
        /// </summary>
        ICsvModel CsvModel { get; }


        /// <summary>
        /// name of sheet
        /// </summary>
        string SheetName { get; }
    }
}
