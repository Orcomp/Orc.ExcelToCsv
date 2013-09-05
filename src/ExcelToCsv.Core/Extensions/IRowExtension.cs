namespace ExcelToCsv.Core.Extensions
{
    using NPOI.SS.UserModel;
    
    /// <summary>
    /// Extensions for NPOI.SS.UserModel.IRow
    /// </summary>
    public static class IRowExtension
    {
        /// <summary>
        /// convert cell from cell index to string that will be acceptable for csv
        /// </summary>
        /// <param name="row">the row</param>
        /// <param name="cellIndex">cell index</param>
        /// <returns>
        /// empty string if cell is empty
        /// trimmed StringCellValue if cell is string
        /// cell.ToString() if cell isn't string
        /// </returns>
        public static string CellToStringCsvValue(this IRow row, int cellIndex)
        {
            ICell cell = row.GetCell(cellIndex);
            if (cell == null)
            {
                return null; // return empty string;
            }
            //return trimmed string cell value or return ToString for cell

            return cell.CellType != CellType.STRING ? cell.ToString() : cell.StringCellValue.Trim();

        }
    }
}
