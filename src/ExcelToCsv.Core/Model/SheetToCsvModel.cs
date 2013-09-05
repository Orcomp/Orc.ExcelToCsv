namespace ExcelToCsv.Core.Model
{
    using Interfaces;
    using NPOI.SS.UserModel;


    /// <summary>
    /// Class that represented sheet to csv model
    /// </summary>
    internal class SheetToCsvModel : ISheetToCsvModel 
    {
        /// <summary>
        /// CsvModel
        /// </summary>
        private readonly ICsvModel _csvModel;
        public ICsvModel CsvModel
        {
            get { return _csvModel; }
        }

        /// <summary>
        /// name of sheet
        /// </summary>
        private readonly string _sheetName;
        public string SheetName
        {
            get { return _sheetName; }
        }
        
        /// <summary>
        /// Constructor that initialize CsvModel
        /// </summary>
        /// <param name="sheet">ISheet taken from external library for processing excel</param>
        public SheetToCsvModel(ISheet sheet)
        {
            var internalCsvModel = new CsvModel();
            
            // go through the all rows
            for (int rownum = 0; rownum <= sheet.LastRowNum; rownum++)
            {
                var row = sheet.GetRow(rownum);
                if (row != null)
                {
                    internalCsvModel.AddRow(row);
                }
            }
            _csvModel = internalCsvModel;
            _sheetName = sheet.SheetName;
        }
    }
}
