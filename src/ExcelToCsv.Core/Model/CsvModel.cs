﻿namespace ExcelToCsv.Core.Model
{
    using CsvHelper;
    using NPOI.SS.UserModel;
    using System.Collections.Generic;
    using System.IO;
    using Extensions;
    using Interfaces;

    
    /// <summary>
    /// model that used to convert ISheet to csv
    /// </summary>
    internal class CsvModel : ICsvModel 
    {
        /// <summary>
        /// field that contain column count (changed while 
        /// </summary>
        private int _prevColumnCount;

        /// <summary>
        /// Count of columns
        /// </summary>
        public int ColumnCount 
        {
            get
            {
                return _prevColumnCount;
            }
        }

        /// <summary>
        /// Collection of rows
        /// </summary>
        private readonly IList<ICsvRowModel> _rows = new List<ICsvRowModel>();
        public IList<ICsvRowModel> Rows {
            get { return _rows; }
        }

        /// <summary>
        /// Add new row from excel row (null or empty cell is added as empty cell)
        /// </summary>
        /// <param name="row">excel row</param>
        /// <returns>
        /// true if we add row to model
        /// otherwize false
        /// </returns>
        public bool AddRow(IRow row)
        {
            bool wasNonEmptyCell = false;
            var values = new List<string>();
            for (int i = 0; i < row.LastCellNum; i++) 
            {
                var cellValue = row.CellToStringCsvValue(i);
                if (cellValue != "")
                {
                    wasNonEmptyCell = true;
                }
                values.Add(cellValue);
            }
            if (wasNonEmptyCell) // ignore cell with all empty cells.
            {
                var rowModel = new CsvRowModel(values);
                NormalizeCsvModel(rowModel);
                Rows.Add(new CsvRowModel(values));
            }
            return wasNonEmptyCell;
        }

        /// <summary>
        /// Make CsvModel rectangle to avoid empty cells.
        /// </summary>
        /// <param name="rowModel"></param>
        private void NormalizeCsvModel(CsvRowModel rowModel)
        {
            if (rowModel.Cells.Count > _prevColumnCount)
            {
                _prevColumnCount = rowModel.Cells.Count;
                foreach (var row in Rows)
                {
                    row.NormalizeRow(_prevColumnCount);
                }
            }
            else
            {
                rowModel.NormalizeRow(_prevColumnCount);
            }
        }

        /// <summary>
        /// Export csv model to file
        /// </summary>
        /// <param name="fileName">the file name of new file</param>
        public void ExportToCsv(string fileName)
        {
            using (var streamWriter = new StreamWriter(fileName))
            {
                using (var csvWriter = new CsvWriter(streamWriter))
                {
                    //write rows 
                    foreach (var row in Rows)
                    {
                        foreach (var cell in row.Cells)
                        {
                            csvWriter.WriteField(cell.Value);
                        }
                        csvWriter.NextRecord(); // move to the next record after each row
                    }
                }
            }
        }
    }
}
