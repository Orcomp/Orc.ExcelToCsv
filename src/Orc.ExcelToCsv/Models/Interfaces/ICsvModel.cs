// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ICsvModel.cs" company="Orcomp development team">
//   Copyright (c) 2008 - 2014 Orcomp development team. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------


namespace Orc.ExcelToCsv.Models
{
    using System.Collections.Generic;

    /// <summary>
    /// Interface that described csv model
    /// </summary>
    public interface ICsvModel
    {
        #region Properties
        /// <summary>
        /// Count of columns
        /// </summary>
        int ColumnCount { get; }

        /// <summary>
        /// Collection of rows
        /// </summary>
        IList<ICsvRowModel> Rows { get; }
        #endregion

        #region Methods
        /// <summary>
        /// Export csv model to file
        /// </summary>
        /// <param name="fileName">the file name of new file</param>
        void ExportToCsv(string fileName);
        #endregion
    }
}