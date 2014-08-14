// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ICsvRowModel.cs" company="Orcomp development team">
//   Copyright (c) 2008 - 2014 Orcomp development team. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------


namespace Orc.ExcelToCsv.Models
{
    using System.Collections.Generic;

    /// <summary>
    /// Interface that described the row model
    /// </summary>
    public interface ICsvRowModel
    {
        #region Properties
        /// <summary>
        /// collection of cells 
        /// </summary>
        IList<ICsvCellModel> Cells { get; }
        #endregion

        #region Methods
        /// <summary>
        /// Add cells to make cell count in row equals to columnCount
        /// </summary>
        /// <param name="columnCount">count of expected columns</param>
        void NormalizeRow(int columnCount);
        #endregion
    }
}