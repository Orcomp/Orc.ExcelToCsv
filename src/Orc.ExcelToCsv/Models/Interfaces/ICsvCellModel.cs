// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ICsvCellModel.cs" company="Orcomp development team">
//   Copyright (c) 2008 - 2014 Orcomp development team. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------


namespace Orc.ExcelToCsv.Models
{
    /// <summary>
    /// Interface that described model of cell
    /// </summary>
    public interface ICsvCellModel
    {
        #region Properties
        /// <summary>
        /// Value of cell 
        /// </summary>
        string Value { get; }
        #endregion
    }
}