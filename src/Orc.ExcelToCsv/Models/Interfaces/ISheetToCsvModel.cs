// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ISheetToCsvModel.cs" company="Orcomp development team">
//   Copyright (c) 2008 - 2014 Orcomp development team. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------


namespace Orc.ExcelToCsv.Models
{
    /// <summary>
    /// Interface that described sheet to csv model
    /// </summary>
    public interface ISheetToCsvModel
    {
        #region Properties
        /// <summary>
        /// CsvModel
        /// </summary>
        ICsvModel CsvModel { get; }

        /// <summary>
        /// name of sheet
        /// </summary>
        string SheetName { get; }
        #endregion
    }
}