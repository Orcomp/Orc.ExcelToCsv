// --------------------------------------------------------------------------------------------------------------------
// <copyright file="IProcessWorkbookResult.cs" company="Orcomp development team">
//   Copyright (c) 2008 - 2014 Orcomp development team. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------


namespace Orc.ExcelToCsv.Models
{
    using System.Collections.Generic;

    public interface IProcessWorkbookResult
    {
        #region Properties
        /// <summary>
        /// collection of sheet to csv models
        /// </summary>
        IList<ISheetToCsvModel> SheetToCsvModels { get; }
        #endregion

        #region Methods
        /// <summary>
        /// Export existing result to csv in working directory
        /// </summary>
        /// <param name="workingDirectory">the working directory where csv files will be placed</param>
        void ExportToCsv(string workingDirectory);
        #endregion
    }
}