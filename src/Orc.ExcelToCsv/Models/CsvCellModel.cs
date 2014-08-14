// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CsvCellModel.cs" company="Orcomp development team">
//   Copyright (c) 2008 - 2014 Orcomp development team. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------


namespace Orc.ExcelToCsv.Models
{
    /// <summary>
    /// Class that represent model of cell
    /// </summary>
    internal class CsvCellModel : ICsvCellModel
    {
        #region Fields
        /// <summary>
        /// Value of cell 
        /// </summary>
        public string _value;
        #endregion

        #region Constructors
        /// <summary>
        /// Initialize the cell model
        /// </summary>
        /// <param name="value">value of cell</param>
        public CsvCellModel(string value)
        {
            _value = value;
        }
        #endregion

        #region ICsvCellModel Members
        public string Value
        {
            get { return _value; }
        }
        #endregion
    }
}