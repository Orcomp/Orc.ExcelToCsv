// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ExcelInvalidException.cs" company="Orcomp development team">
//   Copyright (c) 2008 - 2014 Orcomp development team. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------


namespace Orc.ExcelToCsv.Models
{
    using System;

    public class ExcelInvalidException : Exception
    {
        #region Constructors
        public ExcelInvalidException()
        {
        }

        public ExcelInvalidException(string message) : base(message)
        {
        }
        #endregion
    }
}