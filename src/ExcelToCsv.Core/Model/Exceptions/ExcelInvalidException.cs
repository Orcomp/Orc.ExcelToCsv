namespace ExcelToCsv.Core.Model.Exceptions
{
    using System;

    public class ExcelInvalidException : Exception
    {
        public ExcelInvalidException() 
        {
        }

        public ExcelInvalidException(string message) : base(message)
        {

        }
    }
}
