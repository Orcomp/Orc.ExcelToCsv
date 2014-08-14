using System;
using System.IO;

namespace ExcelToCsv.ConsoleTool
{
    using Orc.ExcelToCsv.Models;

    class Program
    {
        private static bool CheckArgs(string[] args, out string errors)
        {
            errors = "";
            try
            {
                if ((args == null) || (args.Length != 2))
                {
                    errors = "Invalid count of arguments";
                }
                else if (!(new FileInfo(args[0]).Exists))
                {
                    errors = "Excel file doesn't exist";
                }
                else if (!(new DirectoryInfo(args[1]).Exists))
                {
                    errors = "Directory doesn't exist";
                }
                else
                {
                    return true;
                }
            }
            catch (Exception exc)
            {
                errors = exc.Message;
            }
            return false;

        }

        private static void Main(string[] args)
        {
            string error;
            if (CheckArgs(args, out error))
            {
                ProcessExcel.CreateCsv(args[0], args[1]);
            }
            else
            {
                Console.WriteLine(error);
                Console.WriteLine("Error have been occured while processing ");
            }
        }
    }
}
