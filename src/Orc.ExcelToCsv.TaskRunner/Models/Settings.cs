// --------------------------------------------------------------------------------------------------------------------
// <copyright file="Settings.cs" company="Orchestra development team">
//   Copyright (c) 2008 - 2014 Orchestra development team. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------


namespace Orc.ExcelToCsv.TaskRunner.Models
{
	using System.Collections.Generic;
	using System.IO;
	using Catel.Data;

	public class Settings : ModelBase
    {
        public string OutputDirectory { get; set; }

		public string SpreadsheetPath { get; set; }

        protected override void ValidateFields(List<IFieldValidationResult> validationResults)
        {
            base.ValidateFields(validationResults);
            if (string.IsNullOrWhiteSpace(SpreadsheetPath))
            {
				validationResults.Add(FieldValidationResult.CreateError("SpreadsheetPath", "Spreadsheet path is required"));
            }

			if (!string.IsNullOrWhiteSpace(SpreadsheetPath) && !(new FileInfo(SpreadsheetPath).Exists))
			{
				validationResults.Add(FieldValidationResult.CreateError("SpreadsheetPath", "Excel file doesn't exist"));
			}

			if (!string.IsNullOrWhiteSpace(OutputDirectory) && !(new DirectoryInfo(OutputDirectory).Exists))
			{
				throw new DirectoryNotFoundException("There is no output directory");
			}
        }
    }
}