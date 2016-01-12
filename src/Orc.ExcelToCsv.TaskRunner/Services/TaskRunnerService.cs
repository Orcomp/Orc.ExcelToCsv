// --------------------------------------------------------------------------------------------------------------------
// <copyright file="TaskRunnerService.cs" company="Orchestra development team">
//   Copyright (c) 2008 - 2014 Orchestra development team. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------


namespace Orc.ExcelToCsv.TaskRunner.Services
{
	using System;
	using System.IO;
	using System.Threading.Tasks;
	using System.Windows;
	using Catel;
	using Catel.Logging;
	using Models;
	using ExcelToCsv.Models;
	using Orchestra.Models;
	using Orchestra.Services;
	using Views;

	public class TaskRunnerService : ITaskRunnerService
    {
        private static readonly ILog Log = LogManager.GetCurrentClassLogger();

        private string _title = "Custom TaskRunner demo";

        public string Title
        {
            get { return _title; }
            set
            {
                _title = value;
                TitleChanged.SafeInvoke(this);
            }
        }

        public event EventHandler TitleChanged;

        public bool ShowCustomizeShortcutsButton { get { return true; }}

        public object GetViewDataContext()
        {
            return new Settings();
        }

        public FrameworkElement GetView()
        {
            return new SettingsView();
        }

        public async Task RunAsync(object dataContext)
        {
            var settings = (Settings) dataContext;

            Log.Info("Running action with the following settings:");
            Log.Indent();
			Log.Info("Spreadsheet path => {0}", settings.SpreadsheetPath);
            Log.Info("Output directory => {0}", settings.OutputDirectory);
            Log.Unindent();

			var outputDirectory = settings.OutputDirectory;
			if (string.IsNullOrWhiteSpace(settings.OutputDirectory))
			{
				outputDirectory = new FileInfo(settings.SpreadsheetPath).DirectoryName;
				Log.Info("Output directory is set to {0}", outputDirectory);
			}

	        try
	        {
				Log.Info("Processing excel spreadsheet");
				ProcessExcel.CreateCsv(settings.SpreadsheetPath, string.Concat(outputDirectory, "\\"));
				Log.Info("Processing completed");
	        }
	        catch (Exception ex)
			{
				Log.Error("Error has been occurred while processing", ex);
	        }
        }

        public Size GetInitialWindowSize()
        {
            return Size.Empty;
        }

        public AboutInfo GetAboutInfo()
        {
            return new AboutInfo();
        }
    }
}