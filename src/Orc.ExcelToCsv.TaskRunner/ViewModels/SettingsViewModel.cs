// --------------------------------------------------------------------------------------------------------------------
// <copyright file="SettingsViewModel.cs" company="Orchestra development team">
//   Copyright (c) 2008 - 2014 Orchestra development team. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------


namespace Orc.ExcelToCsv.TaskRunner.ViewModels
{
	using System.Threading.Tasks;
	using Catel;
	using Catel.Fody;
	using Catel.Logging;
	using Catel.MVVM;
	using Catel.Services;
	using Models;
	using Orchestra.Services;

	public class SettingsViewModel : ViewModelBase
    {
        #region Fields
        private readonly ILogControlService _logControlService;
        private readonly IDispatcherService _dispatcherService;
        #endregion

        #region Constructors
        public SettingsViewModel(Settings settings, ILogControlService logControlService, IDispatcherService dispatcherService)
        {
            Argument.IsNotNull(() => settings);
            Argument.IsNotNull(() => logControlService);
            Argument.IsNotNull(() => dispatcherService);

            Settings = settings;
            _logControlService = logControlService;
            _dispatcherService = dispatcherService;
        }
        #endregion

        #region Properties
		[Model]
		[Expose("SpreadsheetPath")]
        [Expose("OutputDirectory")]
        public Settings Settings { get; private set; }
        #endregion

        protected override async Task InitializeAsync()
        {
            await base.InitializeAsync();

            _dispatcherService.BeginInvoke(() => _logControlService.SelectedLevel = LogEvent.Debug | LogEvent.Info);
        }
    }
}