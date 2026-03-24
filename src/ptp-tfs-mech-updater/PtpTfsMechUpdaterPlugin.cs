using System;
using System.Windows;
using VANTAGE.Services.Plugins;

namespace PtpTfsMechUpdater
{
    public class PtpTfsMechUpdaterPlugin : IVantagePlugin
    {
        private IPluginHost? _host;

        public string Id => "ptp-tfs-mech-updater";
        public string Name => "PTP TFS MECH Updater";

        public void Initialize(IPluginHost host)
        {
            _host = host;
            host.AddToolsMenuItem("PTP TFS MECH Updater", OnMenuClick, addSeparatorBefore: true);
        }

        public void Shutdown()
        {
        }

        private async void OnMenuClick()
        {
            if (_host == null) return;

            try
            {
                var dialog = new Microsoft.Win32.OpenFileDialog
                {
                    Title = "Select PTP Report",
                    Filter = "Excel Files (*.xlsx)|*.xlsx",
                    DefaultExt = ".xlsx"
                };

                if (dialog.ShowDialog(_host.MainWindow) != true) return;

                var importer = new PtpImporter(_host);
                await importer.RunAsync(dialog.FileName);
            }
            catch (Exception ex)
            {
                _host.LogError(ex, "PtpTfsMechUpdaterPlugin.OnMenuClick");
                _host.ShowError($"An unexpected error occurred:\n\n{ex.Message}");
            }
        }
    }
}
