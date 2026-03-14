using System;
using System.Windows;
using VANTAGE.Services.Plugins;

namespace ConstTfsMechUpdater
{
    public class ConstTfsMechUpdaterPlugin : IVantagePlugin
    {
        private IPluginHost? _host;

        public string Id => "const-tfs-mech-updater";
        public string Name => "CONST TFS MECH Updater";

        public void Initialize(IPluginHost host)
        {
            _host = host;
            host.AddToolsMenuItem("CONST TFS MECH Updater", OnMenuClick, addSeparatorBefore: false);
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
                    Title = "Select CONST Report",
                    Filter = "Excel Files (*.xlsx)|*.xlsx",
                    DefaultExt = ".xlsx"
                };

                if (dialog.ShowDialog(_host.MainWindow) != true) return;

                var importer = new ConstImporter(_host);
                await importer.RunAsync(dialog.FileName);
            }
            catch (Exception ex)
            {
                _host.LogError(ex, "ConstTfsMechUpdaterPlugin.OnMenuClick");
                _host.ShowError($"An unexpected error occurred:\n\n{ex.Message}");
            }
        }
    }
}
