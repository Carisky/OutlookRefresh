using System;
using System.Collections.ObjectModel;
using System.IO;
using Windows.Storage.Pickers;
using Windows.Storage;
using WinRT.Interop;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using Windows.Foundation;
using Windows.Foundation.Collections;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace OutlookRefresh
{
    /// <summary>
    /// An empty window that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainWindow : Window
    {
        public ObservableCollection<PstFileInfo> PstFiles { get; } = new();

        public MainWindow()
        {
            InitializeComponent();
            LoadPstFiles();
        }

        private void LoadPstFiles()
        {
            string documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            try
            {
                var files = Directory.EnumerateFiles(
                    documents,
                    "*.pst",
                    new EnumerationOptions { IgnoreInaccessible = true, RecurseSubdirectories = true });

                foreach (var file in files)
                {
                    double sizeGb = new FileInfo(file).Length / (1024.0 * 1024 * 1024);
                    var color = sizeGb < 35 ? Microsoft.UI.Colors.LightGreen :
                                 sizeGb < 45 ? Microsoft.UI.Colors.Orange :
                                 Microsoft.UI.Colors.Red;
                    var brush = new SolidColorBrush(color);

                    PstFiles.Add(new PstFileInfo { Path = file, SizeGb = sizeGb, Background = brush });
                }
            }
            catch
            {
                // ignore errors
            }
        }

        private async void CreatePstClicked(object sender, RoutedEventArgs e)
        {
            var picker = new FileSavePicker();
            var hwnd = WindowNative.GetWindowHandle(this);
            InitializeWithWindow.Initialize(picker, hwnd);
            picker.FileTypeChoices.Add("Outlook Data File", new[] { ".pst" });
            picker.SuggestedFileName = "NewDataFile";
            StorageFile? file = await picker.PickSaveFileAsync();
            if (file == null)
                return;

            var path = file.Path;
            if (!path.EndsWith(".pst", StringComparison.OrdinalIgnoreCase))
                path += ".pst";

            try
            {
                CreateAndSetDefaultPst(path);
            }
            catch (Exception ex)
            {
                var dialog = new ContentDialog
                {
                    Title = "Error",
                    Content = ex.Message,
                    CloseButtonText = "OK"
                };
                InitializeWithWindow.Initialize(dialog, hwnd);
                await dialog.ShowAsync();
            }

            PstFiles.Clear();
            LoadPstFiles();
        }

        private void CreateAndSetDefaultPst(string path)
        {
            var outlook = new Outlook.Application();
            Outlook.NameSpace ns = outlook.GetNamespace("MAPI");
            ns.AddStoreEx(path, Outlook.OlStoreType.olStoreUnicode);

            // Attempt to set new store as default for the active account.
            // Outlook object model does not expose a direct way to change the
            // default delivery store. This placeholder shows where such logic
            // would be implemented using extended MAPI or other means.
        }
    }
}
