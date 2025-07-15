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

            bool copyStructure = false;
            string? sourcePath = null;

            var checkCopy = new CheckBox { Content = "Copy folder structure" };
            var combo = new ComboBox
            {
                ItemsSource = PstFiles,
                DisplayMemberPath = nameof(PstFileInfo.Path),
                IsEnabled = false
            };
            checkCopy.Checked += (_, _) => combo.IsEnabled = true;
            checkCopy.Unchecked += (_, _) => combo.IsEnabled = false;

            var dialogContent = new StackPanel();
            dialogContent.Children.Add(checkCopy);
            dialogContent.Children.Add(combo);

            var copyDialog = new ContentDialog
            {
                Title = "Copy Folder Structure",
                Content = dialogContent,
                PrimaryButtonText = "OK",
                CloseButtonText = "Cancel"
            };
            InitializeWithWindow.Initialize(copyDialog, hwnd);
            var result = await copyDialog.ShowAsync();
            if (result == ContentDialogResult.Primary)
            {
                copyStructure = checkCopy.IsChecked == true;
                if (copyStructure)
                {
                    sourcePath = (combo.SelectedItem as PstFileInfo)?.Path;
                }
            }

            try
            {
                CreateAndSetDefaultPst(path, copyStructure, sourcePath);
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

        private void CreateAndSetDefaultPst(string path, bool copyStructure, string? sourcePath)
        {
            var outlook = new Outlook.Application();
            Outlook.NameSpace ns = outlook.GetNamespace("MAPI");
            ns.AddStoreEx(path, Outlook.OlStoreType.olStoreUnicode);

            Outlook.Store? newStore = null;
            foreach (Outlook.Store store in ns.Stores)
            {
                if (string.Equals(store.FilePath, path, StringComparison.OrdinalIgnoreCase))
                {
                    newStore = store;
                    break;
                }
            }

            if (copyStructure && !string.IsNullOrEmpty(sourcePath))
            {
                ns.AddStoreEx(sourcePath, Outlook.OlStoreType.olStoreUnicode);

                Outlook.Store? srcStore = null;
                foreach (Outlook.Store store in ns.Stores)
                {
                    if (string.Equals(store.FilePath, sourcePath, StringComparison.OrdinalIgnoreCase))
                    {
                        srcStore = store;
                        break;
                    }
                }

                if (srcStore != null && newStore != null)
                {
                    CopyFolderStructure(srcStore.GetRootFolder(), newStore.GetRootFolder());
                }

                if (srcStore != null)
                {
                    ns.RemoveStore(srcStore.GetRootFolder());
                }
            }

            // Attempt to set new store as default for the active account.
            // Outlook object model does not expose a direct way to change the
            // default delivery store. This placeholder shows where such logic
            // would be implemented using extended MAPI or other means.
        }

        private static void CopyFolderStructure(Outlook.MAPIFolder source, Outlook.MAPIFolder target)
        {
            foreach (Outlook.MAPIFolder child in source.Folders)
            {
                var newChild = target.Folders.Add(child.Name, child.DefaultItemType);
                CopyFolderStructure(child, newChild);
            }
        }
    }
}
