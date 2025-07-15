using System;
using System.Collections.ObjectModel;
using System.IO;
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
                var outlookDirs = Directory.GetDirectories(documents, "*Outlook*", SearchOption.AllDirectories);

                foreach (var dir in outlookDirs)
                {
                    var files = Directory.GetFiles(dir, "*.pst", SearchOption.TopDirectoryOnly);

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
            }
            catch
            {
                // ignore errors
            }
        }
    }
}
