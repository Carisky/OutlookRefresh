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
            string home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            try
            {
                var files = Directory.GetFiles(home, "*.pst", SearchOption.AllDirectories);
                foreach (var file in files)
                {
                    double sizeGb = new FileInfo(file).Length / (1024.0 * 1024 * 1024);
                    var color = sizeGb < 35 ? Windows.UI.Color.FromArgb(128, 0, 255, 0) :
                                 sizeGb < 45 ? Windows.UI.Color.FromArgb(128, 255, 165, 0) :
                                 Windows.UI.Color.FromArgb(128, 255, 0, 0);
                    var brush = new SolidColorBrush(color);
                    PstFiles.Add(new PstFileInfo { Path = file, SizeGb = sizeGb, Background = brush });
                }
            }
            catch
            {
                // ignore errors scanning directories
            }
        }
    }
}
