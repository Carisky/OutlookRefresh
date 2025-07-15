namespace OutlookRefresh
{
    public class PstFileInfo
    {
        public string Path { get; set; } = string.Empty;
        public double SizeGb { get; set; }
        public string SizeGbFormatted => $"{SizeGb:F1} GB";
        public Microsoft.UI.Xaml.Media.Brush? Background { get; set; }
    }
}
