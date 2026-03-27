using System.Windows;

namespace ReportGenerator.App.Views;

public partial class UpdatePromptWindow : Window
{
    public UpdatePromptWindow(string version, string releaseNotes)
    {
        InitializeComponent();
        DataContext = new
        {
            SummaryText = $"Version {version} can be downloaded and applied after the app closes.",
            ReleaseNotes = string.IsNullOrWhiteSpace(releaseNotes) ? "No release notes were published for this version." : releaseNotes
        };
    }

    private void LaterButton_OnClick(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }

    private void InstallButton_OnClick(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
    }
}
