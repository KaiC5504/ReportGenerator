using CommunityToolkit.Mvvm.ComponentModel;
using ReportGenerator.Core.Models;

namespace ReportGenerator.App.ViewModels;

public sealed partial class DetailColumnEditorViewModel : ObservableObject
{
    [ObservableProperty]
    private string _headerText = "Column";

    [ObservableProperty]
    private string _source = string.Empty;

    [ObservableProperty]
    private double _widthWeight = 1;

    [ObservableProperty]
    private ReportTextAlignment _alignment = ReportTextAlignment.Left;

    public static DetailColumnEditorViewModel FromModel(DetailColumnDefinition column)
    {
        return new DetailColumnEditorViewModel
        {
            HeaderText = column.HeaderText,
            Source = column.Source,
            WidthWeight = column.WidthWeight,
            Alignment = column.Alignment
        };
    }

    public DetailColumnDefinition ToModel()
    {
        return new DetailColumnDefinition
        {
            HeaderText = HeaderText,
            Source = Source,
            WidthWeight = WidthWeight,
            Alignment = Alignment
        };
    }
}
