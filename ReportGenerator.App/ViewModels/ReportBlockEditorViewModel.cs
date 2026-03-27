using CommunityToolkit.Mvvm.ComponentModel;
using ReportGenerator.Core.Models;

namespace ReportGenerator.App.ViewModels;

public sealed partial class ReportBlockEditorViewModel : ObservableObject
{
    [ObservableProperty]
    private ReportBlockType _type = ReportBlockType.StaticText;

    [ObservableProperty]
    private string _text = string.Empty;

    [ObservableProperty]
    private string _source = string.Empty;

    [ObservableProperty]
    private ReportTextAlignment _alignment = ReportTextAlignment.Left;

    [ObservableProperty]
    private double _fontSize = 11;

    [ObservableProperty]
    private bool _isBold;

    public static ReportBlockEditorViewModel FromModel(ReportBlock block)
    {
        return new ReportBlockEditorViewModel
        {
            Type = block.Type,
            Text = block.Text,
            Source = block.Source,
            Alignment = block.Alignment,
            FontSize = block.FontSize,
            IsBold = block.IsBold
        };
    }

    public ReportBlock ToModel()
    {
        return new ReportBlock
        {
            Type = Type,
            Text = Text,
            Source = Source,
            Alignment = Alignment,
            FontSize = FontSize,
            IsBold = IsBold
        };
    }
}
