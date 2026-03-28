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
    private int _row = 1;

    [ObservableProperty]
    private double _fontSize = 11;

    [ObservableProperty]
    private bool _isBold;

    [ObservableProperty]
    private bool _isEnabled = true;

    [ObservableProperty]
    private bool _onlyOnFirstPage;

    public static ReportBlockEditorViewModel FromModel(ReportBlock block)
    {
        return new ReportBlockEditorViewModel
        {
            Type = block.Type,
            Text = block.Text,
            Source = block.Source,
            Alignment = block.Alignment,
            Row = Math.Max(1, block.Row),
            FontSize = block.FontSize,
            IsBold = block.IsBold,
            IsEnabled = true,
            OnlyOnFirstPage = block.OnlyOnFirstPage
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
            Row = Math.Max(1, Row),
            FontSize = FontSize,
            IsBold = IsBold,
            OnlyOnFirstPage = OnlyOnFirstPage
        };
    }
}
