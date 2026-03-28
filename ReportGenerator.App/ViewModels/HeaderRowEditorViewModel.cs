using CommunityToolkit.Mvvm.ComponentModel;
using ReportGenerator.Core.Models;

namespace ReportGenerator.App.ViewModels;

public sealed partial class HeaderRowEditorViewModel : ObservableObject
{
    [ObservableProperty]
    private int _rowNumber = 1;

    [ObservableProperty]
    private bool _onlyOnFirstPage;

    public ReportBlockEditorViewModel LeftBlock { get; } = CreateSlot(ReportTextAlignment.Left);

    public ReportBlockEditorViewModel CenterBlock { get; } = CreateSlot(ReportTextAlignment.Center);

    public ReportBlockEditorViewModel RightBlock { get; } = CreateSlot(ReportTextAlignment.Right);

    public static HeaderRowEditorViewModel CreateEmpty(int rowNumber)
    {
        var row = new HeaderRowEditorViewModel
        {
            RowNumber = Math.Max(1, rowNumber)
        };

        row.LeftBlock.IsEnabled = true;
        return row;
    }

    public static HeaderRowEditorViewModel FromBlocks(IEnumerable<ReportBlock> blocks)
    {
        var orderedBlocks = blocks
            .Where(block => block.Row >= 1)
            .OrderBy(block => GetAlignmentOrder(block.Alignment))
            .ToArray();

        var row = new HeaderRowEditorViewModel();
        if (orderedBlocks.Length > 0)
        {
            row.RowNumber = Math.Max(1, orderedBlocks[0].Row);
            row.OnlyOnFirstPage = orderedBlocks[0].OnlyOnFirstPage;
        }

        ApplyBlock(row.LeftBlock, orderedBlocks.FirstOrDefault(block => block.Alignment == ReportTextAlignment.Left), ReportTextAlignment.Left);
        ApplyBlock(row.CenterBlock, orderedBlocks.FirstOrDefault(block => block.Alignment == ReportTextAlignment.Center), ReportTextAlignment.Center);
        ApplyBlock(row.RightBlock, orderedBlocks.FirstOrDefault(block => block.Alignment == ReportTextAlignment.Right), ReportTextAlignment.Right);

        return row;
    }

    public IReadOnlyList<ReportBlock> ToBlocks()
    {
        var blocks = new List<ReportBlock>();
        AppendEnabledBlock(blocks, LeftBlock);
        AppendEnabledBlock(blocks, CenterBlock);
        AppendEnabledBlock(blocks, RightBlock);
        return blocks;
    }

    private void AppendEnabledBlock(ICollection<ReportBlock> blocks, ReportBlockEditorViewModel slot)
    {
        if (!slot.IsEnabled)
        {
            return;
        }

        var block = slot.ToModel();
        block.Row = RowNumber;
        block.Alignment = slot.Alignment;
        block.OnlyOnFirstPage = OnlyOnFirstPage;
        blocks.Add(block);
    }

    private static ReportBlockEditorViewModel CreateSlot(ReportTextAlignment alignment)
    {
        return new ReportBlockEditorViewModel
        {
            Alignment = alignment,
            IsEnabled = false
        };
    }

    private static void ApplyBlock(ReportBlockEditorViewModel slot, ReportBlock? block, ReportTextAlignment alignment)
    {
        slot.Alignment = alignment;

        if (block is null)
        {
            slot.IsEnabled = false;
            slot.Type = ReportBlockType.StaticText;
            slot.Text = string.Empty;
            slot.Source = string.Empty;
            slot.FontSize = 11;
            slot.IsBold = false;
            slot.Row = 1;
            slot.OnlyOnFirstPage = false;
            return;
        }

        slot.Type = block.Type;
        slot.Text = block.Text;
        slot.Source = block.Source;
        slot.Row = Math.Max(1, block.Row);
        slot.FontSize = block.FontSize;
        slot.IsBold = block.IsBold;
        slot.IsEnabled = true;
        slot.OnlyOnFirstPage = block.OnlyOnFirstPage;
    }

    private static int GetAlignmentOrder(ReportTextAlignment alignment)
    {
        return alignment switch
        {
            ReportTextAlignment.Left => 0,
            ReportTextAlignment.Center => 1,
            ReportTextAlignment.Right => 2,
            _ => 3
        };
    }
}
