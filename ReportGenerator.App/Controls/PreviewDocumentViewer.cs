using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Media3D;

namespace ReportGenerator.App.Controls;

public sealed class PreviewDocumentViewer : DocumentViewer
{
    private static readonly PropertyInfo? IsSelectionEnabledProperty = typeof(DocumentViewerBase)
        .GetProperty("IsSelectionEnabled", BindingFlags.Instance | BindingFlags.NonPublic);

    public PreviewDocumentViewer()
    {
        Focusable = false;
        IsTabStop = false;
        Cursor = Cursors.Arrow;
        ForceCursor = true;
    }

    public override void OnApplyTemplate()
    {
        base.OnApplyTemplate();
        DisableTextSelection();
    }

    protected override void OnPreviewMouseLeftButtonDown(MouseButtonEventArgs e)
    {
        if (ShouldSuppressPageClick(e.OriginalSource as DependencyObject))
        {
            DisableTextSelection();
            e.Handled = true;
            return;
        }

        base.OnPreviewMouseLeftButtonDown(e);
    }

    private void DisableTextSelection()
    {
        IsSelectionEnabledProperty?.SetValue(this, false);
    }

    private static bool ShouldSuppressPageClick(DependencyObject? source)
    {
        for (var current = source; current is not null; current = GetParent(current))
        {
            if (current is ScrollBar or Thumb or ButtonBase)
            {
                return false;
            }
        }

        return source is not null;
    }

    private static DependencyObject? GetParent(DependencyObject source)
    {
        return source switch
        {
            Visual or Visual3D => VisualTreeHelper.GetParent(source),
            FrameworkContentElement frameworkContentElement => frameworkContentElement.Parent,
            _ => LogicalTreeHelper.GetParent(source)
        };
    }
}
