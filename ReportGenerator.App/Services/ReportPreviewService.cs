using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Threading;
using ReportGenerator.App.Controls;
using ReportGenerator.App.Models;
using ReportGenerator.Core.Models;
using ReportGenerator.Core.Utilities;

namespace ReportGenerator.App.Services;

public sealed class ReportPreviewService
{
    private const int PreviewYieldInterval = 8;

    public FixedDocument CreateDocument(PagedReport report)
    {
        ArgumentNullException.ThrowIfNull(report);

        return CreateDocument(report, 0, report.Pages.Count - 1);
    }

    public FixedDocument CreateDocument(PagedReport report, int startPageIndex, int endPageIndex)
    {
        ArgumentNullException.ThrowIfNull(report);

        if (report.Pages.Count == 0)
        {
            return CreateEmptyDocument(report);
        }

        if (startPageIndex < 0 || startPageIndex >= report.Pages.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(startPageIndex));
        }

        if (endPageIndex < startPageIndex || endPageIndex >= report.Pages.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(endPageIndex));
        }

        var (pageWidthMm, pageHeightMm) = PageMeasurementHelper.GetPageDimensionsMillimeters(report.PageSettings);
        var pageSize = new Size(
            PageMeasurementHelper.MillimetersToDip(pageWidthMm),
            PageMeasurementHelper.MillimetersToDip(pageHeightMm));

        var document = CreateDocumentShell(pageSize);

        for (var pageIndex = startPageIndex; pageIndex <= endPageIndex; pageIndex++)
        {
            var fixedPage = BuildPageVisual(report, report.Pages[pageIndex], pageSize);
            var pageContent = new PageContent();
            ((IAddChild)pageContent).AddChild(fixedPage);
            document.Pages.Add(pageContent);
        }

        return document;
    }

    public async Task<FixedDocument> CreateDocumentAsync(
        PagedReport report,
        IProgress<PreviewGenerationProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(report);

        var (pageWidthMm, pageHeightMm) = PageMeasurementHelper.GetPageDimensionsMillimeters(report.PageSettings);
        var pageSize = new Size(
            PageMeasurementHelper.MillimetersToDip(pageWidthMm),
            PageMeasurementHelper.MillimetersToDip(pageHeightMm));

        var document = CreateDocumentShell(pageSize);

        await Dispatcher.Yield(DispatcherPriority.Background);

        for (var pageIndex = 0; pageIndex < report.Pages.Count; pageIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var fixedPage = BuildPageVisual(report, report.Pages[pageIndex], pageSize);
            var pageContent = new PageContent();
            ((IAddChild)pageContent).AddChild(fixedPage);
            document.Pages.Add(pageContent);

            progress?.Report(new PreviewGenerationProgress(pageIndex + 1, report.Pages.Count));

            if ((pageIndex + 1) % PreviewYieldInterval == 0)
            {
                await Dispatcher.Yield(DispatcherPriority.Background);
            }
        }

        return document;
    }

    private static FixedPage BuildPageVisual(PagedReport report, ReportPage reportPage, Size pageSize)
    {
        var fixedPage = new FixedPage
        {
            Width = pageSize.Width,
            Height = pageSize.Height,
            Background = Brushes.White,
            Focusable = false,
            // Treat preview pages as rendered output, not interactive text surfaces.
            IsHitTestVisible = false
        };

        var pageVisual = new PreviewPageVisual(report, reportPage, pageSize);
        FixedPage.SetLeft(pageVisual, 0);
        FixedPage.SetTop(pageVisual, 0);
        fixedPage.Children.Add(pageVisual);
        return fixedPage;
    }

    private static FixedDocument CreateEmptyDocument(PagedReport report)
    {
        var (pageWidthMm, pageHeightMm) = PageMeasurementHelper.GetPageDimensionsMillimeters(report.PageSettings);
        var pageSize = new Size(
            PageMeasurementHelper.MillimetersToDip(pageWidthMm),
            PageMeasurementHelper.MillimetersToDip(pageHeightMm));
        return CreateDocumentShell(pageSize);
    }

    private static FixedDocument CreateDocumentShell(Size pageSize)
    {
        return new FixedDocument
        {
            DocumentPaginator =
            {
                PageSize = pageSize
            }
        };
    }
}
