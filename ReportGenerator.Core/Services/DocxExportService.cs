using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ReportGenerator.Core.Abstractions;
using ReportGenerator.Core.Models;
using ReportGenerator.Core.Utilities;

namespace ReportGenerator.Core.Services;

public sealed class DocxExportService : IDocxExportService
{
    public Task ExportAsync(PagedReport report, string outputPath, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(report);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);

        return Task.Run(() => ExportInternal(report, outputPath, cancellationToken), cancellationToken);
    }

    private static void ExportInternal(PagedReport report, string outputPath, CancellationToken cancellationToken)
    {
        var directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        if (File.Exists(outputPath))
        {
            File.Delete(outputPath);
        }

        using var wordDocument = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document);
        var mainPart = wordDocument.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var body = mainPart.Document.Body!;

        var firstPageHeaderBlocks = report.Pages.FirstOrDefault()?.HeaderBlocks ?? Array.Empty<ReportBlockContent>();
        var defaultHeaderBlocks = report.Pages.Skip(1).FirstOrDefault()?.HeaderBlocks ?? firstPageHeaderBlocks;
        var footerBlocks = report.Pages.FirstOrDefault()?.FooterBlocks ?? Array.Empty<ReportBlockContent>();
        var useDistinctFirstPageHeader = report.Pages.Count > 1 && !AreHeaderBlocksEquivalent(firstPageHeaderBlocks, defaultHeaderBlocks);

        var defaultHeaderPart = mainPart.AddNewPart<HeaderPart>();
        defaultHeaderPart.Header = CreateHeader(report.PageSettings, defaultHeaderBlocks);

        HeaderPart? firstHeaderPart = null;
        if (useDistinctFirstPageHeader)
        {
            firstHeaderPart = mainPart.AddNewPart<HeaderPart>();
            firstHeaderPart.Header = CreateHeader(report.PageSettings, firstPageHeaderBlocks);
        }

        var defaultFooterPart = mainPart.AddNewPart<FooterPart>();
        defaultFooterPart.Footer = CreateFooter(footerBlocks);

        FooterPart? firstFooterPart = null;
        if (useDistinctFirstPageHeader)
        {
            firstFooterPart = mainPart.AddNewPart<FooterPart>();
            firstFooterPart.Footer = CreateFooter(footerBlocks);
        }

        foreach (var page in report.Pages)
        {
            cancellationToken.ThrowIfCancellationRequested();
            body.Append(CreatePageTable(report, page));
            AppendFooterBlocks(body, page.FooterBlocks);

            if (page.PageNumber < page.TotalPages)
            {
                body.Append(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
            }
        }

        body.Append(CreateSectionProperties(
            report.PageSettings,
            mainPart.GetIdOfPart(defaultHeaderPart),
            mainPart.GetIdOfPart(defaultFooterPart),
            firstHeaderPart is null ? null : mainPart.GetIdOfPart(firstHeaderPart),
            firstFooterPart is null ? null : mainPart.GetIdOfPart(firstFooterPart)));

        mainPart.Document.Save();
    }

    private static Header CreateHeader(PageSettings settings, IReadOnlyList<ReportBlockContent> blocks)
    {
        var header = new Header();

        if (blocks.Count == 0)
        {
            header.Append(CreateEmptyParagraph());
            return header;
        }

        header.Append(CreateHeaderTable(settings, blocks));
        return header;
    }

    private static Footer CreateFooter(IReadOnlyList<ReportBlockContent> blocks)
    {
        return new Footer(CreateEmptyParagraph());
    }

    private static void AppendFooterBlocks(Body body, IReadOnlyList<ReportBlockContent> blocks)
    {
        foreach (var block in blocks)
        {
            body.Append(CreateDynamicTextParagraph(block));
        }
    }

    private static Table CreatePageTable(PagedReport report, ReportPage page)
    {
        var table = new Table();
        var totalWidthTwips = GetPrintableWidthTwips(report.PageSettings);
        var columnWidths = CalculateColumnWidthsTwips(report.Columns, totalWidthTwips);
        var contentRowHeightTwips = GetContentRowHeightTwips(report);

        var properties = new TableProperties(
            new TableWidth { Width = totalWidthTwips.ToString(), Type = TableWidthUnitValues.Dxa },
            new TableLayout { Type = TableLayoutValues.Fixed },
            new TableCellMarginDefault(
                new TopMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                new BottomMargin { Width = "0", Type = TableWidthUnitValues.Dxa }),
            new TableBorders(
                new TopBorder { Val = BorderValues.Nil },
                new BottomBorder { Val = page.Rows.Count > 0 ? BorderValues.Single : BorderValues.Nil, Size = 8 },
                new LeftBorder { Val = BorderValues.Nil },
                new RightBorder { Val = BorderValues.Nil },
                new InsideHorizontalBorder { Val = BorderValues.Nil },
                new InsideVerticalBorder { Val = BorderValues.Nil }));

        table.AppendChild(properties);
        table.Append(CreateTableGrid(columnWidths));

        var headerRow = new TableRow(new TableRowProperties(
            new TableRowHeight
            {
                HeightType = HeightRuleValues.AtLeast,
                Val = (UInt32Value)(uint)Math.Max(1, GetHeaderRowHeightTwips(report))
            }));
        for (var index = 0; index < report.Columns.Count; index++)
        {
            var column = report.Columns[index];
            headerRow.Append(CreateDetailHeaderCell(
                column.HeaderText,
                column.Alignment,
                columnWidths[index],
                report.DetailHeaderFontSize));
        }

        table.Append(headerRow);

        foreach (var row in page.Rows)
        {
            if (row.IsSpacer)
            {
                table.Append(CreateSpacerRow(report, columnWidths, row.HeightFactor));
                continue;
            }

            var tableRow = CreateDetailRow(contentRowHeightTwips);
            for (var index = 0; index < row.Cells.Count; index++)
            {
                var cell = row.Cells[index];
                tableRow.Append(CreateDetailCell(
                    cell.Text,
                    cell.Alignment,
                    columnWidths[index],
                    report.DetailContentFontSize));
            }

            table.Append(tableRow);
        }

        return table;
    }

    private static Table CreateHeaderTable(PageSettings settings, IReadOnlyList<ReportBlockContent> blocks)
    {
        var totalWidthTwips = GetPrintableWidthTwips(settings);
        var columnWidths = CalculateEqualWidthColumns(totalWidthTwips, 3);
        var table = new Table();

        table.AppendChild(new TableProperties(
            new TableWidth { Width = totalWidthTwips.ToString(), Type = TableWidthUnitValues.Dxa },
            new TableLayout { Type = TableLayoutValues.Fixed },
            new TableBorders(
                new TopBorder { Val = BorderValues.Nil },
                new BottomBorder { Val = BorderValues.Nil },
                new LeftBorder { Val = BorderValues.Nil },
                new RightBorder { Val = BorderValues.Nil },
                new InsideHorizontalBorder { Val = BorderValues.Nil },
                new InsideVerticalBorder { Val = BorderValues.Nil })));
        table.Append(CreateTableGrid(columnWidths));

        foreach (var rowBlocks in blocks.GroupBy(block => block.Row).OrderBy(group => group.Key))
        {
            var tableRow = new TableRow();
            tableRow.Append(CreateHeaderCell(rowBlocks.FirstOrDefault(block => block.Alignment == ReportTextAlignment.Left), ReportTextAlignment.Left, columnWidths[0]));
            tableRow.Append(CreateHeaderCell(rowBlocks.FirstOrDefault(block => block.Alignment == ReportTextAlignment.Center), ReportTextAlignment.Center, columnWidths[1]));
            tableRow.Append(CreateHeaderCell(rowBlocks.FirstOrDefault(block => block.Alignment == ReportTextAlignment.Right), ReportTextAlignment.Right, columnWidths[2]));
            table.Append(tableRow);
        }

        return table;
    }

    private static TableGrid CreateTableGrid(IEnumerable<int> columnWidths)
    {
        var tableGrid = new TableGrid();
        foreach (var width in columnWidths)
        {
            tableGrid.Append(new GridColumn { Width = width.ToString() });
        }

        return tableGrid;
    }

    private static TableCell CreateHeaderCell(ReportBlockContent? block, ReportTextAlignment alignment, int widthTwips)
    {
        var cell = new TableCell();
        cell.Append(
            CreateCellProperties(widthTwips),
            block is null
                ? CreateEmptyParagraph(alignment)
                : CreateDynamicTextParagraph(block, alignment));

        return cell;
    }

    private static TableCell CreateDetailHeaderCell(string text, ReportTextAlignment alignment, int widthTwips, double fontSize)
    {
        var cell = new TableCell();
        cell.Append(
            CreateCellProperties(
                widthTwips,
                new TableCellBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 8 },
                    new BottomBorder { Val = BorderValues.Single, Size = 8 },
                    new LeftBorder { Val = BorderValues.Nil },
                    new RightBorder { Val = BorderValues.Nil })),
            CreateTextParagraph(text, alignment, fontSize, isBold: true));

        return cell;
    }

    private static TableCell CreateDetailCell(string text, ReportTextAlignment alignment, int widthTwips, double fontSize)
    {
        var cell = new TableCell();
        cell.Append(
            CreateCellProperties(widthTwips),
            CreateTextParagraph(text, alignment, fontSize, isBold: false));
        return cell;
    }

    private static TableRow CreateDetailRow(int contentRowHeightTwips)
    {
        return new TableRow(new TableRowProperties(
            new TableRowHeight
            {
                HeightType = HeightRuleValues.AtLeast,
                Val = (UInt32Value)(uint)Math.Max(1, contentRowHeightTwips)
            }));
    }

    private static TableRow CreateSpacerRow(PagedReport report, IReadOnlyList<int> columnWidths, double heightFactor)
    {
        var contentRowHeightTwips = GetContentRowHeightTwips(report);
        var tableRow = new TableRow(new TableRowProperties(
            new TableRowHeight
            {
                HeightType = HeightRuleValues.Exact,
                Val = (UInt32Value)(uint)Math.Max(1, Math.Round(contentRowHeightTwips * Math.Max(0, heightFactor), MidpointRounding.AwayFromZero))
            }));

        foreach (var widthTwips in columnWidths)
        {
            tableRow.Append(CreateSpacerCell(widthTwips));
        }

        return tableRow;
    }

    private static TableCell CreateSpacerCell(int widthTwips)
    {
        var cell = new TableCell();
        cell.Append(CreateCellProperties(
            widthTwips,
            new TableCellBorders(
                new TopBorder { Val = BorderValues.Nil },
                new BottomBorder { Val = BorderValues.Nil },
                new LeftBorder { Val = BorderValues.Nil },
                new RightBorder { Val = BorderValues.Nil })),
            CreateEmptyParagraph());
        return cell;
    }

    private static TableCellProperties CreateCellProperties(int widthTwips, TableCellBorders? borders = null)
    {
        var properties = new TableCellProperties(
            new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = widthTwips.ToString() },
            new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center });

        if (borders is not null)
        {
            properties.Append(borders);
        }

        return properties;
    }

    private static Paragraph CreateTextParagraph(string text, ReportTextAlignment alignment, double fontSize, bool isBold)
    {
        var paragraph = new Paragraph(new ParagraphProperties(
            new Justification { Val = ToJustification(alignment) },
            new SpacingBetweenLines { Before = "0", After = "0" }));

        paragraph.Append(new Run(
            CreateRunProperties(fontSize, isBold),
            new Text(text ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve }));
        return paragraph;
    }

    private static Paragraph CreateDynamicTextParagraph(ReportBlockContent block, ReportTextAlignment? alignmentOverride = null)
    {
        var paragraph = new Paragraph(new ParagraphProperties(
            new Justification { Val = ToJustification(alignmentOverride ?? block.Alignment) },
            new SpacingBetweenLines { Before = "0", After = "0" }));

        if (block.Type == ReportBlockType.PageNumber)
        {
            AppendTextAndFields(paragraph, CreateRunProperties(block.FontSize, block.IsBold), GetPageNumberTemplate(block));
            return paragraph;
        }

        paragraph.Append(new Run(
            CreateRunProperties(block.FontSize, block.IsBold),
            new Text(block.Text ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve }));
        return paragraph;
    }

    private static Paragraph CreateEmptyParagraph(ReportTextAlignment alignment = ReportTextAlignment.Left)
    {
        return new Paragraph(new ParagraphProperties(
            new Justification { Val = ToJustification(alignment) },
            new SpacingBetweenLines { Before = "0", After = "0" }));
    }

    private static void AppendTextAndFields(Paragraph paragraph, RunProperties runProperties, string template)
    {
        var remaining = template;
        while (remaining.Length > 0)
        {
            var pageIndex = remaining.IndexOf("{page}", StringComparison.Ordinal);
            var totalIndex = remaining.IndexOf("{totalPages}", StringComparison.Ordinal);
            var nextIndex = pageIndex >= 0 && totalIndex >= 0 ? Math.Min(pageIndex, totalIndex) : Math.Max(pageIndex, totalIndex);

            if (nextIndex < 0)
            {
                paragraph.Append(new Run(runProperties.CloneNode(true), new Text(remaining) { Space = SpaceProcessingModeValues.Preserve }));
                break;
            }

            if (nextIndex > 0)
            {
                paragraph.Append(new Run(runProperties.CloneNode(true), new Text(remaining[..nextIndex]) { Space = SpaceProcessingModeValues.Preserve }));
            }

            if (pageIndex == nextIndex)
            {
                paragraph.Append(CreateField(" PAGE ", runProperties));
                remaining = remaining[(nextIndex + "{page}".Length)..];
            }
            else
            {
                paragraph.Append(CreateField(" NUMPAGES ", runProperties));
                remaining = remaining[(nextIndex + "{totalPages}".Length)..];
            }
        }
    }

    private static SimpleField CreateField(string instruction, RunProperties runProperties)
    {
        var field = new SimpleField { Instruction = instruction };
        field.Append(new Run(runProperties.CloneNode(true)));
        return field;
    }

    private static RunProperties CreateRunProperties(double fontSize, bool isBold)
    {
        var runProperties = new RunProperties(
            new RunFonts { Ascii = "Segoe UI", HighAnsi = "Segoe UI" },
            new FontSize { Val = ((int)Math.Round(fontSize * 2, MidpointRounding.AwayFromZero)).ToString() });

        if (isBold)
        {
            runProperties.Append(new Bold());
        }

        return runProperties;
    }

    private static int GetHeaderRowHeightTwips(PagedReport report)
    {
        return PageMeasurementHelper.DipToTwips(PageMeasurementHelper.CalculateTableCellMinHeightDip(report.DetailHeaderFontSize));
    }

    private static int GetContentRowHeightTwips(PagedReport report)
    {
        return PageMeasurementHelper.DipToTwips(PageMeasurementHelper.CalculateContentRowHeightDip(
            report.DetailContentFontSize,
            report.DetailContentRowSpacing));
    }

    private static string GetPageNumberTemplate(ReportBlockContent block)
    {
        return string.IsNullOrWhiteSpace(block.TemplateText)
            ? "Page {page} of {totalPages}"
            : block.TemplateText;
    }

    private static SectionProperties CreateSectionProperties(
        PageSettings settings,
        string defaultHeaderPartId,
        string defaultFooterPartId,
        string? firstHeaderPartId,
        string? firstFooterPartId)
    {
        var (pageWidthMm, pageHeightMm) = PageMeasurementHelper.GetPageDimensionsMillimeters(settings);
        var sectionProperties = new SectionProperties();

        if (!string.IsNullOrWhiteSpace(firstHeaderPartId))
        {
            sectionProperties.Append(new HeaderReference { Type = HeaderFooterValues.First, Id = firstHeaderPartId });
        }

        sectionProperties.Append(new HeaderReference { Type = HeaderFooterValues.Default, Id = defaultHeaderPartId });

        if (!string.IsNullOrWhiteSpace(firstFooterPartId))
        {
            sectionProperties.Append(new FooterReference { Type = HeaderFooterValues.First, Id = firstFooterPartId });
        }

        sectionProperties.Append(new FooterReference { Type = HeaderFooterValues.Default, Id = defaultFooterPartId });

        if (!string.IsNullOrWhiteSpace(firstHeaderPartId))
        {
            sectionProperties.Append(new TitlePage());
        }

        sectionProperties.Append(
            new PageSize
            {
                Width = (UInt32Value)(uint)PageMeasurementHelper.MillimetersToTwips(pageWidthMm),
                Height = (UInt32Value)(uint)PageMeasurementHelper.MillimetersToTwips(pageHeightMm)
            },
            new PageMargin
            {
                Top = PageMeasurementHelper.MillimetersToTwips(settings.MarginTopMm),
                Right = (UInt32Value)(uint)PageMeasurementHelper.MillimetersToTwips(settings.MarginRightMm),
                Bottom = PageMeasurementHelper.MillimetersToTwips(settings.MarginBottomMm),
                Left = (UInt32Value)(uint)PageMeasurementHelper.MillimetersToTwips(settings.MarginLeftMm),
                Header = 425U,
                Footer = 425U
            });

        return sectionProperties;
    }

    private static bool AreHeaderBlocksEquivalent(IReadOnlyList<ReportBlockContent> left, IReadOnlyList<ReportBlockContent> right)
    {
        if (left.Count != right.Count)
        {
            return false;
        }

        for (var index = 0; index < left.Count; index++)
        {
            if (!AreHeaderBlocksEquivalent(left[index], right[index]))
            {
                return false;
            }
        }

        return true;
    }

    private static bool AreHeaderBlocksEquivalent(ReportBlockContent left, ReportBlockContent right)
    {
        return left.Type == right.Type
            && left.TemplateText == right.TemplateText
            && left.Row == right.Row
            && left.Alignment == right.Alignment
            && left.FontSize.Equals(right.FontSize)
            && left.IsBold == right.IsBold
            && (left.Type == ReportBlockType.PageNumber || string.Equals(left.Text, right.Text, StringComparison.Ordinal));
    }

    private static int[] CalculateEqualWidthColumns(int totalWidthTwips, int columnCount)
    {
        var widths = new int[columnCount];
        var runningWidth = 0;

        for (var index = 0; index < columnCount; index++)
        {
            widths[index] = index == columnCount - 1
                ? totalWidthTwips - runningWidth
                : totalWidthTwips / columnCount;
            runningWidth += widths[index];
        }

        return widths;
    }

    private static int[] CalculateColumnWidthsTwips(IReadOnlyList<ReportColumnLayout> columns, int totalWidthTwips)
    {
        var widths = new int[columns.Count];
        var totalWeight = columns.Sum(column => column.WidthWeight <= 0 ? 1 : column.WidthWeight);
        var runningWidth = 0;

        for (var index = 0; index < columns.Count; index++)
        {
            widths[index] = index == columns.Count - 1
                ? totalWidthTwips - runningWidth
                : (int)Math.Round(totalWidthTwips * ((columns[index].WidthWeight <= 0 ? 1 : columns[index].WidthWeight) / totalWeight), MidpointRounding.AwayFromZero);
            runningWidth += widths[index];
        }

        return widths;
    }

    private static int GetPrintableWidthTwips(PageSettings settings)
    {
        var (pageWidthMm, _) = PageMeasurementHelper.GetPageDimensionsMillimeters(settings);
        return PageMeasurementHelper.MillimetersToTwips(pageWidthMm - settings.MarginLeftMm - settings.MarginRightMm);
    }

    private static JustificationValues ToJustification(ReportTextAlignment alignment)
    {
        return alignment switch
        {
            ReportTextAlignment.Center => JustificationValues.Center,
            ReportTextAlignment.Right => JustificationValues.Right,
            _ => JustificationValues.Left
        };
    }
}
