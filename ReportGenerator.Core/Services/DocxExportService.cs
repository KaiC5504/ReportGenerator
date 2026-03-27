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

        var headerPart = mainPart.AddNewPart<HeaderPart>();
        var footerPart = mainPart.AddNewPart<FooterPart>();

        headerPart.Header = CreateHeader(report.Pages.FirstOrDefault()?.HeaderBlocks ?? []);
        footerPart.Footer = CreateFooter(report.Pages.FirstOrDefault()?.FooterBlocks ?? []);

        foreach (var page in report.Pages)
        {
            cancellationToken.ThrowIfCancellationRequested();
            body.Append(CreatePageTable(report, page));

            if (page.PageNumber < page.TotalPages)
            {
                body.Append(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
            }
        }

        body.Append(CreateSectionProperties(
            report.PageSettings,
            mainPart.GetIdOfPart(headerPart),
            mainPart.GetIdOfPart(footerPart)));

        mainPart.Document.Save();
    }

    private static Header CreateHeader(IReadOnlyList<ReportBlockContent> blocks)
    {
        var header = new Header();
        foreach (var block in blocks)
        {
            header.Append(CreateTextParagraph(block));
        }

        return header;
    }

    private static Footer CreateFooter(IReadOnlyList<ReportBlockContent> blocks)
    {
        var footer = new Footer();
        foreach (var block in blocks)
        {
            footer.Append(CreateFooterParagraph(block));
        }

        return footer;
    }

    private static Table CreatePageTable(PagedReport report, ReportPage page)
    {
        var table = new Table();
        var totalWidthTwips = GetPrintableWidthTwips(report.PageSettings);
        var totalWeight = report.Columns.Sum(column => column.WidthWeight <= 0 ? 1 : column.WidthWeight);

        var properties = new TableProperties(
            new TableWidth { Width = totalWidthTwips.ToString(), Type = TableWidthUnitValues.Dxa },
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 8 },
                new BottomBorder { Val = BorderValues.Single, Size = 8 },
                new LeftBorder { Val = BorderValues.Single, Size = 8 },
                new RightBorder { Val = BorderValues.Single, Size = 8 },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 6 },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 6 }));

        table.AppendChild(properties);

        var headerRow = new TableRow();
        for (var index = 0; index < report.Columns.Count; index++)
        {
            var column = report.Columns[index];
            headerRow.Append(CreateCell(
                column.HeaderText,
                column.Alignment,
                true,
                (int)Math.Round(totalWidthTwips * ((column.WidthWeight <= 0 ? 1 : column.WidthWeight) / totalWeight), MidpointRounding.AwayFromZero),
                report.DetailHeaderFontSize));
        }

        table.Append(headerRow);

        foreach (var row in page.Rows)
        {
            if (row.IsSpacer)
            {
                table.Append(CreateSpacerRow(report, totalWidthTwips, totalWeight));
                continue;
            }

            var tableRow = new TableRow();
            for (var index = 0; index < row.Cells.Count; index++)
            {
                var cell = row.Cells[index];
                var column = report.Columns[index];
                tableRow.Append(CreateCell(
                    cell.Text,
                    cell.Alignment,
                    false,
                    (int)Math.Round(totalWidthTwips * ((column.WidthWeight <= 0 ? 1 : column.WidthWeight) / totalWeight), MidpointRounding.AwayFromZero),
                    report.DetailContentFontSize));
            }

            table.Append(tableRow);
        }

        return table;
    }

    private static TableCell CreateCell(string text, ReportTextAlignment alignment, bool isBold, int widthTwips, double fontSize = 10)
    {
        var cell = new TableCell();
        var properties = new TableCellProperties(
            new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = widthTwips.ToString() },
            new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center });

        if (isBold)
        {
            properties.Append(new Shading
            {
                Val = ShadingPatternValues.Clear,
                Fill = "F5F7FA"
            });
        }

        var runProperties = new RunProperties();
        if (isBold)
        {
            runProperties.Append(new Bold());
        }

        runProperties.Append(new RunFonts { Ascii = "Segoe UI", HighAnsi = "Segoe UI" });
        runProperties.Append(new FontSize { Val = ((int)Math.Round(fontSize * 2, MidpointRounding.AwayFromZero)).ToString() });

        var paragraphProperties = new ParagraphProperties(
            new Justification { Val = ToJustification(alignment) },
            new SpacingBetweenLines { Before = "0", After = "0" });

        var paragraph = new Paragraph(paragraphProperties, new Run(runProperties, new Text(text ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve }));
        cell.Append(properties, paragraph);
        return cell;
    }

    private static TableRow CreateSpacerRow(PagedReport report, int totalWidthTwips, double totalWeight)
    {
        var tableRow = new TableRow(new TableRowProperties(
            new TableRowHeight
            {
                HeightType = HeightRuleValues.AtLeast,
                Val = (UInt32Value)(uint)Math.Max(220, Math.Round(report.DetailContentFontSize * 22, MidpointRounding.AwayFromZero))
            }));

        foreach (var column in report.Columns)
        {
            var widthTwips = (int)Math.Round(totalWidthTwips * ((column.WidthWeight <= 0 ? 1 : column.WidthWeight) / totalWeight), MidpointRounding.AwayFromZero);
            tableRow.Append(CreateSpacerCell(widthTwips));
        }

        return tableRow;
    }

    private static TableCell CreateSpacerCell(int widthTwips)
    {
        var cell = new TableCell();
        var properties = new TableCellProperties(
            new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = widthTwips.ToString() },
            new TableCellBorders(
                new TopBorder { Val = BorderValues.Nil },
                new BottomBorder { Val = BorderValues.Nil },
                new LeftBorder { Val = BorderValues.Nil },
                new RightBorder { Val = BorderValues.Nil }));

        cell.Append(properties, new Paragraph(new ParagraphProperties(
            new SpacingBetweenLines { Before = "0", After = "0" })));
        return cell;
    }

    private static Paragraph CreateTextParagraph(ReportBlockContent block)
    {
        var paragraphProperties = new ParagraphProperties(
            new Justification { Val = ToJustification(block.Alignment) },
            new SpacingBetweenLines { Before = "0", After = "0" });

        var runProperties = CreateRunProperties(block.FontSize, block.IsBold);
        return new Paragraph(paragraphProperties, new Run(runProperties, new Text(block.Text) { Space = SpaceProcessingModeValues.Preserve }));
    }

    private static Paragraph CreateFooterParagraph(ReportBlockContent block)
    {
        if (!block.Text.Contains("{page}", StringComparison.Ordinal) && !block.Text.Contains("{totalPages}", StringComparison.Ordinal))
        {
            return CreateTextParagraph(block);
        }

        var paragraph = new Paragraph(new ParagraphProperties(
            new Justification { Val = ToJustification(block.Alignment) },
            new SpacingBetweenLines { Before = "0", After = "0" }));

        var runProperties = CreateRunProperties(block.FontSize, block.IsBold);
        AppendTextAndFields(paragraph, runProperties, block.Text);
        return paragraph;
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
                paragraph.Append(new SimpleField { Instruction = " PAGE " });
                remaining = remaining[(nextIndex + "{page}".Length)..];
            }
            else
            {
                paragraph.Append(new SimpleField { Instruction = " NUMPAGES " });
                remaining = remaining[(nextIndex + "{totalPages}".Length)..];
            }
        }
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

    private static SectionProperties CreateSectionProperties(PageSettings settings, string headerPartId, string footerPartId)
    {
        var (pageWidthMm, pageHeightMm) = PageMeasurementHelper.GetPageDimensionsMillimeters(settings);
        return new SectionProperties(
            new HeaderReference { Type = HeaderFooterValues.Default, Id = headerPartId },
            new FooterReference { Type = HeaderFooterValues.Default, Id = footerPartId },
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
