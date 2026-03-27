using System.Text.Json;
using System.Text.Json.Serialization;
using ReportGenerator.Core.Abstractions;
using ReportGenerator.Core.Models;

namespace ReportGenerator.Core.Services;

public sealed class TemplateStorageService : ITemplateStorageService
{
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true,
        Converters =
        {
            new JsonStringEnumConverter(JsonNamingPolicy.CamelCase)
        }
    };

    public ReportTemplate CreateDefaultTemplate()
    {
        return NormalizeTemplate(new ReportTemplate
        {
            Name = "Default Sales Report",
            HeaderBlocks =
            {
                new ReportBlock
                {
                    Type = ReportBlockType.StaticText,
                    Text = "Item Report",
                    FontSize = 18,
                    IsBold = true,
                    Alignment = ReportTextAlignment.Left
                },
                new ReportBlock
                {
                    Type = ReportBlockType.StaticText,
                    Text = "Generated from the selected Excel worksheet",
                    FontSize = 10,
                    Alignment = ReportTextAlignment.Left
                }
            },
            DetailTable = new DetailTableDefinition
            {
                Columns =
                {
                    new DetailColumnDefinition { HeaderText = "Item Id", Source = "ItemId", WidthWeight = 1.4 },
                    new DetailColumnDefinition { HeaderText = "Quantity", Source = "Quantity", WidthWeight = 1, Alignment = ReportTextAlignment.Right },
                    new DetailColumnDefinition { HeaderText = "Quality", Source = "Quality", WidthWeight = 1.1 },
                    new DetailColumnDefinition { HeaderText = "Price", Source = "Price", WidthWeight = 1, Alignment = ReportTextAlignment.Right }
                }
            },
            FooterBlocks =
            {
                new ReportBlock
                {
                    Type = ReportBlockType.PageNumber,
                    Text = "Page {page} of {totalPages}",
                    FontSize = 10,
                    Alignment = ReportTextAlignment.Right
                }
            }
        });
    }

    public async Task SaveAsync(ReportTemplate template, string filePath, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);

        var normalizedTemplate = NormalizeTemplate(template);
        var directory = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        await using var stream = File.Create(filePath);
        await JsonSerializer.SerializeAsync(stream, normalizedTemplate, SerializerOptions, cancellationToken);
    }

    public async Task<ReportTemplate> LoadAsync(string filePath, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);

        await using var stream = File.OpenRead(filePath);
        var template = await JsonSerializer.DeserializeAsync<ReportTemplate>(stream, SerializerOptions, cancellationToken);
        return NormalizeTemplate(template ?? CreateDefaultTemplate());
    }

    public static ReportTemplate NormalizeTemplate(ReportTemplate template)
    {
        template.Name = string.IsNullOrWhiteSpace(template.Name) ? "New Report Template" : template.Name.Trim();
        template.PageSettings ??= new PageSettings();
        template.ImportSettings ??= new ImportSettings();
        template.DetailTable ??= new DetailTableDefinition();
        template.HeaderBlocks ??= [];
        template.FooterBlocks ??= [];
        template.DetailTable.Columns ??= [];

        template.PageSettings.MarginTopMm = Math.Max(0, template.PageSettings.MarginTopMm);
        template.PageSettings.MarginRightMm = Math.Max(0, template.PageSettings.MarginRightMm);
        template.PageSettings.MarginBottomMm = Math.Max(0, template.PageSettings.MarginBottomMm);
        template.PageSettings.MarginLeftMm = Math.Max(0, template.PageSettings.MarginLeftMm);
        template.PageSettings.RowsPerPage = Math.Max(1, template.PageSettings.RowsPerPage);
        template.DetailTable.HeaderFontSize = Math.Clamp(template.DetailTable.HeaderFontSize <= 0 ? 10 : template.DetailTable.HeaderFontSize, 8, 24);
        template.DetailTable.ContentFontSize = Math.Clamp(template.DetailTable.ContentFontSize <= 0 ? 10 : template.DetailTable.ContentFontSize, 8, 24);
        template.DetailTable.GroupEveryRows = Math.Max(0, template.DetailTable.GroupEveryRows);

        foreach (var block in template.HeaderBlocks.Concat(template.FooterBlocks))
        {
            block.Text ??= string.Empty;
            block.Source ??= string.Empty;
            block.FontSize = Math.Clamp(block.FontSize, 8, 28);
        }

        foreach (var column in template.DetailTable.Columns)
        {
            column.HeaderText = string.IsNullOrWhiteSpace(column.HeaderText) ? "Column" : column.HeaderText.Trim();
            column.Source ??= string.Empty;
            column.WidthWeight = column.WidthWeight <= 0 ? 1 : column.WidthWeight;
        }

        return template;
    }
}
