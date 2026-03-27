using System.Collections.ObjectModel;

namespace ReportGenerator.Core.Models;

public sealed class ReportTemplate
{
    public string Name { get; set; } = "New Report Template";

    public PageSettings PageSettings { get; set; } = new();

    public ImportSettings ImportSettings { get; set; } = new();

    public Collection<ReportBlock> HeaderBlocks { get; set; } = [];

    public DetailTableDefinition DetailTable { get; set; } = new();

    public Collection<ReportBlock> FooterBlocks { get; set; } = [];
}
