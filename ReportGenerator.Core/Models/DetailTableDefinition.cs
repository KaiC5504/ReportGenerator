using System.Collections.ObjectModel;

namespace ReportGenerator.Core.Models;

public sealed class DetailTableDefinition
{
    public Collection<DetailColumnDefinition> Columns { get; set; } = [];

    public bool RepeatHeaderOnEveryPage { get; set; } = true;

    public double HeaderFontSize { get; set; } = 10;

    public double ContentFontSize { get; set; } = 10;

    public int GroupEveryRows { get; set; }
}
