using System.Collections.ObjectModel;

namespace ReportGenerator.Core.Models;

public sealed class DetailTableDefinition
{
    public Collection<DetailColumnDefinition> Columns { get; set; } = [];

    public bool RepeatHeaderOnEveryPage { get; set; } = true;

    public double HeaderFontSize { get; set; } = 10;

    public double ContentFontSize { get; set; } = 10;

    public double ContentRowSpacing { get; set; }

    public int GroupEveryRows { get; set; }

    public double GroupSpacingRows { get; set; } = 1;
}
