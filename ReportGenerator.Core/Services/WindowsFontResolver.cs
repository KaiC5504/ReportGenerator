using System.IO;
using PdfSharp.Fonts;

namespace ReportGenerator.Core.Services;

internal sealed class WindowsFontResolver : IFontResolver
{
    private const string RegularFace = "SegoeUI-Regular";
    private const string BoldFace = "SegoeUI-Bold";
    private const string ItalicFace = "SegoeUI-Italic";
    private const string BoldItalicFace = "SegoeUI-BoldItalic";

    public byte[]? GetFont(string faceName)
    {
        return faceName switch
        {
            RegularFace => ReadFontFile("segoeui.ttf"),
            BoldFace => ReadFontFile("segoeuib.ttf"),
            ItalicFace => ReadFontFile("segoeuii.ttf"),
            BoldItalicFace => ReadFontFile("segoeuiz.ttf"),
            _ => null
        };
    }

    public FontResolverInfo? ResolveTypeface(string familyName, bool bold, bool italic)
    {
        var normalizedFamily = familyName.Replace(" ", string.Empty, StringComparison.OrdinalIgnoreCase);
        if (!normalizedFamily.Equals("SegoeUI", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        if (bold && italic)
        {
            return new FontResolverInfo(BoldItalicFace);
        }

        if (bold)
        {
            return new FontResolverInfo(BoldFace);
        }

        if (italic)
        {
            return new FontResolverInfo(ItalicFace);
        }

        return new FontResolverInfo(RegularFace);
    }

    private static byte[] ReadFontFile(string fileName)
    {
        var windowsDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        var fontPath = Path.Combine(windowsDirectory, "Fonts", fileName);
        return File.ReadAllBytes(fontPath);
    }
}
