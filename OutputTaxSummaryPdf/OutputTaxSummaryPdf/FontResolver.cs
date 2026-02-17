using PdfSharp.Fonts;

namespace OutputTaxSummaryPdf;

/// <summary>
/// NotoSansJP（Fonts フォルダの .ttf）を解決するフォントリゾルバー。
/// 参照: nextgen_kaikei_journal の FontResolver と同様の構成。
/// </summary>
sealed class FontResolver : IFontResolver
{
    public static readonly FontResolver Instance = new();

    /// <summary>MigraDoc / XFont で指定するフォント名（ResolveTypeface の familyName に一致させる）</summary>
    public const string NotoSans = "NotoSansJP";

    /// <summary>内部用 Face 名（Light）</summary>
    const string NotoSansLight = "NotoSansJP#Light";

    /// <summary>内部用 Face 名（Regular）</summary>
    const string NotoSansRegular = "NotoSansJP#Regular";

    static readonly string FontsBasePath = Path.Combine(AppContext.BaseDirectory, "Fonts");
    static readonly string PathLight = Path.Combine(FontsBasePath, "NotoSansJP-Light.ttf");
    static readonly string PathRegular = Path.Combine(FontsBasePath, "NotoSansJP-Regular.ttf");

    /// <summary>デフォルトで使用するフォント名（スタイルの Font.Name に設定する値）</summary>
    public string DefaultFontName => NotoSans;

    FontResolver() { }

    public FontResolverInfo? ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        if (string.Equals(familyName, NotoSans, StringComparison.OrdinalIgnoreCase))
        {
            // Bold の場合は Regular、それ以外は Light
            var faceName = isBold ? NotoSansRegular : NotoSansLight;
            return new FontResolverInfo(faceName, false, isItalic);
        }

        return PlatformFontResolver.ResolveTypeface(familyName, isBold, isItalic);
    }

    public byte[]? GetFont(string faceName)
    {
        if (faceName == NotoSansLight && File.Exists(PathLight))
            return File.ReadAllBytes(PathLight);
        if (faceName == NotoSansRegular && File.Exists(PathRegular))
            return File.ReadAllBytes(PathRegular);
        return null;
    }
}
