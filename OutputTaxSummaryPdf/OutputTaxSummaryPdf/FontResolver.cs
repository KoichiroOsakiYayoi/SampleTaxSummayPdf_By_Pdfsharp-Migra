using PdfSharp.Fonts;
using System.Reflection;

namespace OutputTaxSummaryPdf;

/// <summary>
/// NotoSansJP（Fonts フォルダの .ttf）を解決するフォントリゾルバー。
/// 参照: nextgen_kaikei_journal の FontResolver と同様の構成。
/// </summary>
sealed class FontResolver : IFontResolver
{
    public static readonly FontResolver Instance = new();

    /// <summary>MigraDoc / XFont で指定するフォント名（ResolveTypeface の familyName に一致させる）</summary>
    public const string NotoSans = "Yu Mincho";

    const string NotoSansLight = "NotoSansJP#Light";
    const string NotoSansRegular = "NotoSansJP#Regular";

    const string FileNameLight = "NotoSansJP-Light.ttf";
    const string FileNameRegular = "NotoSansJP-Regular.ttf";

    /// <summary>デフォルトで使用するフォント名（スタイルの Font.Name に設定する値）</summary>
    public string DefaultFontName => NotoSans;

    static string? _fontsBasePath;

    /// <summary>Fonts フォルダのパス（複数候補を試す）</summary>
    static string? FontsBasePath
    {
        get
        {
            if (_fontsBasePath != null)
                return _fontsBasePath;

            var candidates = new List<string>
            {
                Path.Combine(AppContext.BaseDirectory, "Fonts"),
                Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? "", "Fonts"),
                Path.Combine(Directory.GetCurrentDirectory(), "Fonts"),
                Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Fonts")
            };

            foreach (var basePath in candidates)
            {
                var resolved = Path.GetFullPath(basePath);
                var lightPath = Path.Combine(resolved, FileNameLight);
                if (File.Exists(lightPath))
                {
                    _fontsBasePath = resolved;
                    return _fontsBasePath;
                }
            }

            return null;
        }
    }

    FontResolver() { }

    public FontResolverInfo? ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        if (string.Equals(familyName, NotoSans, StringComparison.OrdinalIgnoreCase))
        {
            var faceName = isBold ? NotoSansRegular : NotoSansLight;
            return new FontResolverInfo(faceName, false, isItalic);
        }

        return PlatformFontResolver.ResolveTypeface(familyName, isBold, isItalic);
    }

    public byte[]? GetFont(string faceName)
    {
        var basePath = FontsBasePath;
        if (basePath == null)
            throw new InvalidOperationException(
                "NotoSansJP フォントが見つかりません。Fonts フォルダに NotoSansJP-Light.ttf と NotoSansJP-Regular.ttf を配置し、" +
                "プロジェクトの Fonts\\*.ttf が出力にコピーされるようにしてください。探索したパス: " +
                string.Join("; ", new[]
                {
                    Path.Combine(AppContext.BaseDirectory, "Fonts"),
                    Path.Combine(Directory.GetCurrentDirectory(), "Fonts")
                }));

        if (faceName == NotoSansLight)
            return File.ReadAllBytes(Path.Combine(basePath, FileNameLight));
        if (faceName == NotoSansRegular)
            return File.ReadAllBytes(Path.Combine(basePath, FileNameRegular));

        return null;
    }
}
