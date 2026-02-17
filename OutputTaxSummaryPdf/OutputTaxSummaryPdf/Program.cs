using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.Rendering;
using PdfSharp.Fonts;

namespace OutputTaxSummaryPdf;

class Program
{
    static void Main(string[] args)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        GlobalFontSettings.FontResolver = FontResolver.Instance;

        var document = new Document();
        var section = document.AddSection();

        // 用紙 A4
        section.PageSetup.PageFormat = PageFormat.A4;

        CreateHeader(section, document);

        // ヘッダー確認用のダミー本文（空でも可）
        section.AddParagraph();

        var renderer = new PdfDocumentRenderer
        {
            Document = document
        };
        renderer.RenderDocument();

        var outputDir = Path.Combine(AppContext.BaseDirectory, "..", "..", "..");
        var outputPath = Path.GetFullPath(Path.Combine(outputDir, "消費税集計表.pdf"));
        renderer.Save(outputPath);
        Console.WriteLine($"PDF を出力しました: {outputPath}");
    }

    /// <summary>
    /// 消費税集計表のヘッダー1行目を設定する。
    /// 中央に「消 費 税 集 計 表 ( 売 上 ・ 仕 入 )」、右に「1 頁」（ページ番号）を配置する。
    /// </summary>
    static void CreateHeader(Section section, Document document)
    {
        var header = section.Headers.Primary;

        var headerStyle = document.Styles[StyleNames.Header]!;
        headerStyle.Font.Name = FontResolver.NotoSans;
        headerStyle.ParagraphFormat.ClearAll();
        headerStyle.ParagraphFormat.TabStops.AddTabStop(Unit.FromCentimeter(10.5), TabAlignment.Center);
        headerStyle.ParagraphFormat.TabStops.AddTabStop(Unit.FromCentimeter(21), TabAlignment.Right);

        var paragraph = header.AddParagraph();
        paragraph.AddTab();
        paragraph.AddText("消 費 税 集 計 表 ( 売 上 ・ 仕 入 )");
        paragraph.AddTab();
        paragraph.AddPageField();
        paragraph.AddText(" 頁");
    }
}
