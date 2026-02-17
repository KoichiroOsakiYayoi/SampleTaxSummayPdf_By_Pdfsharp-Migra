using MigraDoc.DocumentObjectModel;
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
    /// 中央に「消 費 税 集 計 表 ( 売 上 ・ 仕 入 )」、右に「1 頁」（ページ番号）を1行で配置する。
    /// 各段落に明示的に NotoSansJP を指定して文字化けを防ぐ。
    /// </summary>
    static void CreateHeader(Section section, Document document)
    {
        var titleP = section.Headers.Primary.AddParagraph();
        titleP.Style = "Header";
        titleP.Format.Alignment = ParagraphAlignment.Center;
        titleP.AddText("消費税集計表（売上・仕入）");
        titleP.Format.Font.Name = FontResolver.NotoSans;

        // ページ番号（右寄せ）— タイトルの直下右肩に来る
        var pageP = section.Headers.Primary.AddParagraph();
        pageP.Format.Alignment = ParagraphAlignment.Right;
        pageP.Format.SpaceBefore = Unit.FromPoint(-titleP.GetType().Name.Length * 2);
        //pageP.Format.SpaceBefore = Unit.FromPoint(-10); // ↑重なりすぎる場合はここを微調整（負の余白で近づける）
        pageP.AddPageField();
        pageP.AddText(" 頁");
        pageP.Format.Font.Name = FontResolver.NotoSans;
    }

    static void CreateDetail(Section section, Document document)
    {

    }

}
