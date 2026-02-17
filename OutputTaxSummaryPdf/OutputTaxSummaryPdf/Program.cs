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

        CreateDetail(section, document);

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

    /// <summary>
    /// 消費税集計表の本体テーブル。画像のとりあえず1行目（ヘッダー＋「本体」の1データ行）のみ実装。
    /// </summary>
    static void CreateDetail(Section section, Document document)
    {
        var fontName = FontResolver.NotoSans;

        var table = section.AddTable();
        table.Borders.Visible = true;
        // 列: 課税(項目) | 外税 | 内税 | 別記 | 税込 | 合計
        table.AddColumn(Unit.FromCentimeter(2.5));  // 項目
        table.AddColumn(Unit.FromCentimeter(3.0));  // 外 税
        table.AddColumn(Unit.FromCentimeter(3.0));  // 内 税
        table.AddColumn(Unit.FromCentimeter(3.0));  // 別 記
        table.AddColumn(Unit.FromCentimeter(2.5));  // 税 込
        table.AddColumn(Unit.FromCentimeter(3.0));  // 合 計

        // ヘッダー行
        var headerRow = table.AddRow();
        headerRow.HeadingFormat = true;
        SetCell(headerRow.Cells[0], "", fontName);
        SetCell(headerRow.Cells[1], "外 税", fontName);
        SetCell(headerRow.Cells[2], "内 税", fontName);
        SetCell(headerRow.Cells[3], "別 記", fontName);
        SetCell(headerRow.Cells[4], "税 込", fontName);
        SetCell(headerRow.Cells[5], "合 計", fontName);

        //// 1行目: 売上・本体
        //var row1 = table.AddRow();
        //SetCell(row1.Cells[0], "本 体", fontName);
        //SetCell(row1.Cells[1], "0", fontName);
        //SetCell(row1.Cells[2], "90,781,908", fontName);
        //SetCell(row1.Cells[3], "113,822,395", fontName);
        //SetCell(row1.Cells[4], "", fontName);           // 税込は空
        //SetCell(row1.Cells[5], "204,604,303", fontName);
    }

    static void SetCell(MigraDoc.DocumentObjectModel.Tables.Cell cell, string text, string fontName)
    {
        var p = cell.AddParagraph(text);
        p.Format.Font.Name = fontName;
    }

}
