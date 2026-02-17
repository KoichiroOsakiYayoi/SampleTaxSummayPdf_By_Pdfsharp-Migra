using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
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
        // ===== 列定義（左エリアを 3 列に分割）=====
        table.AddColumn(Unit.FromCentimeter(0.5)); // [区分(縦)] 「課 税」など縦積み表示
        table.AddColumn(Unit.FromCentimeter(3.0)); // [項目]     売上/対価の返還/小計...
        table.AddColumn(Unit.FromCentimeter(1.5)); // [小項目]   本体/消費税/計...
        table.AddColumn(Unit.FromCentimeter(2.5)); // 外 税
        table.AddColumn(Unit.FromCentimeter(2.5)); // 内 税
        table.AddColumn(Unit.FromCentimeter(2.5)); // 別 記
        table.AddColumn(Unit.FromCentimeter(2.5)); // 税 込
        table.AddColumn(Unit.FromCentimeter(2.5)); // 合 計

        // ヘッダー行
        var headerRow = table.AddRow();
        headerRow.HeadingFormat = true;
        SetCell(headerRow.Cells[0], "", fontName);
        headerRow.Cells[0].MergeRight = 2;     // 左3列はまとめて空欄（見出しは出さない）
        SetCell(headerRow.Cells[3], "外 税", fontName);
        SetCell(headerRow.Cells[4], "内 税", fontName);
        SetCell(headerRow.Cells[5], "別 記", fontName);
        SetCell(headerRow.Cells[6], "税 込", fontName);
        SetCell(headerRow.Cells[7], "合 計", fontName);

        // 課税のセット
        // ===== 明細（課税ブロック 13行：差引計含む）=====
        const int TaxableBlockRows = 13;
        var rows = new Row[TaxableBlockRows];
        for (int i = 0; i < TaxableBlockRows; i++)
        {
            rows[i] = table.AddRow();
            rows[i].TopPadding = 0.6;
            rows[i].BottomPadding = 0.6;
            for (int c = 0; c < table.Columns.Count; c++)
                rows[i].Cells[c].VerticalAlignment = VerticalAlignment.Center;
        }
        // --- 左端の縦書き「課税」（= 13行分）---
        rows[0].Cells[0].MergeDown = TaxableBlockRows - 1;
        var pv = rows[0].Cells[0].AddParagraph();
        pv.Format.Alignment = ParagraphAlignment.Center;
        pv.Format.LineSpacingRule = LineSpacingRule.Exactly;
        pv.Format.LineSpacing = Unit.FromPoint(9); // フォントに合わせ調整
        pv.AddText("課"); pv.AddLineBreak(); pv.AddText("税");
        pv.Format.Font.Name = fontName;
        // 課税のセット 終わり

        // 売上行のセット
        // --- 「売上」（= 2行分：rows[0]～rows[1]）---
        rows[0].Cells[1].MergeDown = 1;         // 2行分に結合
        SetCell(rows[0].Cells[1], "売上", fontName);
        // --- 「本体」（= 2行のうち1行目）---
        SetCell(rows[0].Cells[2], "本体", fontName);
        // === 3行目（売上 / 消費税）===
        SetCell(rows[1].Cells[2], "消費税", fontName);
        // === 4行目（売上 / 小計）===
        rows[2].Cells[1].MergeRight = 1;
        SetCell(rows[2].Cells[1], "小　計", fontName);
        // 売上行のセット 終わり

        // 対価の返還ブロック（売上と同じパターン: 本体, 消費税, 小計）
        rows[3].Cells[1].MergeDown = 1;
        SetCell(rows[3].Cells[1], "対価の返還", fontName);
        SetCell(rows[3].Cells[2], "本体", fontName);
        SetCell(rows[4].Cells[2], "消費税", fontName);
        rows[5].Cells[1].MergeRight = 1;
        SetCell(rows[5].Cells[1], "小　計", fontName);
        // 対価の返還 終わり

        // 課税のセット 残り
        // 差引計（貸倒課税分の上）
        rows[6].Cells[1].MergeRight = 1;
        SetCell(rows[6].Cells[1], "差引計", fontName);

        // 貸倒課税分ブロック（本体, 消費税, 小計）
        rows[7].Cells[1].MergeDown = 1;
        SetCell(rows[7].Cells[1], "貸倒課税分", fontName);
        SetCell(rows[7].Cells[2], "本体", fontName);

        SetCell(rows[8].Cells[2], "消費税", fontName);

        rows[9].Cells[1].MergeRight = 1;
        SetCell(rows[9].Cells[1], "計", fontName);

        // 貸倒回収ブロック（本体, 消費税, 小計）
        rows[10].Cells[1].MergeDown = 1;
        SetCell(rows[10].Cells[1], "貸倒回収", fontName);
        SetCell(rows[10].Cells[2], "本体", fontName);

        SetCell(rows[11].Cells[2], "消費税", fontName);

        rows[12].Cells[1].MergeRight = 1;
        SetCell(rows[12].Cells[1], "計", fontName);
        // 課税のセット 終わり

    }

    static void SetCell(MigraDoc.DocumentObjectModel.Tables.Cell cell, string text, string fontName)
    {
        var p = cell.AddParagraph(text);
        p.Format.Font.Name = fontName;
    }

}
