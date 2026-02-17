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
        var fontName = FontResolver.NotoSans;

        // 1行目：タイトル中央 + ページ右
        var hdr = section.Headers.Primary.AddTable();
        hdr.Borders.Width = 0;
        hdr.Rows.LeftIndent = 0;

        // A4縦・左右既定余白前提の目安：合計 ≒16cm
        hdr.AddColumn(Unit.FromCentimeter(5.0));  // 左スペーサ
        hdr.AddColumn(Unit.FromCentimeter(8.0));  // 中央：タイトル
        // ページをもう少し右に


        hdr.AddColumn(Unit.FromCentimeter(3.0));  // 右：ページ

        var r1 = hdr.AddRow();
        r1.TopPadding = 0;
        r1.BottomPadding = 0;

        // タイトル（中央）
        var title = r1.Cells[1].AddParagraph("消費税集計表（売上・仕入）");
        title.Format.Alignment = ParagraphAlignment.Center;
        title.Format.Font.Name = fontName;
        title.Format.Font.Size = 12;
        title.Format.Font.Bold = true;

        // ページ番号（右）
        var page = r1.Cells[2].AddParagraph();
        page.Format.Alignment = ParagraphAlignment.Right;
        page.Format.Font.Name = fontName;
        page.AddPageField(); page.AddText(" 頁");

        // 2行目：タイトル直下の“中央だけ”二重線（幅は中央列の 8cm）
        var underline = section.Headers.Primary.AddTable();
        underline.Borders.Width = 0;
        underline.Rows.LeftIndent = 0;
        underline.AddColumn(Unit.FromCentimeter(5.0)); // 左スペーサ
        underline.AddColumn(Unit.FromCentimeter(8.0)); // 中央：二重線
        underline.AddColumn(Unit.FromCentimeter(3.0)); // 右スペーサ

        var r2 = underline.AddRow();
        r2.HeightRule = RowHeightRule.Exactly;
        r2.Height = Unit.FromPoint(0.45);          // 二重線の“間隔” 0.40〜0.50ptで微調整
        r2.TopPadding = 0;
        r2.BottomPadding = 0;

        // 中央セルだけ上下罫線を引く（= 二重線）
        var mid = r2.Cells[1];
        mid.Borders.Width = 0;
        mid.Borders.Top.Width = 0.35;               // 線の太さ（細め：0.30〜0.40pt）
        mid.Borders.Bottom.Width = 0.35;

        // 左右セルには線を出さない
        r2.Cells[0].Borders.Width = 0;
        r2.Cells[2].Borders.Width = 0;
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


        // 課税ブロックの最終行（例：rows[12]）の直後に「輸出」ブロックの先頭行を追加
        // ===== 二重線（課税 → 輸出）を「スペーサ行」で狭い間隔に表現 =====
        // --- 「課税」→「輸出」の二重線を“専用の1行”で描く ---
        // まず、課税の最終行の下線は消す（重なり防止）
        rows[12].Borders.Bottom.Width = 0;
        for (int c = 0; c < table.Columns.Count; c++)
            rows[12].Cells[c].Borders.Bottom.Width = 0;

        // ① 二重線専用の行を追加（ここが“間隔”を決める唯一の行）
        var dbl = table.AddRow();
        dbl.HeightRule = RowHeightRule.Exactly;
        dbl.Height = Unit.FromPoint(1);   // ★ 二重線の間隔（狭くしたいほど小さく：0.30〜0.50）
        dbl.TopPadding = 0;
        dbl.BottomPadding = 0;
        dbl.Borders.Width = 0;
        for (int c = 0; c < table.Columns.Count; c++)
            dbl.Cells[c].Borders.Width = 0;

        // 全列を1セルに結合して、セルの“上・下”罫線だけを引く
        dbl.Cells[0].MergeRight = table.Columns.Count - 1;
        var cell = dbl.Cells[0];
        cell.Borders.Left.Width = 0;              // 縦線は要らない
        cell.Borders.Right.Width = 0;
        cell.Borders.Top.Width = 0.3;            // ★ 線の太さ（上）
        cell.Borders.Bottom.Width = 0.3;            // ★ 線の太さ（下）

        // ② 二重線のすぐ下に「輸出」先頭行を追加（上罫線は描かない）
        var exportHead = table.AddRow();
        exportHead.TopPadding = 0;               // 二重線に寄せる
        exportHead.BottomPadding = Unit.FromPoint(0.6);
        for (int c = 0; c < table.Columns.Count; c++)
            exportHead.Cells[c].Borders.Top.Width = 0;  // ここに線は描かない（上の dbl 行が担当）
        // 二重線のセット終了

        // 「輸出」行のセット
        exportHead.Cells[0].MergeRight = 1;
        exportHead.Cells[0].VerticalAlignment = VerticalAlignment.Center;
        SetCell(exportHead.Cells[0], "輸出", fontName);
        exportHead.Cells[0].MergeDown = 1;
        SetCell(exportHead.Cells[2], "売上", fontName);

        var exportSecond = table.AddRow();
        SetCell(exportSecond.Cells[2], "返還", fontName);


    }

    static void SetCell(MigraDoc.DocumentObjectModel.Tables.Cell cell, string text, string fontName)
    {
        var p = cell.AddParagraph(text);
        p.Format.Font.Name = fontName;
    }

}
