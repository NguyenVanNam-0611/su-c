using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.Json;
using System.Text.Json.Serialization;

using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace Diff_tool.Services
{
    public class ChecksheetWordService
    {
        private MainDocumentPart? _mainPart; private uint _imageId = 1; private readonly string _imageDir;

        public ChecksheetWordService(IConfiguration config)
        {
            _imageDir = config["ImageStore:Dir"]
                ?? throw new Exception("Thiếu config ImageStore:Dir trong appsettings.json");
        }

        private const string ColorDeleted = "FF0000";
        private const string ColorInserted = "00B050";
        private const string ColorLabel = "6B7280";
        private const string Font = "Times New Roman";
        private const string FontSize = "16"; // half-points = 11pt

        public byte[] GenerateChecksheet(List<SectionDto> sections)
        {
            var templatePath = Path.Combine(
                AppContext.BaseDirectory, "Templates", "CQ_006.docx");

            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Không tìm thấy template: {templatePath}");

            using var ms = new MemoryStream();
            using (var fs = File.OpenRead(templatePath))
                fs.CopyTo(ms);
            ms.Position = 0;

            using (var wordDoc = WordprocessingDocument.Open(ms, true))
            {
                _mainPart = wordDoc.MainDocumentPart
                    ?? throw new Exception("Template thiếu MainDocumentPart");

                var body = _mainPart.Document?.Body
                    ?? throw new Exception("Template không có body");

                var table = body.Elements<Table>().ElementAtOrDefault(1)
                    ?? throw new Exception("Template CQ_006.docx không có table dữ liệu (index 1)");

                var templateRow = table.Elements<TableRow>().ElementAtOrDefault(1)
                    ?? throw new Exception("Template không có data row mẫu (index 1)");

                var allRows = sections
                    .SelectMany(s => (s.Changes ?? new())
                        .SelectMany(c => FlattenChange(c, s.Heading ?? "")))
                    .ToList();

                if (allRows.Count > 0)
                {
                    FillDataRow(templateRow, allRows[0].Heading, allRows[0].Change, allRows[0].SubRow);

                    for (int i = 1; i < allRows.Count; i++)
                    {
                        var cloned = (TableRow)templateRow.CloneNode(deep: true);
                        FillDataRow(cloned, allRows[i].Heading, allRows[i].Change, allRows[i].SubRow);
                        table.AppendChild(cloned);
                    }
                }

                _mainPart.Document!.Save();
            }

            return ms.ToArray();
        }

        private record FlatRow(
            string Heading,
            ChangeDto Change,
            TableRowChangeDto? SubRow,
            bool FullTable = false
        );
        private static string GetNearestHeading(string? changeHeading, string? sectionHeading)
        {
            var heading = !string.IsNullOrWhiteSpace(changeHeading)
                ? changeHeading!
                : sectionHeading ?? "";

            heading = heading.Trim();

            if (string.IsNullOrWhiteSpace(heading))
                return "";

            var separators = new[]
            {
    " > ",
    " / ",
    " \\ ",
    "\n",
    "\r\n"
};

            foreach (var sep in separators)
            {
                if (heading.Contains(sep))
                {
                    var parts = heading
                        .Split(sep, StringSplitOptions.RemoveEmptyEntries)
                        .Select(x => x.Trim())
                        .Where(x => !string.IsNullOrWhiteSpace(x))
                        .ToList();

                    return parts.LastOrDefault() ?? heading;
                }
            }

            return heading;
        }
        private IEnumerable<FlatRow> FlattenChange(ChangeDto change, string sectionHeading)
        {
            string heading = change.Type == "heading_modified"
                ? GetNearestHeading(
                    change.Right?.Node?.Content != null &&
                    change.Right.Node.Content.TryGetValue("text_display", out var hd)
                        ? hd.GetString()
                        : change.Right?.Node?.Content != null &&
                          change.Right.Node.Content.TryGetValue("text", out var ht)
                            ? ht.GetString()
                            : change.Heading,
                    sectionHeading)
                : GetNearestHeading(change.Heading, sectionHeading);

            if (change.Type is "table_modified" or "table_inserted" or "table_deleted")
            {
                var renderMode = change.TableAnalysis?.RenderMode ?? "";

                if (
                    renderMode == "full_table" ||
                    renderMode == "full" ||
                    change.ChangeKind == "insert" ||
                    change.ChangeKind == "delete"
                )
                {
                    yield return new FlatRow(heading, change, null, FullTable: true);
                    yield break;
                }

                var tableChanges = change.TableAnalysis?.TableChanges ?? new();
                if (tableChanges.Count > 0)
                {
                    foreach (var rc in tableChanges)
                        yield return new FlatRow(heading, change, rc);

                    yield break;
                }
            }

            yield return new FlatRow(heading, change, null);
        }

        private static void EnsureCellBorders(TableCell cell)
        {
            var tcPr = cell.GetFirstChild<TableCellProperties>();
            if (tcPr == null)
            {
                tcPr = new TableCellProperties();
                cell.InsertAt(tcPr, 0);
            }
            tcPr.RemoveAllChildren<TableCellBorders>();
            tcPr.AppendChild(new TableCellBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "000000" },
                new LeftBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "000000" },
                new BottomBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "000000" },
                new RightBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "000000" }
            ));
        }
        //Helper Text Inline
        private static bool HasOnlyInlineImageChanged(ChangeDto change)
        {
            if (change.ChangeKind != "replace")
                return false;

            var leftNode = change.Left?.Node;
            var rightNode = change.Right?.Node;

            if (leftNode == null || rightNode == null)
                return false;

            string leftText =
                TryGetText(leftNode.Content);

            string rightText =
                TryGetText(rightNode.Content);

            if (leftText.Trim() != rightText.Trim())
                return false;

            var leftImgs = (leftNode.Children ?? new())
                .Where(x => x.Type == "image")
                .ToList();

            var rightImgs = (rightNode.Children ?? new())
                .Where(x => x.Type == "image")
                .ToList();

            if (leftImgs.Count != rightImgs.Count)
                return true;

            for (int i = 0; i < leftImgs.Count; i++)
            {
                string lSha = GetSha(leftImgs[i]);
                string rSha = GetSha(rightImgs[i]);

                if (lSha != rSha)
                    return true;
            }

            return false;
        }

        private static string TryGetText(Dictionary<string, JsonElement>? content)
        {
            if (content == null) return "";

            if (content.TryGetValue("text_display", out var td))
                return td.GetString() ?? "";

            if (content.TryGetValue("text", out var tx))
                return tx.GetString() ?? "";

            return "";
        }

        private static string GetSha(NodeDto node)
        {
            if (node.Content == null)
                return "";

            if (node.Content.TryGetValue("sha256", out var sha))
                return sha.GetString() ?? "";

            return "";
        }
        private static List<NodeDto> GetInlineImageNodes(NodeDto? node)
        {
            return node?.Children?
                .Where(x => x.Type == "image")
                .ToList() ?? new List<NodeDto>();
        }

        private static string? GetImageUrl(NodeDto node)
        {
            if (node.Content != null &&
                node.Content.TryGetValue("image_url", out var iu))
            {
                return iu.GetString();
            }

            return node.ImageUrl;
        }
        // ────────────────────────────────────────────────────────────────────────
        // SetCellContentAsNestedTable — render đúng thứ tự content_blocks + ảnh
        // ────────────────────────────────────────────────────────────────────────
        private void SetCellContentAsNestedTable(TableCell outerCell, TableRowChangeDto subRow, string side)
        {
            foreach (var child in outerCell.Elements<Paragraph>().ToList()) child.Remove();
            foreach (var child in outerCell.Elements<Table>().ToList()) child.Remove();

            var cells = side == "left"
                ? subRow.LeftDisplayCells ?? subRow.LeftCells ?? new()
                : subRow.RightDisplayCells ?? subRow.RightCells ?? new();

            // Build map: anchorCol → CellChangeDto
            var ccMap = (subRow.CellChanges ?? new())
                .Where(cc => cc.AnchorCol.HasValue)
                .GroupBy(cc => cc.AnchorCol!.Value)
                .ToDictionary(
                    g => g.Key,
                    g => g
                        .OrderByDescending(x => x.Changes?.Count ?? 0)
                        .First()
                );

            var nestedTable = new Table();
            nestedTable.AppendChild(new TableProperties(
                new TableWidth { Width = "0", Type = TableWidthUnitValues.Auto },
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new LeftBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new BottomBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new RightBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = "000000" }
                )
            ));

            var nestedRow = new TableRow();

            if (!cells.Any())
            {
                nestedRow.AppendChild(MakeNestedCell(new Paragraph(
                    MakeParaProps(),
                    MakeTextRun(side == "left" ? "(không có)" : "")
                )));
            }
            else
            {
                foreach (var cell in cells.Where(c => c != null))
                {
                    var colIdx = cell.AnchorCol ?? 0;
                    ccMap.TryGetValue(colIdx, out var cc);

                    // Lấy cell node của đúng side từ cc, fallback về cell
                    var sideCell = cc != null
                        ? (side == "left" ? cc.LeftCell : cc.RightCell) ?? cell
                        : cell;

                    var contentBlocks = sideCell?.ContentBlocks;
                    TableCell nestedTc;

                    if (contentBlocks != null && contentBlocks.Count > 0 && cc?.Changes?.Count > 0)
                    {
                        // Có diff → render với highlight
                        nestedTc = MakeCellFromBlocksWithDiff(contentBlocks, cc.Changes, side);
                    }
                    else if (contentBlocks != null && contentBlocks.Count > 0)
                    {
                        // Không có diff → render bình thường
                        nestedTc = MakeCellFromBlocks(contentBlocks);
                    }
                    else
                    {
                        // Fallback: text thuần
                        nestedTc = MakeNestedCell(new Paragraph(
                            MakeParaProps(),
                            MakeTextRun(cell.TextDisplay ?? cell.Text ?? "")
                        ));
                    }

                    nestedRow.AppendChild(nestedTc);
                }
            }

            nestedTable.AppendChild(nestedRow);
            outerCell.AppendChild(nestedTable);
            outerCell.AppendChild(new Paragraph(MakeParaProps()));
        }
        // ── Render content blocks của cell vào tc (support nested table) ──
        private void AppendCellContentBlocks(TableCell tc, CellDto cell, bool isDeleted, bool isInserted)
        {
            var blocks = cell.ContentBlocks;

            if (blocks == null || blocks.Count == 0)
            {
                var para = new Paragraph(MakeParaProps());
                para.AppendChild(MakeTextRun(
                    cell.TextDisplay ?? cell.Text ?? "",
                    deleted: isDeleted, inserted: isInserted));
                tc.AppendChild(para);
                return;
            }

            bool hasContent = false;
            foreach (var block in blocks)
            {
                if (block.Type == "paragraph" && block.Payload != null)
                {
                    var para = new Paragraph(MakeParaProps());
                    var txt = block.Payload.TextDisplay ?? block.Payload.Text ?? "";
                    if (!string.IsNullOrEmpty(txt))
                        para.AppendChild(MakeTextRun(txt, deleted: isDeleted, inserted: isInserted));
                    // inline images trong paragraph
                    foreach (var img in block.Payload.Images ?? new())
                    {
                        if (!string.IsNullOrEmpty(img.ImageUrl))
                            AppendImage(para, img.ImageUrl, deleted: isDeleted, inserted: isInserted);
                    }
                    tc.AppendChild(para);
                    hasContent = true;
                }
                else if (block.Type == "table"
                    && block.Payload?.Cells != null
                    && block.Payload.Cells.Count > 0)
                                {
                    var blockCells = block.Payload.Cells;
                    int bRows = block.Payload.TotalRows
                        ?? blockCells.Max(c => (c.AnchorRow ?? 0) + Math.Max(1, c.RowSpan ?? 1));
                    int bCols = block.Payload.TotalCols
                        ?? blockCells.Max(c => (c.AnchorCol ?? 0) + Math.Max(1, c.ColSpan ?? 1));

                    var nestedTbl = BuildNestedTableFromCells(blockCells, bRows, bCols, isDeleted, isInserted);
                    tc.AppendChild(nestedTbl);
                    tc.AppendChild(new Paragraph(MakeParaProps()));
                    hasContent = true;
                }
                else if (block.Type == "image" && block.Payload?.ImageUrl != null)
                {
                    var para = new Paragraph(MakeParaProps());
                    AppendImage(para, block.Payload.ImageUrl, deleted: isDeleted, inserted: isInserted);
                    tc.AppendChild(para);
                    hasContent = true;
                }
            }

            // Word yêu cầu cell phải có ít nhất 1 paragraph
            if (!hasContent || !tc.Elements<Paragraph>().Any())
                tc.AppendChild(new Paragraph(MakeParaProps()));
        }

        // ── Build Word Table từ flat cells list (giống SetCellContentAsFullTable nhưng trả Table) ──
        private Table BuildNestedTableFromCells(
                List<CellDto> cells,
                int totalRows,
                int totalCols,
                bool isDeleted,
                bool isInserted,
                List<TableRowChangeDto>? nestedChanges = null,
                string side = "left")
        {
            var tbl = new Table();
            tbl.AppendChild(new TableProperties(
                new TableWidth { Width = "0", Type = TableWidthUnitValues.Auto },
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new LeftBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new BottomBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new RightBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = "000000" }
                )
            ));

            var tblGrid = new TableGrid();
            for (int i = 0; i < totalCols; i++)
                tblGrid.AppendChild(new GridColumn { Width = "0" });
            tbl.AppendChild(tblGrid);

            var rowGroups = cells
                    .GroupBy(c => c.AnchorRow ?? 0)
                .OrderBy(g => g.Key)
                .ToDictionary(g => g.Key, g => g.OrderBy(c => c.AnchorCol ?? 0).ToList());

            var spanTracker = new Dictionary<int, (int remainingRows, int colSpan)>();

            for (int r = 0; r < totalRows; r++)
            {
                var nestedRow = new TableRow();
                var masterCells = rowGroups.ContainsKey(r) ? rowGroups[r] : new List<CellDto>();
                var nestedCcMap = (nestedChanges ?? new())
                    .Where(x => x.AnchorRow == r)
                    .SelectMany(x => x.CellChanges ?? new())
                    .Where(x => x.AnchorCol.HasValue)
                    .GroupBy(x => x.AnchorCol!.Value)
                    .ToDictionary(
                        g => g.Key,
                        g => g.OrderByDescending(x => x.Changes?.Count ?? 0).First()
                    );
                int masterIdx = 0;

                for (int c = 0; c < totalCols;)
                {
                    if (spanTracker.TryGetValue(c, out var spanInfo) && spanInfo.remainingRows > 0)
                    {
                        var contTc = new TableCell();
                        var contPr = new TableCellProperties(
                            new TableCellWidth { Width = "0", Type = TableWidthUnitValues.Auto });
                        if (spanInfo.colSpan > 1)
                            contPr.AppendChild(new GridSpan { Val = spanInfo.colSpan });
                        contPr.AppendChild(new VerticalMerge());
                        contPr.AppendChild(MakeSingleBorders());
                        contTc.AppendChild(contPr);
                        contTc.AppendChild(new Paragraph(MakeParaProps()));
                        nestedRow.AppendChild(contTc);

                        var nextRem = spanInfo.remainingRows - 1;
                        if (nextRem > 0) spanTracker[c] = (nextRem, spanInfo.colSpan);
                        else spanTracker.Remove(c);
                        c += spanInfo.colSpan;
                    }
                    else if (masterIdx < masterCells.Count && (masterCells[masterIdx].AnchorCol ?? 0) == c)
                    {
                        var cell = masterCells[masterIdx++];
                        int colSpan = Math.Max(1, Math.Min(cell.ColSpan ?? 1, totalCols - c));
                        int rowSpan = Math.Max(1, cell.RowSpan ?? 1);

                        var tc = new TableCell();
                        var tcPr = new TableCellProperties(
                            new TableCellWidth { Width = "0", Type = TableWidthUnitValues.Auto },
                            MakeSingleBorders(),
                            new TableCellMargin(
                                new TopMargin { Width = "40", Type = TableWidthUnitValues.Dxa },
                                new BottomMargin { Width = "40", Type = TableWidthUnitValues.Dxa },
                                new LeftMargin { Width = "80", Type = TableWidthUnitValues.Dxa },
                                new RightMargin { Width = "80", Type = TableWidthUnitValues.Dxa }
                            ));
                        if (colSpan > 1) tcPr.AppendChild(new GridSpan { Val = colSpan });
                        if (rowSpan > 1) tcPr.AppendChild(new VerticalMerge { Val = MergedCellValues.Restart });
                        tc.AppendChild(tcPr);

                        nestedCcMap.TryGetValue(c, out var nestedCc);

                        if (nestedCc?.Changes?.Count > 0 &&
                            cell.ContentBlocks?.Count > 0)
                        {
                            var rendered = MakeCellFromBlocksWithDiff(
                                cell.ContentBlocks,
                                nestedCc.Changes,
                                side
                            );

                            foreach (var child in rendered.ChildElements.ToList())
                                tc.Append(child.CloneNode(true));
                        }
                        else
                        {
                            AppendCellContentBlocks(
                                tc,
                                cell,
                                isDeleted,
                                isInserted
                            );
                        }

                        nestedRow.AppendChild(tc);
                        if (rowSpan > 1) spanTracker[c] = (rowSpan - 1, colSpan);
                        c += colSpan;
                    }
                    else
                    {
                        var emptyTc = new TableCell();
                        emptyTc.AppendChild(new TableCellProperties(
                            new TableCellWidth { Width = "0", Type = TableWidthUnitValues.Auto },
                            MakeSingleBorders()));
                        emptyTc.AppendChild(new Paragraph(MakeParaProps()));
                        nestedRow.AppendChild(emptyTc);
                        c++;
                    }
                }
                tbl.AppendChild(nestedRow);
            }

            return tbl;
        }

        // ── Helper tạo borders cho cell (tránh lặp code) ─────────────
        private static TableCellBorders MakeSingleBorders() =>
            new TableCellBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "000000" },
                new LeftBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "000000" },
                new BottomBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "000000" },
                new RightBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "000000" }
            );
        private void SetCellContentAsFullTable(
    TableCell outerCell, List<CellDto> cells, bool isDeleted, bool isInserted)
        {
            foreach (var child in outerCell.Elements<Paragraph>().ToList()) child.Remove();
            foreach (var child in outerCell.Elements<Table>().ToList()) child.Remove();

            int totalRows = cells.Max(c => (c.AnchorRow ?? 0) + (c.RowSpan ?? 1));
            int totalCols = cells.Max(c => (c.AnchorCol ?? 0) + (c.ColSpan ?? 1));

            var tbl = BuildNestedTableFromCells(cells, totalRows, totalCols, isDeleted, isInserted);
            outerCell.AppendChild(tbl);
            outerCell.AppendChild(new Paragraph(MakeParaProps()));
        }
        // ── Render content_blocks bình thường (không diff) ───────────────────────
        private TableCell MakeCellFromBlocks(List<ContentBlockDto> blocks)
        {
            var tc = new TableCell();
            tc.AppendChild(new TableCellProperties(
                new TableCellWidth { Width = "0", Type = TableWidthUnitValues.Auto },
                MakeSingleBorders(),
                new TableCellMargin(
                    new TopMargin { Width = "40", Type = TableWidthUnitValues.Dxa },
                    new BottomMargin { Width = "40", Type = TableWidthUnitValues.Dxa },
                    new LeftMargin { Width = "80", Type = TableWidthUnitValues.Dxa },
                    new RightMargin { Width = "80", Type = TableWidthUnitValues.Dxa }
                )
            ));

            AppendBlocksToCell(tc, blocks,
                deletedParaUids: new HashSet<string>(),
                insertedParaUids: new HashSet<string>(),
                deletedImgShas: new HashSet<string>(),
                insertedImgShas: new HashSet<string>(),
                modifiedParaMap: new Dictionary<string, ParagraphChangeDto>(),
                allChanges: new List<ParagraphChangeDto>(),
                side: "left");

            if (!tc.Elements<Paragraph>().Any() && !tc.Elements<Table>().Any())
                tc.AppendChild(new Paragraph(MakeParaProps()));

            return tc;
        }

        // ── Render content_blocks với diff highlights ─────────────────────────────
        private TableCell MakeCellFromBlocksWithDiff(
            List<ContentBlockDto> blocks,
            List<ParagraphChangeDto> changes,
            string side)
        {
            // Build lookup maps từ changes
            var deletedParaUids = new HashSet<string>();
            var insertedParaUids = new HashSet<string>();
            var deletedImgShas = new HashSet<string>();
            var insertedImgShas = new HashSet<string>();
            var modifiedParaMap = new Dictionary<string, ParagraphChangeDto>();

            foreach (var ch in changes)
            {
                switch (ch.Type)
                {
                    case "paragraph_deleted":
                        if (ch.Original?.Uid != null) deletedParaUids.Add(ch.Original.Uid);
                        break;
                    case "paragraph_inserted":
                        if (ch.Modified?.Uid != null) insertedParaUids.Add(ch.Modified.Uid);
                        break;
                    case "paragraph_modified":
                        if (ch.Original?.Uid != null) modifiedParaMap[ch.Original.Uid] = ch;
                        if (ch.Modified?.Uid != null) modifiedParaMap[ch.Modified.Uid] = ch;
                        break;
                    case "image_deleted":
                        if (ch.Original?.Sha256 != null) deletedImgShas.Add(ch.Original.Sha256);
                        break;
                    case "image_added":
                        if (ch.Modified?.Sha256 != null) insertedImgShas.Add(ch.Modified.Sha256);
                        break;
                }
            }

            var tc = new TableCell();
            tc.AppendChild(new TableCellProperties(
                new TableCellWidth { Width = "0", Type = TableWidthUnitValues.Auto },
                MakeSingleBorders(),
                new TableCellMargin(
                    new TopMargin { Width = "40", Type = TableWidthUnitValues.Dxa },
                    new BottomMargin { Width = "40", Type = TableWidthUnitValues.Dxa },
                    new LeftMargin { Width = "80", Type = TableWidthUnitValues.Dxa },
                    new RightMargin { Width = "80", Type = TableWidthUnitValues.Dxa }
                )
            ));

            AppendBlocksToCell(tc, blocks,
                deletedParaUids, insertedParaUids,
                deletedImgShas, insertedImgShas,
                modifiedParaMap, changes, side);

            // side == "right": append image_added nếu không có trong content_blocks
            if (side == "right")
            {
                var lastPara = tc.Elements<Paragraph>().LastOrDefault();
                if (lastPara == null)
                {
                    lastPara = new Paragraph(MakeParaProps());
                    tc.AppendChild(lastPara);
                }
                foreach (var ch in changes.Where(c => c.Type == "image_added" && c.Modified?.ImageUrl != null))
                {
                    var sha = ch.Modified!.Sha256 ?? "";
                    bool alreadyIn = blocks.Any(b => b.Type == "image" && b.Payload?.Sha256 == sha);
                    if (!alreadyIn)
                        AppendImage(lastPara, ch.Modified.ImageUrl!, inserted: true);
                }
            }

            if (!tc.Elements<Paragraph>().Any() && !tc.Elements<Table>().Any())
                tc.AppendChild(new Paragraph(MakeParaProps()));

            return tc;
        }
        // ── Core: duyệt blocks và emit runs/images vào paragraph ─────────────────
        private void AppendBlocksToParagraph(
            Paragraph para,
            List<ContentBlockDto> blocks,
            HashSet<string> deletedParaUids,
            HashSet<string> insertedParaUids,
            HashSet<string> deletedImgShas,
            HashSet<string> insertedImgShas,
            Dictionary<string, ParagraphChangeDto> modifiedParaMap,
            List<ParagraphChangeDto> allChanges,
            string side)
        {
            foreach (var block in blocks)
            {
                if (block.Type == "paragraph" && block.Payload != null)
                {
                    var p = block.Payload;
                    var uid = p.Uid ?? "";
                    var txt = p.TextDisplay ?? p.Text ?? "";
                    bool isDeleted = deletedParaUids.Contains(uid);
                    bool isInserted = insertedParaUids.Contains(uid);

                    // Skip deleted on right / inserted on left
                    if (isDeleted && side == "right") continue;
                    if (isInserted && side == "left") continue;

                    if (modifiedParaMap.TryGetValue(uid, out var modCh) && modCh.Spans?.Count > 0)
                    {
                        // Word-level diff
                        bool firstSpan = true;
                        foreach (var span in modCh.Spans)
                        {
                            string prefix = firstSpan ? "" : " ";
                            firstSpan = false;
                            switch (span.Type)
                            {
                                case "equal":
                                    para.AppendChild(MakeTextRun(prefix + (span.Text ?? "")));
                                    break;
                                case "delete" when side == "left":
                                    para.AppendChild(MakeTextRun(prefix + (span.OldText ?? span.Text ?? ""), deleted: true));
                                    break;
                                case "insert" when side == "right":
                                    para.AppendChild(MakeTextRun(prefix + (span.NewText ?? span.Text ?? ""), inserted: true));
                                    break;
                                case "replace":
                                    var rText = side == "left" ? span.OldText ?? "" : span.NewText ?? "";
                                    para.AppendChild(MakeTextRun(prefix + rText,
                                        deleted: side == "left", inserted: side == "right"));
                                    break;
                            }
                        }
                        para.AppendChild(new Run(new Break()));
                    }
                    else if (!string.IsNullOrEmpty(txt))
                    {
                        // Paragraph bình thường hoặc toàn bộ paragraph bị deleted/inserted
                        var lines = txt.Split('\n');
                        for (int i = 0; i < lines.Length; i++)
                        {
                            if (i > 0) para.AppendChild(new Run(new Break()));
                            if (!string.IsNullOrEmpty(lines[i]))
                                para.AppendChild(MakeTextRun(lines[i],
                                    deleted: isDeleted, inserted: isInserted));
                        }
                        para.AppendChild(new Run(new Break()));
                    }

                    // Inline images trong paragraph
                    foreach (var img in p.Images ?? new())
                    {
                        if (!string.IsNullOrEmpty(img.ImageUrl))
                            AppendImage(para, img.ImageUrl,
                                deleted: isDeleted, inserted: isInserted);
                    }
                }
                else if (block.Type == "image" && block.Payload?.ImageUrl != null)
                {
                    var sha = block.Payload.Sha256 ?? "";
                    bool isDeleted = deletedImgShas.Contains(sha);
                    bool isInserted = insertedImgShas.Contains(sha);

                    // Skip deleted on right / inserted on left
                    if (isDeleted && side == "right") continue;
                    if (isInserted && side == "left") continue;

                    AppendImage(para, block.Payload.ImageUrl,
                        deleted: isDeleted, inserted: isInserted);
                }
                // shape: skip
            }
        }
        // ── Append blocks vào TableCell — hỗ trợ cả paragraph lẫn nested table ──
        private void AppendBlocksToCell(
            TableCell tc,
            List<ContentBlockDto> blocks,
            HashSet<string> deletedParaUids,
            HashSet<string> insertedParaUids,
            HashSet<string> deletedImgShas,
            HashSet<string> insertedImgShas,
            Dictionary<string, ParagraphChangeDto> modifiedParaMap,
            List<ParagraphChangeDto> allChanges,
            string side)
        {
            // Gom các paragraph/image blocks liên tiếp vào 1 para
            // Khi gặp block type="table" thì flush para hiện tại, append table, rồi tiếp tục
            var currentPara = new Paragraph(MakeParaProps());
            bool paraHasContent = false;

            void FlushPara()
            {
                if (paraHasContent)
                {
                    tc.AppendChild(currentPara);
                    currentPara = new Paragraph(MakeParaProps());
                    paraHasContent = false;
                }
            }

            foreach (var block in blocks)
            {
                if (block.Type == "table" && block.Payload?.Cells != null && block.Payload.Cells.Count > 0)
                {
                    FlushPara();

                    var blockCells = block.Payload.Cells;
                    int bRows = block.Payload.TotalRows
                        ?? blockCells.Max(c => (c.AnchorRow ?? 0) + Math.Max(1, c.RowSpan ?? 1));
                    int bCols = block.Payload.TotalCols
                        ?? blockCells.Max(c => (c.AnchorCol ?? 0) + Math.Max(1, c.ColSpan ?? 1));

                    bool isDeleted = side == "left";
                    bool isInserted = side == "right";
                    var nestedDiff = allChanges
                        .FirstOrDefault(x => x.Type == "nested_table_modified");

                    var nestedTbl = BuildNestedTableFromCells(
                        blockCells,
                        bRows,
                        bCols,
                        false,
                        false,
                        nestedDiff?.Changes,
                        side
                    );
                    tc.AppendChild(nestedTbl);
                    tc.AppendChild(new Paragraph(MakeParaProps()));
                }
                else if (block.Type == "paragraph" && block.Payload != null)
                {
                    var p = block.Payload;
                    var uid = p.Uid ?? "";
                    var txt = p.TextDisplay ?? p.Text ?? "";
                    bool isDeleted = deletedParaUids.Contains(uid);
                    bool isInserted = insertedParaUids.Contains(uid);

                    if (isDeleted && side == "right") continue;
                    if (isInserted && side == "left") continue;

                    if (modifiedParaMap.TryGetValue(uid, out var modCh) && modCh.Spans?.Count > 0)
                    {
                        bool firstSpan = true;
                        foreach (var span in modCh.Spans)
                        {
                            string prefix = firstSpan ? "" : " ";
                            firstSpan = false;
                            switch (span.Type)
                            {
                                case "equal":
                                    currentPara.AppendChild(MakeTextRun(prefix + (span.Text ?? "")));
                                    break;
                                case "delete" when side == "left":
                                    currentPara.AppendChild(MakeTextRun(prefix + (span.OldText ?? span.Text ?? ""), deleted: true));
                                    break;
                                case "insert" when side == "right":
                                    currentPara.AppendChild(MakeTextRun(prefix + (span.NewText ?? span.Text ?? ""), inserted: true));
                                    break;
                                case "replace":
                                    var rText = side == "left" ? span.OldText ?? "" : span.NewText ?? "";
                                    currentPara.AppendChild(MakeTextRun(prefix + rText,
                                        deleted: side == "left", inserted: side == "right"));
                                    break;
                            }
                        }
                        currentPara.AppendChild(new Run(new Break()));
                        paraHasContent = true;
                    }
                    else if (!string.IsNullOrEmpty(txt))
                    {
                        var lines = txt.Split('\n');
                        for (int i = 0; i < lines.Length; i++)
                        {
                            if (i > 0) currentPara.AppendChild(new Run(new Break()));
                            if (!string.IsNullOrEmpty(lines[i]))
                                currentPara.AppendChild(MakeTextRun(lines[i],
                                    deleted: isDeleted, inserted: isInserted));
                        }
                        currentPara.AppendChild(new Run(new Break()));
                        paraHasContent = true;
                    }

                    foreach (var img in p.Images ?? new())
                    {
                        if (!string.IsNullOrEmpty(img.ImageUrl))
                        {
                            AppendImage(currentPara, img.ImageUrl, deleted: isDeleted, inserted: isInserted);
                            paraHasContent = true;
                        }
                    }
                }
                else if (block.Type == "image" && block.Payload?.ImageUrl != null)
                {
                    var sha = block.Payload.Sha256 ?? "";
                    bool isDeleted = deletedImgShas.Contains(sha);
                    bool isInserted = insertedImgShas.Contains(sha);

                    if (isDeleted && side == "right") continue;
                    if (isInserted && side == "left") continue;

                    AppendImage(currentPara, block.Payload.ImageUrl, deleted: isDeleted, inserted: isInserted);
                    paraHasContent = true;
                }
            }

            FlushPara();
        }
        // ── Append image run vào paragraph ───────────────────────────────────────
        private void AppendImage(Paragraph para, string imageUrl,
            bool deleted = false, bool inserted = false)
        {
            if (_mainPart == null) return;
            var imagePath = ResolveImagePath(imageUrl);
            if (imagePath == null)
            {
                para.AppendChild(MakeTextRun($"[Ảnh không tìm thấy: {imageUrl}]"));
                return;
            }

            if (deleted || inserted)
            {
                // Label text trước ảnh để biết trạng thái
                para.AppendChild(MakeTextRun(
                    deleted ? "[Ảnh đã xóa] " : "[Ảnh mới] ",
                    deleted: deleted, inserted: inserted));
            }

            para.AppendChild(MakeImageRun(_mainPart, imagePath, 280, 180));
            para.AppendChild(new Run(new Break()));
        }

        // ── Tạo TableCell bọc 1 paragraph ────────────────────────────────────────
        private static TableCell MakeNestedCell(Paragraph para)
        {
            var tc = new TableCell();
            tc.AppendChild(new TableCellProperties(
                new TableCellWidth { Width = "0", Type = TableWidthUnitValues.Auto },
                new TableCellBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "000000" },
                    new LeftBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "000000" },
                    new BottomBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "000000" },
                    new RightBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "000000" }
                ),
                new TableCellMargin(
                    new TopMargin { Width = "40", Type = TableWidthUnitValues.Dxa },
                    new BottomMargin { Width = "40", Type = TableWidthUnitValues.Dxa },
                    new LeftMargin { Width = "80", Type = TableWidthUnitValues.Dxa },
                    new RightMargin { Width = "80", Type = TableWidthUnitValues.Dxa }
                )
            ));
            tc.AppendChild(para);
            return tc;
        }

        private void FillDataRow(TableRow row, string heading, ChangeDto change, TableRowChangeDto? subRow)
        {
            var cells = row.Elements<TableCell>().ToList();
            if (cells.Count < 4) return;

            foreach (var cell in cells)
                EnsureCellBorders(cell);

            SetCellContent(cells[0], new[] { new RunInfo { Text = GetPageText(change) } }, center: true);
            SetCellContent(cells[1], new[] { new RunInfo { Text = heading } });

            if (subRow != null)
            {
                var leftCells = subRow.LeftDisplayCells ?? subRow.LeftCells ?? new();
                var rightCells = subRow.RightDisplayCells ?? subRow.RightCells ?? new();

                if (leftCells.Any())
                    SetCellContentAsNestedTable(cells[2], subRow, "left");
                else
                    SetCellContent(cells[2], new[] { new RunInfo { Text = "" } });

                if (rightCells.Any())
                    SetCellContentAsNestedTable(cells[3], subRow, "right");
                else
                    SetCellContent(cells[3], new[] { new RunInfo { Text = "" } });
            }
            else
            {
                var renderMode = change.TableAnalysis?.RenderMode ?? "";
                bool isTableChange = change.Type?.StartsWith("table_") == true;

                bool isFullTable =
                    isTableChange &&
                    (
                        renderMode == "table_deleted" ||
                        renderMode == "table_added" ||
                        renderMode == "full_table" ||
                        renderMode == "full" ||
                        change.ChangeKind == "delete" ||
                        change.ChangeKind == "insert"
                    );

                if (isFullTable)
                {
                    bool isDelete = change.ChangeKind == "delete" || change.Type == "table_deleted";
                    bool isInsert = change.ChangeKind == "insert" || change.Type == "table_inserted";

                    var originalCells = change.TableAnalysis?.FullTableOriginal?.Cells;
                    var modifiedCells = change.TableAnalysis?.FullTableModified?.Cells;

                    if (originalCells != null && originalCells.Count > 0)
                        SetCellContentAsFullTable(cells[2], originalCells, isDelete, false);
                    else
                        SetCellContent(cells[2], new[] { new RunInfo { Text = "" } });

                    if (modifiedCells != null && modifiedCells.Count > 0)
                        SetCellContentAsFullTable(cells[3], modifiedCells, false, isInsert);
                    else
                        SetCellContent(cells[3], new[] { new RunInfo { Text = "" } });
                }
                else
                {
                    SetCellContent(cells[2], BuildRunsFromChange(change, "left"));
                    SetCellContent(cells[3], BuildRunsFromChange(change, "right"));
                }
            }
        }

        private static string GetPageText(ChangeDto change)
        {
            int? page =
                change.Right?.Page
                ?? change.Left?.Page
                ?? GetPageFromNode(change.Right?.Node)
                ?? GetPageFromNode(change.Left?.Node);
            return page?.ToString() ?? "";
        }

        private static int? GetPageFromNode(NodeDto? node)
        {
            var content = node?.Content;
            if (content == null) return null;
            if (!content.TryGetValue("page", out var pageEl)) return null;
            if (pageEl.ValueKind == JsonValueKind.Number && pageEl.TryGetInt32(out var n)) return n;
            if (pageEl.ValueKind == JsonValueKind.String && int.TryParse(pageEl.GetString(), out n)) return n;
            return null;
        }

        private void SetCellContent(TableCell cell, IEnumerable<RunInfo> runs, bool center = false)
        {
            foreach (var p in cell.Elements<Paragraph>().ToList()) p.Remove();
            foreach (var t in cell.Elements<Table>().ToList()) t.Remove();

            var paraProps = MakeParaProps();
            if (center)
                paraProps.AppendChild(new Justification { Val = JustificationValues.Center });

            var para = new Paragraph();
            para.AppendChild(paraProps);

            foreach (var runInfo in runs)
            {
                if (!string.IsNullOrWhiteSpace(runInfo.ImageUrl))
                {
                    var imagePath = ResolveImagePath(runInfo.ImageUrl);
                    para.AppendChild(imagePath != null && _mainPart != null
                        ? MakeImageRun(_mainPart, imagePath, 280, 180)
                        : MakeTextRun("[IMAGE NOT FOUND]"));
                }

                if (!string.IsNullOrWhiteSpace(runInfo.Text))
                {
                    if (!string.IsNullOrWhiteSpace(runInfo.ImageUrl))
                        para.AppendChild(new Run(new Text(" ") { Space = SpaceProcessingModeValues.Preserve }));

                    var lines = (runInfo.Text ?? "").Split('\n');
                    for (int i = 0; i < lines.Length; i++)
                    {
                        if (i > 0) para.AppendChild(new Run(new Break()));
                        para.AppendChild(MakeTextRun(lines[i],
                            bold: runInfo.IsLabel,
                            deleted: runInfo.IsDeleted,
                            inserted: runInfo.IsInserted,
                            label: runInfo.IsLabel));
                    }
                }
            }

            cell.AppendChild(para);
        }

        private List<RunInfo> BuildRunsFromChange(ChangeDto change, string direction)
        {
            if (HasOnlyInlineImageChanged(change))
            {
                var leftImgs = GetInlineImageNodes(change.Left?.Node);
                var rightImgs = GetInlineImageNodes(change.Right?.Node);

                var changedImgs = direction == "left"
                    ? leftImgs.Where((img, idx) =>
                        idx >= rightImgs.Count || GetSha(img) != GetSha(rightImgs[idx]))
                    : rightImgs.Where((img, idx) =>
                        idx >= leftImgs.Count || GetSha(img) != GetSha(leftImgs[idx]));

                var result = new List<RunInfo>
    {
        new()
        {
            Text = "[Chỉ thay đổi ảnh inline - nội dung đoạn văn không đổi]",
            IsLabel = true
        }
    };

                foreach (var img in changedImgs)
                {
                    var imageUrl = GetImageUrl(img);

                    if (!string.IsNullOrWhiteSpace(imageUrl))
                    {
                        result.Add(new RunInfo
                        {
                            ImageUrl = imageUrl
                        });
                    }
                }

                return result;
            }

            if (change.Type?.StartsWith("shape") == true)
            {
                var side = direction == "left" ? change.Left : change.Right;
                if (side == null) return new List<RunInfo> { new() { Text = "" } };

                bool isDeleted = direction == "left" && change.ChangeKind == "delete";
                bool isInserted = direction == "right" && change.ChangeKind == "insert";

                var shapeWordDiff = change.ShapeChanges?
                .Select(sc => sc.WordDiff)
                .FirstOrDefault(wd => wd?.Spans?.Count > 0);

                            var contentRuns = GetRunsFromShapeChildren(
                                side.Node,
                                isDeleted,
                                isInserted,
                                shapeWordDiff,
                                direction
                            );

                if (contentRuns.Count == 0)
                {
                    var text = side.PreviewText ?? "";
                    if (text.StartsWith("[SHAPE:") && text.EndsWith("]")) text = text[7..^1].Trim();
                    if (text == "[SHAPE]" || string.IsNullOrWhiteSpace(text)) text = "[Ảnh]";
                    contentRuns.Add(new RunInfo { Text = text, IsDeleted = isDeleted, IsInserted = isInserted });
                }

                contentRuns.Insert(0, new RunInfo { Text = "[Shape] ", IsLabel = true });
                return contentRuns;
            }

            {
                var side = direction == "left" ? change.Left : change.Right;
                if (side == null) return new List<RunInfo> { new() { Text = "" } };

                string label = GetLabel(side.Node?.Content);
                var runs = GetRunsFromWordDiff(change.WordDiff, direction)
                        ?? GetRuns(side.Node?.Content);

                if (runs == null || runs.Count == 0)
                    return new List<RunInfo> { new() { Text = GetPlainText(side, label) } };

                var result = new List<RunInfo>();
                if (!string.IsNullOrEmpty(label))
                    result.Add(new RunInfo { Text = $"{label} ", IsLabel = true });
                result.AddRange(runs);
                return result;
            }
        }

        private static List<RunInfo> GetRunsFromShapeChildren(
    NodeDto? node,
    bool isDeleted = false,
    bool isInserted = false,
    WordDiffDto? wordDiff = null,
    string direction = "left")
        {
            if (node?.Children == null)
                return new();

            var result = new List<RunInfo>();
            bool firstParagraph = true;
            bool usedWordDiff = false;

            foreach (var paragraph in node.Children)
            {
                if (!firstParagraph)
                    result.Add(new RunInfo { Text = "\n" });

                firstParagraph = false;

                // Nếu có text diff → dùng highlight
                if (!usedWordDiff &&
                    wordDiff?.Spans?.Count > 0)
                {
                    var diffRuns =
                        GetRunsFromWordDiff(wordDiff, direction);

                    if (diffRuns != null)
                        result.AddRange(diffRuns);

                    usedWordDiff = true;
                }
                else
                {
                    // Text thường
                    if (paragraph.Content != null &&
                        paragraph.Content.TryGetValue("text", out var tx))
                    {
                        var t = tx.GetString() ?? "";

                        if (!string.IsNullOrWhiteSpace(t))
                        {
                            result.Add(new RunInfo
                            {
                                Text = t,
                                IsDeleted = isDeleted,
                                IsInserted = isInserted
                            });
                        }
                    }
                }

                // Luôn export ảnh
                if (paragraph.Children == null)
                    continue;

                foreach (var child in paragraph.Children)
                {
                    if (child.Type != "image")
                        continue;

                    string? imageUrl = null;

                    if (child.Content != null &&
                        child.Content.TryGetValue("image_url", out var iu))
                    {
                        imageUrl = iu.GetString();
                    }

                    imageUrl ??= child.ImageUrl;

                    if (!string.IsNullOrWhiteSpace(imageUrl))
                    {
                        result.Add(new RunInfo
                        {
                            ImageUrl = imageUrl,
                            IsDeleted = isDeleted,
                            IsInserted = isInserted
                        });
                    }
                }
            }

            return result;
        }

        private static List<RunInfo>? GetRunsFromWordDiff(WordDiffDto? wordDiff, string direction)
        {
            if (wordDiff?.Spans == null || wordDiff.Spans.Count == 0) return null;
            var list = new List<RunInfo>();
            bool isFirst = true;

            foreach (var span in wordDiff.Spans)
            {
                string prefix = !isFirst ? " " : "";
                switch (span.Type)
                {
                    case "equal":
                        list.Add(new RunInfo { Text = prefix + (span.Text ?? "") });
                        break;
                    case "delete":
                        if (direction == "left")
                            list.Add(new RunInfo { Text = prefix + (span.OldText ?? span.Text ?? ""), IsDeleted = true });
                        break;
                    case "insert":
                        if (direction == "right")
                            list.Add(new RunInfo { Text = prefix + (span.NewText ?? span.Text ?? ""), IsInserted = true });
                        break;
                    case "replace":
                        if (direction == "left")
                            list.Add(new RunInfo { Text = prefix + (span.OldText ?? ""), IsDeleted = true });
                        else
                            list.Add(new RunInfo { Text = prefix + (span.NewText ?? ""), IsInserted = true });
                        break;
                }
                isFirst = false;
            }
            return list.Count > 0 ? list : null;
        }

        private static List<RunInfo>? GetRuns(Dictionary<string, JsonElement>? content)
        {
            if (content == null) return null;
            var result = new List<RunInfo>();

            string? imageUrl = null;
            if (content.TryGetValue("image_url", out var iu)) imageUrl = iu.GetString();
            else if (content.TryGetValue("url", out var ul)) imageUrl = ul.GetString();
            else if (content.TryGetValue("src", out var sc)) imageUrl = sc.GetString();
            else if (content.TryGetValue("path", out var pt)) imageUrl = pt.GetString();

            if (!string.IsNullOrWhiteSpace(imageUrl))
                result.Add(new RunInfo { ImageUrl = imageUrl });

            if (content.TryGetValue("runs", out var runsEl) && runsEl.ValueKind == JsonValueKind.Array)
            {
                foreach (var r in runsEl.EnumerateArray())
                    result.Add(new RunInfo
                    {
                        Text = r.TryGetProperty("text", out var t) ? t.GetString() ?? "" : "",
                        IsDeleted = r.TryGetProperty("deleted", out var d) && d.GetBoolean(),
                        IsInserted = r.TryGetProperty("inserted", out var ins) && ins.GetBoolean(),
                    });
            }
            else if (result.Count == 0 || !string.IsNullOrWhiteSpace(imageUrl))
            {
                if (content.TryGetValue("text", out var tx) && tx.ValueKind == JsonValueKind.String)
                {
                    var text = tx.GetString() ?? "";
                    if (!string.IsNullOrEmpty(text)) result.Add(new RunInfo { Text = text });
                }
                else if (content.TryGetValue("text_display", out var td) && td.ValueKind == JsonValueKind.String)
                {
                    var text = td.GetString() ?? "";
                    if (!string.IsNullOrEmpty(text)) result.Add(new RunInfo { Text = text });
                }
            }
            return result.Count > 0 ? result : null;
        }

        private static string GetLabel(Dictionary<string, JsonElement>? content)
        {
            if (content == null) return "";
            if (content.TryGetValue("numbering", out var numEl) &&
                numEl.ValueKind == JsonValueKind.Object &&
                numEl.TryGetProperty("label", out var labelEl))
                return labelEl.GetString() ?? "";
            return "";
        }

        private static string GetPlainText(BlockSideDto side, string label)
        {
            string text = "";
            var content = side.Node?.Content;
            if (content != null)
            {
                if (content.TryGetValue("text_display", out var td)) text = td.GetString() ?? "";
                else if (content.TryGetValue("text", out var tx)) text = tx.GetString() ?? "";
                else if (content.TryGetValue("image_url", out var iu)) text = iu.GetString() ?? "";
                else if (content.TryGetValue("url", out var ul)) text = ul.GetString() ?? "";
                else if (content.TryGetValue("src", out var sc)) text = sc.GetString() ?? "";
            }
            if (string.IsNullOrEmpty(text)) text = side.PreviewText ?? "";
            return string.IsNullOrEmpty(label) ? text : $"{label} {text}";
        }

        private static ParagraphProperties MakeParaProps() =>
            new ParagraphProperties(
                new SpacingBetweenLines { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }
            );

        private static Run MakeTextRun(
            string text,
            bool bold = false,
            bool deleted = false,
            bool inserted = false,
            bool label = false)
        {
            var rPr = new RunProperties();
            rPr.AppendChild(new RunFonts { Ascii = Font, HighAnsi = Font, ComplexScript = Font });
            rPr.AppendChild(new FontSize { Val = FontSize });
            rPr.AppendChild(new FontSizeComplexScript { Val = FontSize });

            if (bold) rPr.AppendChild(new Bold());

            if (deleted)
            {
                rPr.AppendChild(new Color { Val = ColorDeleted });
                rPr.AppendChild(new Highlight { Val = HighlightColorValues.Yellow });
            }
            else if (inserted)
            {
                rPr.AppendChild(new Color { Val = ColorInserted });
                rPr.AppendChild(new Highlight { Val = HighlightColorValues.Cyan });
            }
            else if (label)
            {
                rPr.AppendChild(new Color { Val = ColorLabel });
            }

            return new Run(rPr, new Text(text ?? "") { Space = SpaceProcessingModeValues.Preserve });
        }

        private string? ResolveImagePath(string? imageUrl)
        {
            if (string.IsNullOrWhiteSpace(imageUrl)) return null;
            if (File.Exists(imageUrl)) return imageUrl;
            var fileName = Path.GetFileName(imageUrl);
            if (string.IsNullOrWhiteSpace(fileName)) return null;
            var path = Path.Combine(_imageDir, fileName);
            return File.Exists(path) ? path : null;
        }

        private Run MakeImageRun(MainDocumentPart mainPart, string imagePath, long widthPx, long heightPx)
        {
            var ext = Path.GetExtension(imagePath).ToLowerInvariant();
            var partType = ext switch
            {
                ".jpg" or ".jpeg" => ImagePartType.Jpeg,
                ".gif" => ImagePartType.Gif,
                ".bmp" => ImagePartType.Bmp,
                _ => ImagePartType.Png
            };

            var imagePart = mainPart.AddImagePart(partType);
            using (var stream = File.OpenRead(imagePath))
                imagePart.FeedData(stream);

            var relId = mainPart.GetIdOfPart(imagePart);
            long wEmu = widthPx * 9525;
            long hEmu = heightPx * 9525;

            return new Run(new Drawing(
                new DW.Inline(
                    new DW.Extent { Cx = wEmu, Cy = hEmu },
                    new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                    new DW.DocProperties { Id = _imageId++, Name = Path.GetFileName(imagePath) },
                    new DW.NonVisualGraphicFrameDrawingProperties(
                        new A.GraphicFrameLocks { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties { Id = 0U, Name = Path.GetFileName(imagePath) },
                                    new PIC.NonVisualPictureDrawingProperties()),
                                new PIC.BlipFill(
                                    new A.Blip { Embed = relId },
                                    new A.Stretch(new A.FillRectangle())),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset { X = 0, Y = 0 },
                                        new A.Extents { Cx = wEmu, Cy = hEmu }),
                                    new A.PresetGeometry(new A.AdjustValueList())
                                    { Preset = A.ShapeTypeValues.Rectangle })))
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                )
                { DistanceFromTop = 0, DistanceFromBottom = 0, DistanceFromLeft = 0, DistanceFromRight = 0 }));
        }

        private class RunInfo
        {
            public string Text { get; set; } = "";
            public bool IsDeleted { get; set; }
            public bool IsInserted { get; set; }
            public bool IsLabel { get; set; }
            public string? ImageUrl { get; set; }
        }
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // DTOs
    // ═══════════════════════════════════════════════════════════════════════════

    public class ChecksheetRequest
    {
        [JsonPropertyName("job_id")] public string? JobId { get; set; }
        [JsonPropertyName("original_file")] public string? OriginalFile { get; set; }
        [JsonPropertyName("modified_file")] public string? ModifiedFile { get; set; }
        [JsonPropertyName("sections")] public List<SectionDto> Sections { get; set; } = new();
    }

    public class SectionDto
    {
        [JsonPropertyName("heading")] public string? Heading { get; set; }
        [JsonPropertyName("changes")] public List<ChangeDto> Changes { get; set; } = new();
    }

    public class ChangeDto
    {
        [JsonPropertyName("heading")] public string? Heading { get; set; }
        [JsonPropertyName("type")] public string? Type { get; set; }
        [JsonPropertyName("change_kind")] public string? ChangeKind { get; set; }
        [JsonPropertyName("left")] public BlockSideDto? Left { get; set; }
        [JsonPropertyName("right")] public BlockSideDto? Right { get; set; }
        [JsonPropertyName("word_diff")] public WordDiffDto? WordDiff { get; set; }
        [JsonPropertyName("shape_changes")] public List<ShapeChangeDto>? ShapeChanges { get; set; }
        [JsonPropertyName("shape_text_lines")] public List<string>? ShapeTextLines { get; set; }
        [JsonPropertyName("table_analysis")] public TableAnalysisDto? TableAnalysis { get; set; }
    }

    public class ShapeChangeDto
    {
        [JsonPropertyName("word_diff")] public WordDiffDto? WordDiff { get; set; }
        [JsonPropertyName("original")] public ShapeNodeDto? Original { get; set; }
        [JsonPropertyName("modified")] public ShapeNodeDto? Modified { get; set; }
    }

    public class ShapeNodeDto
    {
        [JsonPropertyName("content")] public Dictionary<string, JsonElement>? Content { get; set; }
    }

    public class BlockSideDto
    {
        [JsonPropertyName("preview_text")] public string? PreviewText { get; set; }
        [JsonPropertyName("page")] public int? Page { get; set; }
        [JsonPropertyName("node")] public NodeDto? Node { get; set; }
    }

    public class NodeDto
    {
        [JsonPropertyName("type")] public string? Type { get; set; }
        [JsonPropertyName("image_url")] public string? ImageUrl { get; set; }
        [JsonPropertyName("content")] public Dictionary<string, JsonElement>? Content { get; set; }
        [JsonPropertyName("children")] public List<NodeDto>? Children { get; set; }
    }

    public class WordDiffDto
    {
        [JsonPropertyName("spans")] public List<SpanDto>? Spans { get; set; }
    }

    public class SpanDto
    {
        [JsonPropertyName("type")] public string? Type { get; set; }
        [JsonPropertyName("old_text")] public string? OldText { get; set; }
        [JsonPropertyName("new_text")] public string? NewText { get; set; }
        [JsonPropertyName("text")] public string? Text { get; set; }
        [JsonPropertyName("space_before")] public bool SpaceBefore { get; set; }
    }

    public class TableAnalysisDto
    {
        [JsonPropertyName("render_mode")] public string? RenderMode { get; set; }
        [JsonPropertyName("full_table_original")] public FullTableDto? FullTableOriginal { get; set; }
        [JsonPropertyName("full_table_modified")] public FullTableDto? FullTableModified { get; set; }
        [JsonPropertyName("table_changes")] public List<TableRowChangeDto> TableChanges { get; set; } = new();
    }

    public class FullTableDto
    {
        [JsonPropertyName("total_rows")] public int TotalRows { get; set; }
        [JsonPropertyName("total_cols")] public int TotalCols { get; set; }
        [JsonPropertyName("cells")] public List<CellDto> Cells { get; set; } = new();
    }

    public class TableRowChangeDto
    {
        [JsonPropertyName("type")] public string? Type { get; set; }
        [JsonPropertyName("anchor_row")] public int? AnchorRow { get; set; }
        [JsonPropertyName("left_cells")] public List<CellDto>? LeftCells { get; set; }
        [JsonPropertyName("right_cells")] public List<CellDto>? RightCells { get; set; }
        [JsonPropertyName("left_display_cells")] public List<CellDto>? LeftDisplayCells { get; set; }
        [JsonPropertyName("right_display_cells")] public List<CellDto>? RightDisplayCells { get; set; }
        [JsonPropertyName("cell_changes")] public List<CellChangeDto>? CellChanges { get; set; }
    }

    public class CellDto
    {
        [JsonPropertyName("anchor_col")] public int? AnchorCol { get; set; }
        [JsonPropertyName("anchor_row")] public int? AnchorRow { get; set; }
        [JsonPropertyName("col_span")] public int? ColSpan { get; set; }
        [JsonPropertyName("row_span")] public int? RowSpan { get; set; }
        [JsonPropertyName("text")] public string? Text { get; set; }
        [JsonPropertyName("text_display")] public string? TextDisplay { get; set; }
        [JsonPropertyName("content_blocks")] public List<ContentBlockDto>? ContentBlocks { get; set; }
        [JsonPropertyName("standalone_images")] public List<ContentBlockPayloadDto>? StandaloneImages { get; set; }
    }

    public class CellChangeDto
    {
        [JsonPropertyName("anchor_col")] public int? AnchorCol { get; set; }
        [JsonPropertyName("type")] public string? Type { get; set; }
        [JsonPropertyName("changes")] public List<ParagraphChangeDto>? Changes { get; set; }
        [JsonPropertyName("left_cell")] public CellDto? LeftCell { get; set; }
        [JsonPropertyName("right_cell")] public CellDto? RightCell { get; set; }
    }

    public class ParagraphChangeDto
    {
        [JsonPropertyName("type")] public string? Type { get; set; }
        [JsonPropertyName("spans")] public List<SpanDto>? Spans { get; set; }
        [JsonPropertyName("original")] public ContentBlockPayloadDto? Original { get; set; }
        [JsonPropertyName("modified")] public ContentBlockPayloadDto? Modified { get; set; }
        [JsonPropertyName("change_kind")] public string? ChangeKind { get; set; }

        // dùng cho nested_table_modified
        [JsonPropertyName("changes")] public List<TableRowChangeDto>? Changes { get; set; }
    }

    public class ContentBlockDto
    {
        [JsonPropertyName("type")] public string? Type { get; set; }
        [JsonPropertyName("payload")] public ContentBlockPayloadDto? Payload { get; set; }
    }

    public class ContentBlockPayloadDto
    {
        [JsonPropertyName("uid")] public string? Uid { get; set; }
        [JsonPropertyName("text")] public string? Text { get; set; }
        [JsonPropertyName("text_display")] public string? TextDisplay { get; set; }
        [JsonPropertyName("images")] public List<InlineImageDto>? Images { get; set; }
        [JsonPropertyName("image_url")] public string? ImageUrl { get; set; }
        [JsonPropertyName("sha256")] public string? Sha256 { get; set; }
        [JsonPropertyName("width_px")] public int? WidthPx { get; set; }
        [JsonPropertyName("height_px")] public int? HeightPx { get; set; }
        [JsonPropertyName("total_rows")] public int? TotalRows { get; set; }
        [JsonPropertyName("total_cols")] public int? TotalCols { get; set; }
        [JsonPropertyName("cells")] public List<CellDto>? Cells { get; set; }
    }

    public class InlineImageDto
    {
        [JsonPropertyName("image_url")] public string? ImageUrl { get; set; }
        [JsonPropertyName("sha256")] public string? Sha256 { get; set; }
    }

}