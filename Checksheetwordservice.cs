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
        private MainDocumentPart? _mainPart;
        private uint _imageId = 1;
        private readonly string _imageDir;

        public ChecksheetWordService(IConfiguration config)
        {
            _imageDir = config["ImageStore:Dir"]
                ?? throw new Exception("Thiếu config ImageStore:Dir trong appsettings.json");
        }

        private const string ColorDeleted = "FF0000";
        private const string ColorInserted = "00B050";
        private const string ColorLabel = "6B7280";
        private const string Font = "Times New Roman";
        private const string FontSize = "22"; // half-points = 11pt

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

                var allChanges = sections
                    .SelectMany(s => (s.Changes ?? new())
                        .Select(c => (heading: c.Heading ?? s.Heading ?? "", change: c)))
                    .ToList();

                if (allChanges.Count > 0)
                {
                    FillDataRow(templateRow, allChanges[0].heading, allChanges[0].change);

                    for (int i = 1; i < allChanges.Count; i++)
                    {
                        var cloned = (TableRow)templateRow.CloneNode(deep: true);
                        FillDataRow(cloned, allChanges[i].heading, allChanges[i].change);
                        table.AppendChild(cloned);
                    }
                }

                _mainPart.Document!.Save();
            }

            return ms.ToArray();
        }

        // ── Row filler ──────────────────────────────────────────────────────────

        private void FillDataRow(TableRow row, string heading, ChangeDto change)
        {
            var cells = row.Elements<TableCell>().ToList();
            if (cells.Count < 4) return;

            // Cột 0: Trang sửa đổi
            SetCellContent(cells[0], new[] { new RunInfo { Text = GetPageText(change) } }, center: true);

            // Cột 1: Đầu mục
            SetCellContent(cells[1], new[] { new RunInfo { Text = heading } });

            // Cột 2: Trước sửa đổi
            SetCellContent(cells[2], BuildRunsFromChange(change, "left"));

            // Cột 3: Sau sửa đổi
            SetCellContent(cells[3], BuildRunsFromChange(change, "right"));

            // Cột 4 & 5: Đào tạo thực hành / Mục đích sửa đổi — giữ nguyên template
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

            if (!content.TryGetValue("page", out var pageEl))
                return null;

            if (pageEl.ValueKind == JsonValueKind.Number &&
                pageEl.TryGetInt32(out var pageNum))
                return pageNum;

            if (pageEl.ValueKind == JsonValueKind.String &&
                int.TryParse(pageEl.GetString(), out pageNum))
                return pageNum;

            return null;
        }

        // ── Cell writer ─────────────────────────────────────────────────────────

        private void SetCellContent(TableCell cell, IEnumerable<RunInfo> runs, bool center = false)
        {
            foreach (var p in cell.Elements<Paragraph>().ToList())
                p.Remove();

            var paraProps = new ParagraphProperties(
                new SpacingBetweenLines { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }
            );
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

                    var parts = (runInfo.Text ?? "").Split('\n');
                    for (int i = 0; i < parts.Length; i++)
                    {
                        if (i > 0) para.AppendChild(new Run(new Break()));
                        para.AppendChild(MakeTextRun(parts[i],
                            bold: runInfo.IsLabel,
                            deleted: runInfo.IsDeleted,
                            inserted: runInfo.IsInserted,
                            label: runInfo.IsLabel));
                    }
                }
            }

            cell.AppendChild(para);
        }

        // ── Run builders ────────────────────────────────────────────────────────

        private List<RunInfo> BuildRunsFromChange(ChangeDto change, string direction)
        {
            // ── Shape modified: có shape_changes[].word_diff ───────────────────────
            if (change.ShapeChanges?.Count > 0)
            {
                var shapeRuns = new List<RunInfo>();
                foreach (var sc in change.ShapeChanges)
                {
                    var runs = GetRunsFromWordDiff(sc.WordDiff, direction);
                    if (runs != null) shapeRuns.AddRange(runs);
                }
                if (shapeRuns.Count > 0)
                {
                    shapeRuns.Insert(0, new RunInfo { Text = "[Shape] ", IsLabel = true });
                    return shapeRuns;
                }
                // word_diff rỗng (e.g. chỉ thay đổi format) → fallthrough xuống shape deleted/added
            }

            // ── Shape deleted / inserted: lấy content từ children đệ quy ──────────
            if (change.Type?.StartsWith("shape") == true)
            {
                var side = direction == "left" ? change.Left : change.Right;
                if (side == null) return new List<RunInfo> { new() { Text = "" } };

                bool isDeleted = direction == "left" && change.ChangeKind == "delete";
                bool isInserted = direction == "right" && change.ChangeKind == "insert";

                var contentRuns = GetRunsFromShapeChildren(side.Node, isDeleted, isInserted);

                if (contentRuns.Count == 0)
                {
                    var text = side.PreviewText ?? "";
                    if (text.StartsWith("[SHAPE:") && text.EndsWith("]"))
                        text = text[7..^1].Trim();
                    if (text == "[SHAPE]" || string.IsNullOrWhiteSpace(text))
                        text = "[Ảnh]";
                    contentRuns.Add(new RunInfo { Text = text, IsDeleted = isDeleted, IsInserted = isInserted });
                }

                contentRuns.Insert(0, new RunInfo { Text = "[Shape] ", IsLabel = true });
                return contentRuns;
            }

            // ── Paragraph: logic cũ ────────────────────────────────────────────────
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

        // THAY THẾ GetTextFromShapeChildren bằng method này
        private static List<RunInfo> GetRunsFromShapeChildren(
            NodeDto? node, bool isDeleted = false, bool isInserted = false)
        {
            if (node?.Children == null) return new();
            var result = new List<RunInfo>();
            bool firstParagraph = true;

            foreach (var paragraph in node.Children)
            {
                // Ngắt dòng giữa các paragraph
                if (!firstParagraph)
                    result.Add(new RunInfo { Text = "\n" });
                firstParagraph = false;

                // Text của paragraph
                if (paragraph.Content != null &&
                    paragraph.Content.TryGetValue("text", out var tx))
                {
                    var t = tx.GetString() ?? "";
                    if (!string.IsNullOrEmpty(t))
                        result.Add(new RunInfo { Text = t, IsDeleted = isDeleted, IsInserted = isInserted });
                }

                // Children của paragraph: tìm image nodes
                if (paragraph.Children == null) continue;
                foreach (var child in paragraph.Children)
                {
                    if (child.Type != "image") continue;

                    string? imageUrl = null;
                    if (child.Content != null &&
                        child.Content.TryGetValue("image_url", out var iu))
                        imageUrl = iu.GetString();
                    if (string.IsNullOrEmpty(imageUrl))
                        imageUrl = child.ImageUrl;

                    if (!string.IsNullOrEmpty(imageUrl))
                        result.Add(new RunInfo { ImageUrl = imageUrl, IsDeleted = isDeleted, IsInserted = isInserted });
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

            // Detect image URL
            string? imageUrl = null;
            if (content.TryGetValue("image_url", out var iu)) imageUrl = iu.GetString();
            else if (content.TryGetValue("url", out var ul)) imageUrl = ul.GetString();
            else if (content.TryGetValue("src", out var sc)) imageUrl = sc.GetString();
            else if (content.TryGetValue("path", out var pt)) imageUrl = pt.GetString();

            if (!string.IsNullOrWhiteSpace(imageUrl))
                result.Add(new RunInfo { ImageUrl = imageUrl });

            // Đọc runs kèm theo (text trong shape có thể nằm ở đây)
            if (content.TryGetValue("runs", out var runsEl) && runsEl.ValueKind == JsonValueKind.Array)
            {
                foreach (var r in runsEl.EnumerateArray())
                    result.Add(new RunInfo
                    {
                        Text = r.TryGetProperty("text", out var t) ? t.GetString() ?? "" : "",
                        IsDeleted = r.TryGetProperty("deleted", out var d) && d.GetBoolean(),
                        IsInserted = r.TryGetProperty("inserted", out var i) && i.GetBoolean(),
                    });
            }
            // Fallback: text đơn trong content (không có runs array)
            else if (result.Count == 0 || !string.IsNullOrWhiteSpace(imageUrl))
            {
                if (content.TryGetValue("text", out var tx) && tx.ValueKind == JsonValueKind.String)
                {
                    var text = tx.GetString() ?? "";
                    if (!string.IsNullOrEmpty(text))
                        result.Add(new RunInfo { Text = text });
                }
                else if (content.TryGetValue("text_display", out var td) && td.ValueKind == JsonValueKind.String)
                {
                    var text = td.GetString() ?? "";
                    if (!string.IsNullOrEmpty(text))
                        result.Add(new RunInfo { Text = text });
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

        // ── Text / Image run factories ──────────────────────────────────────────

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
    // Shared DTOs
    // ═══════════════════════════════════════════════════════════════════════════

    public class ChecksheetRequest
    {
        [JsonPropertyName("job_id")]
        public string? JobId { get; set; }

        [JsonPropertyName("original_file")]
        public string? OriginalFile { get; set; }

        [JsonPropertyName("modified_file")]
        public string? ModifiedFile { get; set; }

        [JsonPropertyName("sections")]
        public List<SectionDto> Sections { get; set; } = new();
    }

    public class SectionDto
    {
        [JsonPropertyName("heading")]
        public string? Heading { get; set; }

        [JsonPropertyName("changes")]
        public List<ChangeDto> Changes { get; set; } = new();
    }

    public class ChangeDto
    {
        [JsonPropertyName("heading")]
        public string? Heading { get; set; }

        [JsonPropertyName("type")]
        public string? Type { get; set; }

        [JsonPropertyName("change_kind")]
        public string? ChangeKind { get; set; }

        [JsonPropertyName("left")]
        public BlockSideDto? Left { get; set; }

        [JsonPropertyName("right")]
        public BlockSideDto? Right { get; set; }

        [JsonPropertyName("word_diff")]
        public WordDiffDto? WordDiff { get; set; }

        [JsonPropertyName("shape_changes")]
        public List<ShapeChangeDto>? ShapeChanges { get; set; }

        [JsonPropertyName("shape_text_lines")]
        public List<string>? ShapeTextLines { get; set; }
    }
    public class ShapeChangeDto
    {
        [JsonPropertyName("word_diff")]
        public WordDiffDto? WordDiff { get; set; }

        [JsonPropertyName("original")]
        public ShapeNodeDto? Original { get; set; }

        [JsonPropertyName("modified")]
        public ShapeNodeDto? Modified { get; set; }
    }

    public class ShapeNodeDto
    {
        [JsonPropertyName("content")]
        public Dictionary<string, JsonElement>? Content { get; set; }
    }

    public class BlockSideDto
    {
        [JsonPropertyName("preview_text")]
        public string? PreviewText { get; set; }

        [JsonPropertyName("page")]
        public int? Page { get; set; }

        [JsonPropertyName("node")]
        public NodeDto? Node { get; set; }

    }

    public class NodeDto
    {
        [JsonPropertyName("type")]
        public string? Type { get; set; }         

        [JsonPropertyName("image_url")]
        public string? ImageUrl { get; set; }
        [JsonPropertyName("content")]
        public Dictionary<string, JsonElement>? Content { get; set; }

        [JsonPropertyName("children")]
        public List<NodeDto>? Children { get; set; }
    }

    public class WordDiffDto
    {
        [JsonPropertyName("spans")]
        public List<SpanDto>? Spans { get; set; }
    }

    public class SpanDto
    {
        [JsonPropertyName("type")]
        public string? Type { get; set; }

        [JsonPropertyName("old_text")]
        public string? OldText { get; set; }

        [JsonPropertyName("new_text")]
        public string? NewText { get; set; }

        [JsonPropertyName("text")]
        public string? Text { get; set; }

        [JsonPropertyName("space_before")]
        public bool SpaceBefore { get; set; }
    }
}