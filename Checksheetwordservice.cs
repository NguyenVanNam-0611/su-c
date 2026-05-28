using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.Json;

using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace Diff_tool.Services
{
    public class ChecksheetWordService
    {
        private MainDocumentPart? _mainPart;
        private uint _imageId = 1;

        private const string ColorDeleted = "FF0000";
        private const string ColorInserted = "00B050";
        private const string ColorLabel = "6B7280";

        private const string Font = "Times New Roman";
        private const string FontSize = "22";

        public byte[] GenerateChecksheet(List<SectionDto> sections)
        {
            using var ms = new MemoryStream();

            using (var wordDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document, true))
            {
                var mainPart = wordDoc.AddMainDocumentPart();
                _mainPart = mainPart;

                mainPart.Document = new Document(new Body());
                var body = mainPart.Document.Body!;

                AddStyles(mainPart);

                body.AppendChild(MakeTitle("CHECKSHEET SỬA ĐỔI TÀI LIỆU"));

                var table = MakeTable();
                table.AppendChild(MakeHeaderRow());

                bool hasRows = false;

                foreach (var section in sections)
                {
                    if (section.Changes == null || section.Changes.Count == 0)
                        continue;

                    foreach (var change in section.Changes)
                    {
                        string heading = change.Heading ?? section.Heading ?? "";
                        table.AppendChild(MakeDataRow(heading, change));
                        hasRows = true;
                    }
                }

                if (!hasRows)
                    table.AppendChild(MakeEmptyRow());

                body.AppendChild(table);

                body.AppendChild(new Paragraph());
                body.AppendChild(MakeSectionProperties());

                mainPart.Document.Save();
            }

            return ms.ToArray();
        }

        private static string? ResolveImagePath(string? imageUrl)
        {
            if (string.IsNullOrWhiteSpace(imageUrl))
                return null;

            if (File.Exists(imageUrl))
                return imageUrl;

            var fileName = Path.GetFileName(imageUrl);
            if (string.IsNullOrWhiteSpace(fileName))
                return null;

            var tmpDir =
                Environment.GetEnvironmentVariable("DOCX_IMAGE_TMP")
                ?? @"D:\comparison_tool\tmp\docx_images";

            var path = Path.Combine(tmpDir, fileName);

            Console.WriteLine($"IMAGE URL  = {imageUrl}");
            Console.WriteLine($"IMAGE PATH = {path}");
            Console.WriteLine($"EXISTS     = {File.Exists(path)}");

            return File.Exists(path) ? path : null;
        }

        private static Table MakeTable()
        {
            int[] colWidths = { 1000, 2000, 3500, 3500, 1500, 1500 };

            var tblPr = new TableProperties(
                new TableWidth { Width = "13000", Type = TableWidthUnitValues.Dxa },
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new BottomBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new LeftBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new RightBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = "000000" }
                )
            );

            var grid = new TableGrid(
                colWidths.Select(w => new GridColumn { Width = w.ToString() }).ToArray()
            );

            var table = new Table();
            table.AppendChild(tblPr);
            table.AppendChild(grid);

            return table;
        }

        private TableRow MakeHeaderRow()
        {
            string[] headers =
            {
                "Trang\nsửa đổi",
                "Đầu mục",
                "Trước sửa đổi",
                "Sau sửa đổi",
                "Đào tạo\nthực hành",
                "Mục đích\nsửa đổi"
            };

            int[] colWidths = { 1000, 2000, 3500, 3500, 1500, 1500 };

            var row = new TableRow();

            for (int i = 0; i < headers.Length; i++)
            {
                row.AppendChild(MakeCell(
                    colWidths[i],
                    new[] { new RunInfo { Text = headers[i] } },
                    shading: "D9D9D9",
                    bold: true,
                    center: true
                ));
            }

            return row;
        }

        private TableRow MakeDataRow(string heading, ChangeDto change)
        {
            int[] colWidths = { 1000, 2000, 3500, 3500, 1500, 1500 };

            var leftRuns = BuildRunsFromChange(change, "left");
            var rightRuns = BuildRunsFromChange(change, "right");

            var row = new TableRow();

            row.AppendChild(MakeCell(colWidths[0], new[] { new RunInfo { Text = "" } }));
            row.AppendChild(MakeCell(colWidths[1], new[] { new RunInfo { Text = heading } }));
            row.AppendChild(MakeCell(colWidths[2], leftRuns));
            row.AppendChild(MakeCell(colWidths[3], rightRuns));
            row.AppendChild(MakeCell(colWidths[4], new[] { new RunInfo { Text = "" } }));
            row.AppendChild(MakeCell(colWidths[5], new[] { new RunInfo { Text = "" } }));

            return row;
        }

        private TableRow MakeEmptyRow()
        {
            int[] colWidths = { 1000, 2000, 3500, 3500, 1500, 1500 };
            var row = new TableRow();

            foreach (var width in colWidths)
                row.AppendChild(MakeCell(width, new[] { new RunInfo { Text = " " } }));

            return row;
        }

        private TableCell MakeCell(
            int width,
            IEnumerable<RunInfo> runs,
            string? shading = null,
            bool bold = false,
            bool center = false)
        {
            var cellPr = new TableCellProperties(
                new TableCellWidth { Width = width.ToString(), Type = TableWidthUnitValues.Dxa },
                new TableCellMargin(
                    new TopMargin { Width = "60", Type = TableWidthUnitValues.Dxa },
                    new BottomMargin { Width = "60", Type = TableWidthUnitValues.Dxa },
                    new LeftMargin { Width = "100", Type = TableWidthUnitValues.Dxa },
                    new RightMargin { Width = "100", Type = TableWidthUnitValues.Dxa }
                ),
                new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }
            );

            if (shading != null)
                cellPr.AppendChild(new Shading { Val = ShadingPatternValues.Clear, Fill = shading, Color = "auto" });

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

                    if (imagePath != null && _mainPart != null)
                    {
                        para.AppendChild(MakeImageRun(_mainPart, imagePath, 280, 180));
                    }
                    else
                    {
                        para.AppendChild(MakeTextRun("[IMAGE NOT FOUND]", bold: false));
                    }

                    continue;
                }

                var parts = (runInfo.Text ?? "").Split('\n');

                for (int i = 0; i < parts.Length; i++)
                {
                    if (i > 0)
                        para.AppendChild(new Run(new Break()));

                    para.AppendChild(MakeTextRun(
                        parts[i],
                        bold || runInfo.IsLabel,
                        runInfo.IsDeleted,
                        runInfo.IsInserted,
                        runInfo.IsLabel
                    ));
                }
            }

            var cell = new TableCell();
            cell.AppendChild(cellPr);
            cell.AppendChild(para);

            return cell;
        }

        private static Run MakeTextRun(
            string text,
            bool bold = false,
            bool deleted = false,
            bool inserted = false,
            bool label = false)
        {
            var rPr = new RunProperties();

            rPr.AppendChild(new RunFonts
            {
                Ascii = Font,
                HighAnsi = Font,
                ComplexScript = Font
            });

            rPr.AppendChild(new FontSize { Val = FontSize });
            rPr.AppendChild(new FontSizeComplexScript { Val = FontSize });

            if (bold)
                rPr.AppendChild(new Bold());

            if (deleted)
            {
                rPr.AppendChild(new Color { Val = ColorDeleted });
                rPr.AppendChild(new Highlight { Val = HighlightColorValues.Yellow });
                rPr.AppendChild(new Strike());
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

            return new Run(
                rPr,
                new Text(text ?? "") { Space = SpaceProcessingModeValues.Preserve }
            );
        }

        private List<RunInfo> BuildRunsFromChange(ChangeDto change, string direction)
        {
            var side = direction == "left" ? change.Left : change.Right;

            if (side == null)
                return new List<RunInfo> { new() { Text = "" } };

            string label = GetLabel(side.Node?.Content);

            var runs =
                GetRunsFromWordDiff(change.WordDiff, direction)
                ?? GetRuns(side.Node?.Content);

            if (runs == null || runs.Count == 0)
            {
                string plain = GetPlainText(side, label);
                return new List<RunInfo> { new() { Text = plain } };
            }

            var result = new List<RunInfo>();

            if (!string.IsNullOrEmpty(label))
                result.Add(new RunInfo { Text = $"{label} ", IsLabel = true });

            result.AddRange(runs);

            return result;
        }

        private static List<RunInfo>? GetRunsFromWordDiff(WordDiffDto? wordDiff, string direction)
        {
            if (wordDiff?.Spans == null || wordDiff.Spans.Count == 0)
                return null;

            var list = new List<RunInfo>();
            bool isFirst = true;

            foreach (var span in wordDiff.Spans)
            {
                string prefix = span.SpaceBefore && !isFirst ? " " : "";

                switch (span.Type)
                {
                    case "equal":
                        list.Add(new RunInfo { Text = prefix + (span.Text ?? "") });
                        break;

                    case "delete":
                        if (direction == "left")
                            list.Add(new RunInfo { Text = prefix + (span.OldText ?? span.Text ?? ""), IsDeleted = true });
                        else if (prefix != "")
                            list.Add(new RunInfo { Text = prefix });
                        break;

                    case "insert":
                        if (direction == "right")
                            list.Add(new RunInfo { Text = prefix + (span.NewText ?? span.Text ?? ""), IsInserted = true });
                        else if (prefix != "")
                            list.Add(new RunInfo { Text = prefix });
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
            if (content == null)
                return null;

            string? imageUrl = null;

            if (content.TryGetValue("image_url", out var imageUrlEl))
                imageUrl = imageUrlEl.GetString();
            else if (content.TryGetValue("url", out var urlEl))
                imageUrl = urlEl.GetString();
            else if (content.TryGetValue("src", out var srcEl))
                imageUrl = srcEl.GetString();
            else if (content.TryGetValue("path", out var pathEl))
                imageUrl = pathEl.GetString();

            if (!string.IsNullOrWhiteSpace(imageUrl))
            {
                return new List<RunInfo>
                {
                    new RunInfo { ImageUrl = imageUrl }
                };
            }

            if (!content.TryGetValue("runs", out var runsEl))
                return null;

            if (runsEl.ValueKind != JsonValueKind.Array)
                return null;

            var list = new List<RunInfo>();

            foreach (var r in runsEl.EnumerateArray())
            {
                list.Add(new RunInfo
                {
                    Text = r.TryGetProperty("text", out var t) ? t.GetString() ?? "" : "",
                    IsDeleted = r.TryGetProperty("deleted", out var d) && d.GetBoolean(),
                    IsInserted = r.TryGetProperty("inserted", out var i) && i.GetBoolean(),
                });
            }

            return list.Count > 0 ? list : null;
        }

        private static string GetLabel(Dictionary<string, JsonElement>? content)
        {
            if (content == null)
                return "";

            if (content.TryGetValue("numbering", out var numEl) &&
                numEl.ValueKind == JsonValueKind.Object &&
                numEl.TryGetProperty("label", out var labelEl))
            {
                return labelEl.GetString() ?? "";
            }

            return "";
        }

        private static string GetPlainText(BlockSideDto side, string label)
        {
            string text = "";
            var content = side.Node?.Content;

            if (content != null)
            {
                if (content.TryGetValue("text_display", out var textDisplayEl))
                    text = textDisplayEl.GetString() ?? "";
                else if (content.TryGetValue("text", out var textEl))
                    text = textEl.GetString() ?? "";
                else if (content.TryGetValue("image_url", out var imageUrlEl))
                    text = imageUrlEl.GetString() ?? "";
                else if (content.TryGetValue("url", out var urlEl))
                    text = urlEl.GetString() ?? "";
                else if (content.TryGetValue("src", out var srcEl))
                    text = srcEl.GetString() ?? "";
            }

            if (string.IsNullOrEmpty(text))
                text = side.PreviewText ?? "";

            return string.IsNullOrEmpty(label) ? text : $"{label} {text}";
        }

        private Run MakeImageRun(
            MainDocumentPart mainPart,
            string imagePath,
            long widthPx,
            long heightPx)
        {
            var ext = Path.GetExtension(imagePath).ToLowerInvariant();

            var imagePartType = ext switch
            {
                ".jpg" or ".jpeg" => ImagePartType.Jpeg,
                ".gif" => ImagePartType.Gif,
                ".bmp" => ImagePartType.Bmp,
                _ => ImagePartType.Png
            };

            var imagePart = mainPart.AddImagePart(imagePartType);

            using (var stream = File.OpenRead(imagePath))
            {
                imagePart.FeedData(stream);
            }

            var relationshipId = mainPart.GetIdOfPart(imagePart);

            const long emusPerPixel = 9525;
            long widthEmus = widthPx * emusPerPixel;
            long heightEmus = heightPx * emusPerPixel;

            var drawing = new Drawing(
                new DW.Inline(
                    new DW.Extent
                    {
                        Cx = widthEmus,
                        Cy = heightEmus
                    },
                    new DW.EffectExtent
                    {
                        LeftEdge = 0L,
                        TopEdge = 0L,
                        RightEdge = 0L,
                        BottomEdge = 0L
                    },
                    new DW.DocProperties
                    {
                        Id = _imageId++,
                        Name = Path.GetFileName(imagePath)
                    },
                    new DW.NonVisualGraphicFrameDrawingProperties(
                        new A.GraphicFrameLocks
                        {
                            NoChangeAspect = true
                        }
                    ),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties
                                    {
                                        Id = 0U,
                                        Name = Path.GetFileName(imagePath)
                                    },
                                    new PIC.NonVisualPictureDrawingProperties()
                                ),
                                new PIC.BlipFill(
                                    new A.Blip
                                    {
                                        Embed = relationshipId
                                    },
                                    new A.Stretch(
                                        new A.FillRectangle()
                                    )
                                ),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset
                                        {
                                            X = 0L,
                                            Y = 0L
                                        },
                                        new A.Extents
                                        {
                                            Cx = widthEmus,
                                            Cy = heightEmus
                                        }
                                    ),
                                    new A.PresetGeometry(
                                        new A.AdjustValueList()
                                    )
                                    {
                                        Preset = A.ShapeTypeValues.Rectangle
                                    }
                                )
                            )
                        )
                        {
                            Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                        }
                    )
                )
                {
                    DistanceFromTop = 0U,
                    DistanceFromBottom = 0U,
                    DistanceFromLeft = 0U,
                    DistanceFromRight = 0U
                }
            );

            return new Run(drawing);
        }

        private static Paragraph MakeTitle(string text)
        {
            var pPr = new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { After = "240" }
            );

            var rPr = new RunProperties(
                new RunFonts { Ascii = Font, HighAnsi = Font, ComplexScript = Font },
                new Bold(),
                new FontSize { Val = "28" }
            );

            return new Paragraph(
                pPr,
                new Run(rPr, new Text(text))
            );
        }

        private static void AddStyles(MainDocumentPart mainPart)
        {
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();

            stylesPart.Styles = new Styles(
                new DocDefaults(
                    new RunPropertiesDefault(
                        new RunPropertiesBaseStyle(
                            new RunFonts
                            {
                                Ascii = Font,
                                HighAnsi = Font,
                                ComplexScript = Font
                            },
                            new FontSize { Val = FontSize },
                            new FontSizeComplexScript { Val = FontSize }
                        )
                    )
                )
            );

            stylesPart.Styles.Save();
        }

        private static SectionProperties MakeSectionProperties()
        {
            return new SectionProperties(
                new PageSize
                {
                    Width = 16838,
                    Height = 11906,
                    Orient = PageOrientationValues.Landscape
                },
                new PageMargin
                {
                    Top = 720,
                    Right = 720,
                    Bottom = 720,
                    Left = 720,
                    Header = 0,
                    Footer = 0
                }
            );
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
}