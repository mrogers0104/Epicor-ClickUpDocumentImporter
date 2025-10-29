using ClickUpDocumentImporter.Helpers;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClickUpDocumentImporter.DocumentConverter
{
    public class WordToMarkdownFormatter
    {
        private readonly WordprocessingDocument _wordDoc;

        public WordToMarkdownFormatter(WordprocessingDocument wordDoc)
        {
            _wordDoc = wordDoc;
        }

        public void ProcessParagraph(Paragraph para, ClickUpDocumentBuilder builder)
        {
            var paragraphProperties = para.ParagraphProperties;
            var styleId = paragraphProperties?.ParagraphStyleId?.Val?.Value;
            var numProperties = paragraphProperties?.NumberingProperties;

            // Check if it's a heading
            if (styleId != null && styleId.StartsWith("Heading"))
            {
                int level = int.TryParse(styleId.Replace("Heading", ""), out int l) ? Math.Min(l, 6) : 1;
                string text = ExtractFormattedText(para);
                if (!string.IsNullOrWhiteSpace(text))
                {
                    builder.AddHeading(text, level);
                    ConsoleHelper.WriteInfo($"Added heading (level {level}): {text}");
                }
                return;
            }

            // Check if it's a list item
            if (numProperties != null)
            {
                string text = ExtractFormattedText(para);
                if (!string.IsNullOrWhiteSpace(text))
                {
                    var numId = numProperties.NumberingId?.Val?.Value ?? 0;
                    var ilvl = numProperties.NumberingLevelReference?.Val?.Value ?? 0;

                    // Get numbering format to determine bullet vs numbered
                    bool isOrdered = IsOrderedList(numId, ilvl);
                    string indent = new string(' ', (int)ilvl * 2);
                    string prefix = isOrdered ? "1. " : "- ";

                    builder.AddMarkdown($"{indent}{prefix}{text}");
                    ConsoleHelper.WriteInfo($"Added list item: {text}");
                }
                return;
            }

            // Check if it's a blockquote (typically styled as "Quote" or "IntenseQuote")
            if (styleId != null && (styleId.Contains("Quote") || styleId.Contains("Emphasis")))
            {
                string text = ExtractFormattedText(para);
                if (!string.IsNullOrWhiteSpace(text))
                {
                    builder.AddBlockquote(text);
                    ConsoleHelper.WriteInfo($"Added blockquote: {text}");
                }
                return;
            }

            // Check for code block (typically styled as "Code" or "HTMLCode")
            if (styleId != null && styleId.Contains("Code"))
            {
                string text = para.InnerText;
                if (!string.IsNullOrWhiteSpace(text))
                {
                    builder.AddCodeBlock(text);
                    ConsoleHelper.WriteInfo($"Added code block: {text}");
                }
                return;
            }

            // Check for horizontal rule (empty paragraph with border)
            if (HasBottomBorder(paragraphProperties) && string.IsNullOrWhiteSpace(para.InnerText))
            {
                builder.AddHorizontalRule();
                ConsoleHelper.WriteInfo("Added horizontal rule");
                return;
            }

            // Regular paragraph with inline formatting
            string formattedText = ExtractFormattedText(para);
            if (!string.IsNullOrWhiteSpace(formattedText))
            {
                builder.AddParagraph(formattedText);
                ConsoleHelper.WriteInfo($"Added paragraph: {formattedText.Substring(0, Math.Min(50, formattedText.Length))}...");
            }
        }

        private string ExtractFormattedText(Paragraph para)
        {
            var sb = new StringBuilder();

            foreach (var run in para.Elements<Run>())
            {
                string text = GetRunText(run);
                if (string.IsNullOrEmpty(text)) continue;

                var runProps = run.RunProperties;
                if (runProps == null)
                {
                    sb.Append(text);
                    continue;
                }

                // Track formatting flags
                //bool isBold = runProps.Bold != null && (runProps.Bold.Val != null || runProps.Bold.Val != false);
                //bool isItalic = runProps.Italic != null && (runProps.Italic.Val != null || runProps.Italic.Val != false);
                //bool isStrikethrough = runProps.Strike != null && (runProps.Strike.Val != null || runProps.Strike.Val != false);
                bool isBold = runProps.Bold?.Val ?? false;
                bool isItalic = runProps.Italic?.Val ?? false;
                bool isStrikethrough = runProps.Strike?.Val ?? false;
                bool isCode = IsCodeStyle(runProps);
                bool isSubscript = runProps.VerticalTextAlignment?.Val?.Value == VerticalPositionValues.Subscript;
                bool isSuperscript = runProps.VerticalTextAlignment?.Val?.Value == VerticalPositionValues.Superscript;
                bool isUnderline = runProps.Underline != null && runProps.Underline.Val != null;
                bool isHighlight = runProps.Highlight != null;

                // Check for hyperlink
                var hyperlink = run.Parent as Hyperlink;
                if (hyperlink != null)
                {
                    string url = GetHyperlinkUrl(hyperlink);
                    if (!string.IsNullOrEmpty(url))
                    {
                        sb.Append($"[{text}]({url})");
                        continue;
                    }
                }

                // Apply markdown formatting in proper order
                string formatted = text;

                // Code takes precedence
                if (isCode)
                {
                    formatted = $"`{formatted}`";
                }
                else
                {
                    // Bold and Italic combined = ***text***
                    if (isBold && isItalic)
                    {
                        formatted = $"***{formatted}***";
                    }
                    else if (isBold)
                    {
                        formatted = $"**{formatted}**";
                    }
                    else if (isItalic)
                    {
                        formatted = $"*{formatted}*";
                    }

                    // Strikethrough
                    if (isStrikethrough)
                    {
                        formatted = $"~~{formatted}~~";
                    }

                    // !!! ClickUp Markdown does not support HTML (October 15, 2025)
                    //// Subscript and Superscript (HTML fallback in markdown)
                    //if (isSubscript)
                    //{
                    //    formatted = $"<sub>{formatted}</sub>";
                    //}
                    //else if (isSuperscript)
                    //{
                    //    formatted = $"<sup>{formatted}</sup>";
                    //}

                    //// Underline (HTML fallback in markdown)
                    //if (isUnderline)
                    //{
                    //    formatted = $"<u>{formatted}</u>";
                    //}


                    // Highlight (HTML fallback)
                    if (isHighlight)
                    {
                        var color = GetHighlightColor(runProps.Highlight);
                        formatted = $"<mark style=\"background-color:{color}\">{formatted}</mark>";
                    }
                }

                sb.Append(formatted);
            }

            // Handle hyperlinks at paragraph level
            foreach (var hyperlink in para.Elements<Hyperlink>())
            {
                string text = hyperlink.InnerText;
                string url = GetHyperlinkUrl(hyperlink);
                if (!string.IsNullOrEmpty(url) && !string.IsNullOrEmpty(text))
                {
                    // This is already handled in run processing
                }
            }

            return sb.ToString();
        }

        private string GetRunText(Run run)
        {
            var sb = new StringBuilder();

            foreach (var element in run.Elements())
            {
                if (element is Text text)
                {
                    sb.Append(text.Text);
                }
                else if (element is TabChar)
                {
                    sb.Append("\t");
                }
                else if (element is Break br)
                {
                    // Line break
                    if (br.Type == null || br.Type == BreakValues.TextWrapping)
                    {
                        sb.Append("  \n"); // Markdown line break
                    }
                    else
                    {
                        sb.Append("\n");
                    }
                }
            }

            return sb.ToString();
        }

        private bool IsCodeStyle(RunProperties runProps)
        {
            // Check if font is monospace (common code fonts)
            var font = runProps.RunFonts?.Ascii?.Value;
            if (font != null)
            {
                var codeFonts = new[] { "Courier", "Consolas", "Monaco", "Menlo", "Monospace", "Courier New" };
                if (codeFonts.Any(cf => font.Contains(cf, StringComparison.OrdinalIgnoreCase)))
                    return true;
            }

            // Check for gray shading (common code style)
            var shading = runProps.Shading;
            if (shading?.Fill?.Value != null)
            {
                var fill = shading.Fill.Value;
                if (fill.Equals("E7E6E6", StringComparison.OrdinalIgnoreCase) ||
                    fill.Equals("F3F3F3", StringComparison.OrdinalIgnoreCase) ||
                    fill.Equals("D3D3D3", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private string GetHyperlinkUrl(Hyperlink hyperlink)
        {
            var id = hyperlink.Id?.Value;
            if (string.IsNullOrEmpty(id)) return null;

            try
            {
                var relationship = _wordDoc.MainDocumentPart.HyperlinkRelationships
                    .FirstOrDefault(r => r.Id == id);
                return relationship?.Uri?.ToString();
            }
            catch
            {
                return null;
            }
        }

        private bool IsOrderedList(int? numId, int ilvl)
        {
            if (numId == null) return false;

            try
            {
                var numberingPart = _wordDoc.MainDocumentPart.NumberingDefinitionsPart;
                if (numberingPart == null) return false;

                var numbering = numberingPart.Numbering;
                var numInstance = numbering.Elements<NumberingInstance>()
                    .FirstOrDefault(ni => ni.NumberID?.Value == numId);

                if (numInstance == null) return false;

                var abstractNumId = numInstance.AbstractNumId?.Val?.Value;
                if (abstractNumId == null) return false;

                var abstractNum = numbering.Elements<AbstractNum>()
                    .FirstOrDefault(an => an.AbstractNumberId?.Value == abstractNumId);

                if (abstractNum == null) return false;

                var level = abstractNum.Elements<Level>()
                    .FirstOrDefault(l => l.LevelIndex?.Value == ilvl);

                if (level?.NumberingFormat?.Val?.Value != null)
                {
                    var format = level.NumberingFormat.Val.Value;
                    return format != NumberFormatValues.Bullet;
                }
            }
            catch
            {
                // Default to bullet if we can't determine
            }

            return false;
        }

        private bool HasBottomBorder(ParagraphProperties props)
        {
            if (props?.ParagraphBorders?.BottomBorder != null)
            {
                var border = props.ParagraphBorders.BottomBorder;
                return border.Val != null && border.Val != BorderValues.None;
            }
            return false;
        }

        private string GetHighlightColor(Highlight highlight)
        {
            if (highlight?.Val?.Value == null) return "yellow";

            // Get the string value of the highlight color
            string colorValue = highlight.Val.Value.ToString().ToLower();

            // Map color names to CSS colors
            return colorValue switch
            {
                "yellow" => "yellow",
                "green" => "lightgreen",
                "cyan" => "cyan",
                "magenta" => "magenta",
                "blue" => "lightblue",
                "red" => "lightcoral",
                "darkblue" => "darkblue",
                "darkcyan" => "darkcyan",
                "darkgreen" => "darkgreen",
                "darkmagenta" => "darkmagenta",
                "darkred" => "darkred",
                "darkyellow" => "gold",
                "darkgray" => "darkgray",
                "lightgray" => "lightgray",
                "black" => "black",
                "white" => "white",
                _ => "yellow"
            };
        }

        public void ProcessTable(Table table, ClickUpDocumentBuilder builder)
        {
            var rows = table.Elements<TableRow>().ToList();
            if (!rows.Any()) return;

            var markdown = new StringBuilder();
            markdown.AppendLine(); // Blank line before table

            // Process each row
            for (int i = 0; i < rows.Count; i++)
            {
                var row = rows[i];
                var cells = row.Elements<TableCell>().ToList();

                markdown.Append("|");
                foreach (var cell in cells)
                {
                    string cellText = string.Join(" ", cell.Elements<Paragraph>()
                        .Select(p => ExtractFormattedText(p)));
                    markdown.Append($" {cellText.Trim()} |");
                }
                markdown.AppendLine();

                // Add header separator after first row
                if (i == 0)
                {
                    markdown.Append("|");
                    for (int j = 0; j < cells.Count; j++)
                    {
                        markdown.Append(" --- |");
                    }
                    markdown.AppendLine();
                }
            }

            markdown.AppendLine(); // Blank line after table
            builder.AddMarkdown(markdown.ToString());
            ConsoleHelper.WriteInfo($"Added table with {rows.Count} rows");
        }
    }

    // Extension methods for ClickUpDocumentBuilder
    public static class ClickUpDocumentBuilderExtensions
    {
        public static void AddMarkdown(this ClickUpDocumentBuilder builder, string markdown)
        {
            // Access the internal markdown content through reflection or add this as a public method
            // For now, use AddParagraph which appends to markdown
            typeof(ClickUpDocumentBuilder)
                .GetField("_markdownContent", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                ?.GetValue(builder)
                ?.GetType()
                .GetMethod("Append")
                ?.Invoke(
                    typeof(ClickUpDocumentBuilder)
                        .GetField("_markdownContent", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                        ?.GetValue(builder),
                    new object[] { markdown }
                );
        }

        public static void AddBlockquote(this ClickUpDocumentBuilder builder, string text)
        {
            var lines = text.Split('\n');
            foreach (var line in lines)
            {
                builder.AddMarkdown($"> {line}\n");
            }
            builder.AddMarkdown("\n");
        }

        public static void AddCodeBlock(this ClickUpDocumentBuilder builder, string code, string language = "")
        {
            builder.AddMarkdown($"```{language}\n{code}\n```\n\n");
        }

        public static void AddHorizontalRule(this ClickUpDocumentBuilder builder)
        {
            builder.AddMarkdown("\n---\n\n");
        }
    }
}
