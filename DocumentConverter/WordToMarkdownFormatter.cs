using ClickUpDocumentImporter.Helpers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;

using W = DocumentFormat.OpenXml.Wordprocessing;

namespace ClickUpDocumentImporter.DocumentConverter
{
    public class WordToMarkdownFormatter
    {
        private readonly WordprocessingDocument _wordDoc;
        private readonly MainDocumentPart _mainPart;
        private Dictionary<string, int> _listCounters; // Track list item numbers
        private int _previousListLevel = -1; // Track previous list level for reset logic

        public WordToMarkdownFormatter(WordprocessingDocument wordDoc)
        {
            _wordDoc = wordDoc;
            _mainPart = wordDoc?.MainDocumentPart;

            _listCounters = new Dictionary<string, int>();
            _previousListLevel = -1;
        }

        public void ProcessParagraph(Paragraph para, ClickUpDocumentBuilder builder)
        {
            var paragraphProperties = para.ParagraphProperties;
            var styleId = paragraphProperties?.ParagraphStyleId?.Val?.Value;
            var numProperties = paragraphProperties?.NumberingProperties;

            var numberingProperties = para.Descendants<W.NumberingProperties>().FirstOrDefault();

            if (numberingProperties != null)
            {
                var numIdElement = numberingProperties.Descendants<W.NumberingId>().FirstOrDefault();
                var levelIdElement = numberingProperties.Descendants<W.NumberingLevelReference>().FirstOrDefault();

                if (numIdElement != null && levelIdElement != null)
                {
                    int numId = numIdElement.Val.Value;
                    int ilvl = levelIdElement.Val.Value;
                    string listKey = $"{numId}-{ilvl}"; // e.g., "3-0"

                    // 1. Reset Lower Levels: If we hit a new list or a higher level (e.g., going from 1.1 to 2.),
                    // you must reset counters for all *lower* levels (e.g., level 1.1.X).
                    // This is advanced, but necessary for complex lists. For a simple fix, focus on incrementing.

                    // 2. Get/Increment the Counter:
                    if (!_listCounters.ContainsKey(listKey))
                    {
                        _listCounters[listKey] = 1; // Start at 1
                    }
                    else
                    {
                        _listCounters[listKey]++; // Increment
                    }

                    int currentCount = _listCounters[listKey];

                    // 3. Apply Markdown: Prepend the number to the text
                    string plainText = para.InnerText;

                    // NOTE: For ClickUp/Markdown, you usually just need "1. " or "  1. " (indented).
                    // We'll use the simplest form here:
                    string listItemText = $"{currentCount}. {plainText.Trim()}";

                    // You should adapt this to your builder's specific list method.
                    // For example:
                    builder.AddListItem(listItemText);
                    Console.WriteLine($"Added numbered paragraph: {listItemText.Substring(0, Math.Min(50, listItemText.Length))}...");

                    // Stop processing this paragraph as a standard paragraph now.
                    return;
                }
            }

            // Check if it's a heading
            if (styleId != null && styleId.StartsWith("Heading"))
            {
                int level = int.TryParse(styleId.Replace("Heading", ""), out int l) ? Math.Min(l, 6) : 1;
                string text = ExtractFormattedText(para);
                if (!string.IsNullOrWhiteSpace(text))
                {
                    builder.AddHeading(text, level);
                    Console.WriteLine($"Added heading (level {level}): {text}");
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
                    Console.WriteLine($"Added blockquote: {text}");
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
                    Console.WriteLine($"Added code block: {text}");
                }
                return;
            }

            // Check for horizontal rule (empty paragraph with border)
            if (HasBottomBorder(paragraphProperties) && string.IsNullOrWhiteSpace(para.InnerText))
            {
                builder.AddHorizontalRule();
                Console.WriteLine("Added horizontal rule");
                return;
            }

            // Regular paragraph with inline formatting
            string formattedText = ExtractFormattedText(para);
            if (!string.IsNullOrWhiteSpace(formattedText))
            {
                builder.AddParagraph(formattedText);
                Console.WriteLine($"Added paragraph: {formattedText.Substring(0, Math.Min(50, formattedText.Length))}...");
            }
        }

        private string ExtractFormattedText(Paragraph para)
        {
            var sb = new StringBuilder();
            string lastFormatting = ""; // Track the last formatting applied

            foreach (var run in para.Elements<Run>())
            {
                string text = GetRunText(run);
                if (string.IsNullOrEmpty(text)) continue;

                var runProps = run.RunProperties;
                if (runProps == null)
                {
                    sb.Append(text);
                    lastFormatting = "";
                    continue;
                }

                // Track formatting flags - Fixed to handle OnOffValue correctly
                bool isBold = IsBoldSet(runProps.Bold);
                bool isItalic = IsItalicSet(runProps.Italic);
                bool isStrikethrough = IsStrikeSet(runProps.Strike);
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
                        lastFormatting = "";
                        continue;
                    }
                }

                // Apply markdown formatting in proper order
                string formatted = text;
                string currentFormatting = "";

                // Extract leading and trailing whitespace
                string leadingSpace = "";
                string trailingSpace = "";
                string trimmedText = text;

                if (!string.IsNullOrEmpty(text))
                {
                    int leadingCount = text.Length - text.TrimStart().Length;
                    int trailingCount = text.Length - text.TrimEnd().Length;

                    if (leadingCount > 0)
                        leadingSpace = text.Substring(0, leadingCount);

                    if (trailingCount > 0)
                        trailingSpace = text.Substring(text.Length - trailingCount);

                    trimmedText = text.Trim();
                }

                // Code takes precedence
                if (isCode)
                {
                    currentFormatting = "code";
                    // Add space if previous run had same formatting and no trailing space
                    if (lastFormatting == currentFormatting && string.IsNullOrEmpty(leadingSpace))
                    {
                        sb.Append(" ");
                    }
                    formatted = leadingSpace + $"`{trimmedText}`" + trailingSpace;
                }
                else
                {
                    // Bold and Italic combined = ***text***
                    if (isBold && isItalic)
                    {
                        currentFormatting = "bolditalic";
                        if (lastFormatting == currentFormatting && string.IsNullOrEmpty(leadingSpace))
                        {
                            sb.Append(" ");
                        }
                        formatted = leadingSpace + $"***{trimmedText}***" + trailingSpace;
                    }
                    else if (isBold)
                    {
                        currentFormatting = "bold";
                        if (lastFormatting == currentFormatting && string.IsNullOrEmpty(leadingSpace))
                        {
                            sb.Append(" ");
                        }
                        formatted = leadingSpace + $"**{trimmedText}**" + trailingSpace;
                    }
                    else if (isItalic)
                    {
                        currentFormatting = "italic";
                        if (lastFormatting == currentFormatting && string.IsNullOrEmpty(leadingSpace))
                        {
                            sb.Append(" ");
                        }
                        formatted = leadingSpace + $"*{trimmedText}*" + trailingSpace;
                    }
                    else
                    {
                        formatted = text; // No formatting, keep original with spaces
                        currentFormatting = "";
                    }

                    // Strikethrough (can combine with bold/italic)
                    if (isStrikethrough)
                    {
                        string strikeFormatting = currentFormatting + "strike";
                        if (lastFormatting == strikeFormatting && string.IsNullOrEmpty(leadingSpace))
                        {
                            sb.Append(" ");
                        }

                        if (isBold || isItalic)
                        {
                            // Wrap the already formatted text
                            formatted = leadingSpace + $"~~{trimmedText}~~" + trailingSpace;
                        }
                        else
                        {
                            formatted = leadingSpace + $"~~{trimmedText}~~" + trailingSpace;
                        }
                        currentFormatting = strikeFormatting;
                    }
                }

                sb.Append(formatted);
                lastFormatting = currentFormatting;
            }

            //// Handle hyperlinks at paragraph level
            //foreach (var hyperlink in para.Elements<Hyperlink>())
            //{
            //    string text = hyperlink.InnerText;
            //    string url = GetHyperlinkUrl(hyperlink);
            //    if (!string.IsNullOrEmpty(url) && !string.IsNullOrEmpty(text))
            //    {
            //        // This is already handled in run processing
            //    }
            //}

            return sb.ToString();
        }

        // Helper method for Bold property
        /// <summary>
        /// Determines whether the specified <see cref="Bold"/> instance represents a bold setting.
        /// </summary>
        /// <remarks>If the <paramref name="bold"/> parameter is null, the method returns <see
        /// langword="false"/>. If the <see cref="Bold.Val"/> property is null, the method assumes a default value of
        /// <see langword="true"/>.</remarks>
        /// <param name="bold">The <see cref="Bold"/> instance to evaluate. Can be null.</param>
        /// <returns><see langword="true"/> if the <paramref name="bold"/> instance is not null and its value indicates bold;
        /// otherwise, <see langword="false"/>.</returns>
        private bool IsBoldSet(Bold bold)
        {
            if (bold == null)
                return false;
            return bold.Val?.Value ?? true;
        }

        // Helper method for Italic property
        private bool IsItalicSet(Italic italic)
        {
            if (italic == null)
                return false;
            return italic.Val?.Value ?? true;
        }

        // Helper method for Strike property
        private bool IsStrikeSet(Strike strike)
        {
            if (strike == null)
                return false;
            return strike.Val?.Value ?? true;
        }

        /// <summary>
        /// Extracts and concatenates the text content from the specified <see cref="Run"/> object,  including handling
        /// special elements such as tabs and line breaks.
        /// </summary>
        /// <remarks>This method processes the elements within the <see cref="Run"/> object and converts
        /// them  into a plain text representation. Tabs are represented as tab characters, and line breaks  are
        /// represented as either Markdown-style line breaks ("  \n") or standard line breaks ("\n"),  depending on the
        /// type of the break.</remarks>
        /// <param name="run">The <see cref="Run"/> object from which to extract text content.</param>
        /// <returns>A string containing the concatenated text content of the <paramref name="run"/>,  with tabs and line breaks
        /// appropriately represented.</returns>
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
            return false;

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

        private bool HasBottomBorder(ParagraphProperties props)
        {
            if (props?.ParagraphBorders?.BottomBorder != null)
            {
                var border = props.ParagraphBorders.BottomBorder;
                return border.Val != null && border.Val != BorderValues.None;
            }
            return false;
        }

        //private string GetHighlightColor(Highlight highlight)
        //{
        //    if (highlight?.Val?.Value == null) return "yellow";

        //    // Get the string value of the highlight color
        //    string colorValue = highlight.Val.Value.ToString().ToLower();

        //    // Map color names to CSS colors
        //    return colorValue switch
        //    {
        //        "yellow" => "yellow",
        //        "green" => "lightgreen",
        //        "cyan" => "cyan",
        //        "magenta" => "magenta",
        //        "blue" => "lightblue",
        //        "red" => "lightcoral",
        //        "darkblue" => "darkblue",
        //        "darkcyan" => "darkcyan",
        //        "darkgreen" => "darkgreen",
        //        "darkmagenta" => "darkmagenta",
        //        "darkred" => "darkred",
        //        "darkyellow" => "gold",
        //        "darkgray" => "darkgray",
        //        "lightgray" => "lightgray",
        //        "black" => "black",
        //        "white" => "white",
        //        _ => "yellow"
        //    };
        //}

        //// Helper method to safely check boolean properties
        //// In OpenXML, if a property exists but Val is null, it's considered true
        //private bool IsBoolPropertyOn(OnOffType property)
        //{
        //    if (property == null) return false;

        //    // If Val is null, the property is considered "on" (true)
        //    if (property.Val == null) return true;

        //    // Otherwise check the actual value
        //    return property.Val.Value != false;
        //}

        //// Reset list counters (call between document sections if needed)
        //public void ResetListCounters()
        //{
        //    _listCounters.Clear();
        //    _previousListLevel = -1;
        //}

        //// Reset counter for a specific list level (useful for nested lists)
        //private void ResetCounterForLevel(int numId, int level)
        //{
        //    string listKey = $"{numId}_{level}";
        //    if (_listCounters.ContainsKey(listKey))
        //    {
        //        _listCounters.Remove(listKey);
        //    }
        //}

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