using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Office2019.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Runtime.InteropServices.JavaScript.JSType;
using iTextSharp.text;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Data;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Pdf.Xobject;


namespace ClickUpDocumentImporter.DocumentConverter
{
    /// <summary>
    /// Custom extraction strategy to get formatted text
    /// </summary>
    /// <remarks>
    /// Some Key Points:
    /// 1. Added renderInfo.PreserveGraphicsState(): This is called immediately after receiving
    ///    the TextRenderInfo to preserve the graphics state before it's deleted
    /// 2. Created PreservedTextRenderInfo class: This stores the extracted information (text,
    ///    position, font details) so we don't need to access the graphics state later
    /// 3. Extract data immediately: All data from TextRenderInfo is extracted and stored in
    ///    PreservedTextRenderInfo during the EventOccurred method, before the graphics state is deleted
    /// 4. Updated all methods: Changed to work with PreservedTextRenderInfo instead of TextRenderInfo
    ///
    /// This approach extracts all necessary information from the TextRenderInfo object
    /// immediately and stores it, so you don't need to access the graphics state after the
    /// event has been processed.
    ///
    /// ** Key improvements:**
    ///
    /// 1. ** Better line grouping**: Groups all text chunks into lines first based on Y position, then combines chunks within each line
    /// 2. ** Smart spacing**: Adds spaces between chunks based on horizontal gaps
    /// 3. ** Paragraph detection**: Groups lines into paragraphs based on vertical spacing, indentation changes, and font changes
    /// 4. ** Bullet point preservation**: Detects bullet points(including "o") at the start of paragraphs and keeps the entire bullet item together
    /// 5. ** Single block per paragraph**: Each bullet point or paragraph becomes one contiguous text block instead of being fragmented
    ///
    /// </remarks>
    ///
    internal class FormattedTextExtractionStrategy : ITextExtractionStrategy
    {
        private List<FormattedTextBlock> blocks = null; // Changed to null to track if processed
        private List<PreservedTextRenderInfo> allChunks = new List<PreservedTextRenderInfo>();
        private const float LINE_SPACING_THRESHOLD = 12f;
        private const float HORIZONTAL_SPACING_THRESHOLD = 2f;
        private bool hasProcessed = false; // Track if we've already processed

        public void EventOccurred(IEventData data, EventType type)
        {
            if (type == EventType.RENDER_TEXT)
            {
                var renderInfo = (TextRenderInfo)data;

                // Preserve the graphics state before processing
                renderInfo.PreserveGraphicsState();

                var baseline = renderInfo.GetBaseline();

                // Store preserved information - ONLY store, don't create blocks yet
                var preservedInfo = new PreservedTextRenderInfo
                {
                    Text = renderInfo.GetText(),
                    X = baseline.GetStartPoint().Get(0),
                    Y = baseline.GetStartPoint().Get(1),
                    FontSize = renderInfo.GetFontSize(),
                    FontName = renderInfo.GetFont()?.GetFontProgram()?.GetFontNames()?.GetFontName() ?? "",
                    Width = baseline.GetLength()
                };

                allChunks.Add(preservedInfo);
            }
        }

        public ICollection<EventType> GetSupportedEvents()
        {
            return new List<EventType> { EventType.RENDER_TEXT };
        }

        public string GetResultantText()
        {
            if (!hasProcessed)
            {
                ProcessAllChunks();
            }
            return string.Join("\n", blocks.Select(b => b.Text));
        }

        public List<FormattedTextBlock> GetFormattedBlocks()
        {
            // Only process once
            if (!hasProcessed)
            {
                ProcessAllChunks();
            }
            return blocks;
        }

        private void ProcessAllChunks()
        {
            if (hasProcessed)
                return;

            hasProcessed = true;
            blocks = new List<FormattedTextBlock>(); // Initialize blocks here

            if (allChunks.Count == 0)
                return;

            // Sort chunks by Y (descending - top to bottom) then X (ascending - left to right)
            var sortedChunks = allChunks
                .OrderByDescending(c => c.Y)
                .ThenBy(c => c.X)
                .ToList();

            // Group chunks into lines
            var lines = GroupIntoLines(sortedChunks);

            // Process each line and combine into paragraphs
            var paragraphs = GroupLinesIntoParagraphs(lines);

            // Convert paragraphs to formatted blocks
            foreach (var paragraph in paragraphs)
            {
                blocks.Add(paragraph);
            }

            Console.WriteLine($"Processed {allChunks.Count} chunks into {blocks.Count} text blocks");
        }

        private List<LineInfo> GroupIntoLines(List<PreservedTextRenderInfo> chunks)
        {
            var lines = new List<LineInfo>();
            LineInfo currentLine = null;

            foreach (var chunk in chunks)
            {
                if (currentLine == null)
                {
                    currentLine = new LineInfo
                    {
                        Y = chunk.Y,
                        Chunks = new List<PreservedTextRenderInfo> { chunk }
                    };
                }
                else
                {
                    // Check if this chunk belongs to the current line (similar Y position)
                    if (Math.Abs(currentLine.Y - chunk.Y) < LINE_SPACING_THRESHOLD)
                    {
                        currentLine.Chunks.Add(chunk);
                    }
                    else
                    {
                        // New line detected
                        lines.Add(currentLine);
                        currentLine = new LineInfo
                        {
                            Y = chunk.Y,
                            Chunks = new List<PreservedTextRenderInfo> { chunk }
                        };
                    }
                }
            }

            if (currentLine != null && currentLine.Chunks.Count > 0)
            {
                lines.Add(currentLine);
            }

            return lines;
        }

        private List<FormattedTextBlock> GroupLinesIntoParagraphs(List<LineInfo> lines)
        {
            var paragraphs = new List<FormattedTextBlock>();

            if (lines.Count == 0)
                return paragraphs;

            StringBuilder currentParagraphText = new StringBuilder();
            float currentY = 0;
            float currentX = 0;
            float currentFontSize = 0;
            string currentFontName = "";
            bool currentIsBold = false;
            bool currentIsItalic = false;
            bool isBulletPoint = false;
            bool isNumberedList = false;

            for (int i = 0; i < lines.Count; i++)
            {
                var line = lines[i];

                // Sort chunks in line by X position (left to right)
                var sortedChunks = line.Chunks.OrderBy(c => c.X).ToList();

                // Combine chunks in the line with proper spacing
                string lineText = CombineChunksInLine(sortedChunks);

                // Get formatting from first chunk
                var firstChunk = sortedChunks[0];
                float fontSize = firstChunk.FontSize;
                string fontName = firstChunk.FontName;
                bool isBold = IsBoldFont(fontName);
                bool isItalic = IsItalicFont(fontName);

                // Check if this is the start of a new paragraph
                bool isNewParagraph = false;

                if (i == 0)
                {
                    isNewParagraph = true;
                }
                else
                {
                    var prevLine = lines[i - 1];
                    float lineSpacing = Math.Abs(prevLine.Y - line.Y);

                    // Detect paragraph breaks
                    if (lineSpacing > LINE_SPACING_THRESHOLD * 1.5 ||
                        Math.Abs(firstChunk.X - currentX) > 20 ||
                        Math.Abs(fontSize - currentFontSize) > 1)
                    {
                        isNewParagraph = true;
                    }
                }

                // Check for bullet points and numbered lists
                string trimmedLine = lineText.TrimStart();
                bool lineStartsWithBullet = trimmedLine.StartsWith("o ") ||
                                           trimmedLine.StartsWith("• ") ||
                                           trimmedLine.StartsWith("· ") ||
                                           trimmedLine.StartsWith("- ");
                bool lineStartsWithNumber = System.Text.RegularExpressions.Regex.IsMatch(
                    trimmedLine, @"^\d+[\.\)]\s");

                if (isNewParagraph && currentParagraphText.Length > 0)
                {
                    // Save the current paragraph
                    paragraphs.Add(CreateFormattedBlock(
                        currentParagraphText.ToString().Trim(),
                        currentX,
                        currentY,
                        currentFontSize,
                        currentFontName,
                        currentIsBold,
                        currentIsItalic,
                        isBulletPoint,
                        isNumberedList
                    ));

                    currentParagraphText.Clear();
                }

                // Start new paragraph or continue current one
                if (currentParagraphText.Length == 0)
                {
                    // Starting new paragraph
                    currentParagraphText.Append(lineText);
                    currentY = line.Y;
                    currentX = firstChunk.X;
                    currentFontSize = fontSize;
                    currentFontName = fontName;
                    currentIsBold = isBold;
                    currentIsItalic = isItalic;
                    isBulletPoint = lineStartsWithBullet;
                    isNumberedList = lineStartsWithNumber;
                }
                else
                {
                    // Continue current paragraph - add space if needed
                    if (!currentParagraphText.ToString().EndsWith(" ") &&
                        !lineText.StartsWith(" "))
                    {
                        currentParagraphText.Append(" ");
                    }
                    currentParagraphText.Append(lineText);
                }
            }

            // Add the last paragraph
            if (currentParagraphText.Length > 0)
            {
                paragraphs.Add(CreateFormattedBlock(
                    currentParagraphText.ToString().Trim(),
                    currentX,
                    currentY,
                    currentFontSize,
                    currentFontName,
                    currentIsBold,
                    currentIsItalic,
                    isBulletPoint,
                    isNumberedList
                ));
            }

            return paragraphs;
        }

        private string CombineChunksInLine(List<PreservedTextRenderInfo> chunks)
        {
            if (chunks.Count == 0)
                return "";

            StringBuilder lineText = new StringBuilder();

            for (int i = 0; i < chunks.Count; i++)
            {
                var chunk = chunks[i];

                // Add the chunk text
                lineText.Append(chunk.Text);

                // Check if we need to add a space before the next chunk
                if (i < chunks.Count - 1)
                {
                    var nextChunk = chunks[i + 1];
                    float gap = nextChunk.X - (chunk.X + chunk.Width);

                    // If there's a gap between chunks, add a space
                    if (gap > HORIZONTAL_SPACING_THRESHOLD)
                    {
                        lineText.Append(" ");
                    }
                }
            }

            return lineText.ToString();
        }

        private FormattedTextBlock CreateFormattedBlock(
            string text,
            float x,
            float y,
            float fontSize,
            string fontName,
            bool isBold,
            bool isItalic,
            bool isBulletPoint,
            bool isNumberedList)
        {
            bool isHeading = fontSize > 12 && isBold && !isBulletPoint && !isNumberedList;

            return new FormattedTextBlock
            {
                Text = text,
                X = x,
                Y = y,
                FontSize = fontSize,
                FontName = fontName,
                IsBold = isBold,
                IsItalic = isItalic,
                IsHeading = isHeading,
                IsBulletPoint = isBulletPoint,
                IsNumberedList = isNumberedList,
                IsUnderlined = false,
                IsStrikethrough = false,
                IsCode = IsMonospaceFont(fontName),
                IsCodeBlock = false,
                IsBlockQuote = text.TrimStart().StartsWith(">"),
                IsLink = false
            };
        }

        private bool IsBoldFont(string fontName)
        {
            if (string.IsNullOrEmpty(fontName))
                return false;

            fontName = fontName.ToLower();
            return fontName.Contains("bold") ||
                   fontName.Contains("heavy") ||
                   fontName.Contains("black");
        }

        private bool IsItalicFont(string fontName)
        {
            if (string.IsNullOrEmpty(fontName))
                return false;

            fontName = fontName.ToLower();
            return fontName.Contains("italic") ||
                   fontName.Contains("oblique");
        }

        private bool IsMonospaceFont(string fontName)
        {
            if (string.IsNullOrEmpty(fontName))
                return false;

            fontName = fontName.ToLower();
            return fontName.Contains("courier") ||
                   fontName.Contains("mono") ||
                   fontName.Contains("console") ||
                   fontName.Contains("code");
        }

        // Helper classes
        private class PreservedTextRenderInfo
        {
            public string Text { get; set; }
            public float X { get; set; }
            public float Y { get; set; }
            public float FontSize { get; set; }
            public string FontName { get; set; }
            public float Width { get; set; }
        }

        private class LineInfo
        {
            public float Y { get; set; }
            public List<PreservedTextRenderInfo> Chunks { get; set; }
        }
    }

    //internal class FormattedTextExtractionStrategy : ITextExtractionStrategy
    //{
    //    private List<FormattedTextBlock> blocks = new List<FormattedTextBlock>();
    //    private List<PreservedTextRenderInfo> allChunks = new List<PreservedTextRenderInfo>();
    //    private const float LINE_SPACING_THRESHOLD = 12f; // Increased for better line detection
    //    private const float HORIZONTAL_SPACING_THRESHOLD = 2f; // Space between words

    //    public void EventOccurred(IEventData data, EventType type)
    //    {
    //        if (type == EventType.RENDER_TEXT)
    //        {
    //            var renderInfo = (TextRenderInfo)data;

    //            // Preserve the graphics state before processing
    //            renderInfo.PreserveGraphicsState();

    //            var baseline = renderInfo.GetBaseline();

    //            // Store preserved information
    //            var preservedInfo = new PreservedTextRenderInfo
    //            {
    //                Text = renderInfo.GetText(),
    //                X = baseline.GetStartPoint().Get(0),
    //                Y = baseline.GetStartPoint().Get(1),
    //                FontSize = renderInfo.GetFontSize(),
    //                FontName = renderInfo.GetFont()?.GetFontProgram()?.GetFontNames()?.GetFontName() ?? "",
    //                Width = baseline.GetLength()
    //            };

    //            allChunks.Add(preservedInfo);
    //        }
    //    }

    //    public ICollection<EventType> GetSupportedEvents()
    //    {
    //        return new List<EventType> { EventType.RENDER_TEXT };
    //    }

    //    public string GetResultantText()
    //    {
    //        ProcessAllChunks();
    //        return string.Join("\n", blocks.Select(b => b.Text));
    //    }

    //    public List<FormattedTextBlock> GetFormattedBlocks()
    //    {
    //        ProcessAllChunks();
    //        return blocks;
    //    }

    //    private void ProcessAllChunks()
    //    {
    //        if (allChunks.Count == 0)
    //            return;

    //        // Sort chunks by Y (descending - top to bottom) then X (ascending - left to right)
    //        var sortedChunks = allChunks
    //            .OrderByDescending(c => c.Y)
    //            .ThenBy(c => c.X)
    //            .ToList();

    //        // Group chunks into lines
    //        var lines = GroupIntoLines(sortedChunks);

    //        // Process each line and combine into paragraphs
    //        var paragraphs = GroupLinesIntoParagraphs(lines);

    //        // Convert paragraphs to formatted blocks
    //        foreach (var paragraph in paragraphs)
    //        {
    //            blocks.Add(paragraph);
    //        }
    //    }

    //    private List<LineInfo> GroupIntoLines(List<PreservedTextRenderInfo> chunks)
    //    {
    //        var lines = new List<LineInfo>();
    //        LineInfo currentLine = null;

    //        foreach (var chunk in chunks)
    //        {
    //            if (currentLine == null)
    //            {
    //                currentLine = new LineInfo
    //                {
    //                    Y = chunk.Y,
    //                    Chunks = new List<PreservedTextRenderInfo> { chunk }
    //                };
    //            }
    //            else
    //            {
    //                // Check if this chunk belongs to the current line (similar Y position)
    //                if (Math.Abs(currentLine.Y - chunk.Y) < LINE_SPACING_THRESHOLD)
    //                {
    //                    currentLine.Chunks.Add(chunk);
    //                }
    //                else
    //                {
    //                    // New line detected
    //                    lines.Add(currentLine);
    //                    currentLine = new LineInfo
    //                    {
    //                        Y = chunk.Y,
    //                        Chunks = new List<PreservedTextRenderInfo> { chunk }
    //                    };
    //                }
    //            }
    //        }

    //        if (currentLine != null && currentLine.Chunks.Count > 0)
    //        {
    //            lines.Add(currentLine);
    //        }

    //        return lines;
    //    }

    //    private List<FormattedTextBlock> GroupLinesIntoParagraphs(List<LineInfo> lines)
    //    {
    //        var paragraphs = new List<FormattedTextBlock>();

    //        if (lines.Count == 0)
    //            return paragraphs;

    //        StringBuilder currentParagraphText = new StringBuilder();
    //        float currentY = 0;
    //        float currentX = 0;
    //        float currentFontSize = 0;
    //        string currentFontName = "";
    //        bool currentIsBold = false;
    //        bool currentIsItalic = false;
    //        bool isBulletPoint = false;
    //        bool isNumberedList = false;

    //        for (int i = 0; i < lines.Count; i++)
    //        {
    //            var line = lines[i];

    //            // Sort chunks in line by X position (left to right)
    //            var sortedChunks = line.Chunks.OrderBy(c => c.X).ToList();

    //            // Combine chunks in the line with proper spacing
    //            string lineText = CombineChunksInLine(sortedChunks);

    //            // Get formatting from first chunk
    //            var firstChunk = sortedChunks[0];
    //            float fontSize = firstChunk.FontSize;
    //            string fontName = firstChunk.FontName;
    //            bool isBold = IsBoldFont(fontName);
    //            bool isItalic = IsItalicFont(fontName);

    //            // Check if this is the start of a new paragraph
    //            bool isNewParagraph = false;

    //            if (i == 0)
    //            {
    //                isNewParagraph = true;
    //            }
    //            else
    //            {
    //                var prevLine = lines[i - 1];
    //                float lineSpacing = Math.Abs(prevLine.Y - line.Y);

    //                // Detect paragraph breaks based on:
    //                // 1. Large vertical spacing (more than 1.5x normal line height)
    //                // 2. Change in indentation (bullet points, new sections)
    //                // 3. Change in font size (headings)
    //                if (lineSpacing > LINE_SPACING_THRESHOLD * 1.5 ||
    //                    Math.Abs(firstChunk.X - currentX) > 20 ||
    //                    Math.Abs(fontSize - currentFontSize) > 1)
    //                {
    //                    isNewParagraph = true;
    //                }
    //            }

    //            // Check for bullet points and numbered lists
    //            string trimmedLine = lineText.TrimStart();
    //            bool lineStartsWithBullet = trimmedLine.StartsWith("o ") ||
    //                                       trimmedLine.StartsWith("• ") ||
    //                                       trimmedLine.StartsWith("· ") ||
    //                                       trimmedLine.StartsWith("- ");
    //            bool lineStartsWithNumber = System.Text.RegularExpressions.Regex.IsMatch(
    //                trimmedLine, @"^\d+[\.\)]\s");

    //            if (isNewParagraph && currentParagraphText.Length > 0)
    //            {
    //                // Save the current paragraph
    //                paragraphs.Add(CreateFormattedBlock(
    //                    currentParagraphText.ToString().Trim(),
    //                    currentX,
    //                    currentY,
    //                    currentFontSize,
    //                    currentFontName,
    //                    currentIsBold,
    //                    currentIsItalic,
    //                    isBulletPoint,
    //                    isNumberedList
    //                ));

    //                currentParagraphText.Clear();
    //            }

    //            // Start new paragraph or continue current one
    //            if (currentParagraphText.Length == 0)
    //            {
    //                // Starting new paragraph
    //                currentParagraphText.Append(lineText);
    //                currentY = line.Y;
    //                currentX = firstChunk.X;
    //                currentFontSize = fontSize;
    //                currentFontName = fontName;
    //                currentIsBold = isBold;
    //                currentIsItalic = isItalic;
    //                isBulletPoint = lineStartsWithBullet;
    //                isNumberedList = lineStartsWithNumber;
    //            }
    //            else
    //            {
    //                // Continue current paragraph - add space if needed
    //                if (!currentParagraphText.ToString().EndsWith(" ") &&
    //                    !lineText.StartsWith(" "))
    //                {
    //                    currentParagraphText.Append(" ");
    //                }
    //                currentParagraphText.Append(lineText);
    //            }
    //        }

    //        // Add the last paragraph
    //        if (currentParagraphText.Length > 0)
    //        {
    //            paragraphs.Add(CreateFormattedBlock(
    //                currentParagraphText.ToString().Trim(),
    //                currentX,
    //                currentY,
    //                currentFontSize,
    //                currentFontName,
    //                currentIsBold,
    //                currentIsItalic,
    //                isBulletPoint,
    //                isNumberedList
    //            ));
    //        }

    //        return paragraphs;
    //    }

    //    private string CombineChunksInLine(List<PreservedTextRenderInfo> chunks)
    //    {
    //        if (chunks.Count == 0)
    //            return "";

    //        StringBuilder lineText = new StringBuilder();

    //        for (int i = 0; i < chunks.Count; i++)
    //        {
    //            var chunk = chunks[i];

    //            // Add the chunk text
    //            lineText.Append(chunk.Text);

    //            // Check if we need to add a space before the next chunk
    //            if (i < chunks.Count - 1)
    //            {
    //                var nextChunk = chunks[i + 1];
    //                float gap = nextChunk.X - (chunk.X + chunk.Width);

    //                // If there's a gap between chunks, add a space
    //                if (gap > HORIZONTAL_SPACING_THRESHOLD)
    //                {
    //                    lineText.Append(" ");
    //                }
    //            }
    //        }

    //        return lineText.ToString();
    //    }

    //    private FormattedTextBlock CreateFormattedBlock(
    //        string text,
    //        float x,
    //        float y,
    //        float fontSize,
    //        string fontName,
    //        bool isBold,
    //        bool isItalic,
    //        bool isBulletPoint,
    //        bool isNumberedList)
    //    {
    //        bool isHeading = fontSize > 12 && isBold && !isBulletPoint && !isNumberedList;

    //        return new FormattedTextBlock
    //        {
    //            Text = text,
    //            X = x,
    //            Y = y,
    //            FontSize = fontSize,
    //            FontName = fontName,
    //            IsBold = isBold,
    //            IsItalic = isItalic,
    //            IsHeading = isHeading,
    //            IsBulletPoint = isBulletPoint,
    //            IsNumberedList = isNumberedList,
    //            IsUnderlined = false,
    //            IsStrikethrough = false,
    //            IsCode = IsMonospaceFont(fontName),
    //            IsCodeBlock = false,
    //            IsBlockQuote = text.TrimStart().StartsWith(">"),
    //            IsLink = false
    //        };
    //    }

    //    private bool IsBoldFont(string fontName)
    //    {
    //        if (string.IsNullOrEmpty(fontName))
    //            return false;

    //        fontName = fontName.ToLower();
    //        return fontName.Contains("bold") ||
    //               fontName.Contains("heavy") ||
    //               fontName.Contains("black");
    //    }

    //    private bool IsItalicFont(string fontName)
    //    {
    //        if (string.IsNullOrEmpty(fontName))
    //            return false;

    //        fontName = fontName.ToLower();
    //        return fontName.Contains("italic") ||
    //               fontName.Contains("oblique");
    //    }

    //    private bool IsMonospaceFont(string fontName)
    //    {
    //        if (string.IsNullOrEmpty(fontName))
    //            return false;

    //        fontName = fontName.ToLower();
    //        return fontName.Contains("courier") ||
    //               fontName.Contains("mono") ||
    //               fontName.Contains("console") ||
    //               fontName.Contains("code");
    //    }

    //    // Helper classes
    //    private class PreservedTextRenderInfo
    //    {
    //        public string Text { get; set; }
    //        public float X { get; set; }
    //        public float Y { get; set; }
    //        public float FontSize { get; set; }
    //        public string FontName { get; set; }
    //        public float Width { get; set; }
    //    }

    //    private class LineInfo
    //    {
    //        public float Y { get; set; }
    //        public List<PreservedTextRenderInfo> Chunks { get; set; }
    //    }
    //}
}
