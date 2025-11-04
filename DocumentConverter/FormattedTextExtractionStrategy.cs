using iText.Kernel.Font;

//using iTextSharp.text;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Data;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using System.Text;

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
        private List<FormattedTextBlock> blocks = null;
        private List<PreservedTextRenderInfo> allChunks = new List<PreservedTextRenderInfo>();
        private const float LINE_SPACING_THRESHOLD = 12f;
        private const float HORIZONTAL_SPACING_THRESHOLD = 2f;
        private bool hasProcessed = false;

        /// <summary>
        /// Track events
        /// </summary>
        /// <remarks>
        /// <b>NOTE</b>
        /// Even though it appears this method is not being used,
        /// it must remain for ITextExtractionStragety to work properly.
        /// </remarks>
        /// <param name="data"></param>
        /// <param name="type"></param>
        public void EventOccurred(IEventData data, EventType type)
        {
            if (type == EventType.RENDER_TEXT)
            {
                var renderInfo = (TextRenderInfo)data;

                // Preserve the graphics state before processing
                renderInfo.PreserveGraphicsState();

                var baseline = renderInfo.GetBaseline();
                var font = renderInfo.GetFont();
                string fontName = font?.GetFontProgram()?.GetFontNames()?.GetFontName() ?? "";

                // Get text rendering mode for strikethrough detection
                int textRenderingMode = renderInfo.GetGraphicsState().GetTextRenderingMode();

                // Get rise (for superscript/subscript)
                float rise = renderInfo.GetRise();

                // Detect formatting from font properties and rendering mode
                bool isBold = IsBoldFont(fontName) || IsBoldFromFontWeight(font);
                bool isItalic = IsItalicFont(fontName) || IsItalicFromFontStyle(font);
                bool isUnderlined = false; // Will be detected from text decoration
                bool isStrikethrough = textRenderingMode == 3; // Mode 3 is typically strikethrough
                bool isMonospace = IsMonospaceFont(fontName);

                // Store preserved information
                var preservedInfo = new PreservedTextRenderInfo
                {
                    Text = renderInfo.GetText(),
                    X = baseline.GetStartPoint().Get(0),
                    Y = baseline.GetStartPoint().Get(1),
                    FontSize = renderInfo.GetFontSize(),
                    FontName = fontName,
                    Width = baseline.GetLength(),
                    IsBold = isBold,
                    IsItalic = isItalic,
                    IsUnderlined = isUnderlined,
                    IsStrikethrough = isStrikethrough,
                    IsMonospace = false, // isMonospace,
                    Rise = rise
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
            blocks = new List<FormattedTextBlock>();

            if (allChunks.Count == 0)
                return;

            var sortedChunks = allChunks
                .OrderByDescending(c => c.Y)
                .ThenBy(c => c.X)
                .ToList();

            var lines = GroupIntoLines(sortedChunks);
            var paragraphs = GroupLinesIntoParagraphs(lines);

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
                    if (Math.Abs(currentLine.Y - chunk.Y) < LINE_SPACING_THRESHOLD)
                    {
                        currentLine.Chunks.Add(chunk);
                    }
                    else
                    {
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

        // --------------------------------------------------------------------------------
        private List<FormattedTextBlock> GroupLinesIntoParagraphs(List<LineInfo> lines)
        {
            var paragraphs = new List<FormattedTextBlock>();

            if (lines.Count == 0)
                return paragraphs;

            // Track current paragraph being built
            List<FormattedSegment> currentParagraphSegments = new List<FormattedSegment>();
            float currentParagraphY = 0;
            float currentParagraphX = 0;
            float currentParagraphFontSize = 0;

            for (int i = 0; i < lines.Count; i++)
            {
                var line = lines[i];
                var sortedChunks = line.Chunks.OrderBy(c => c.X).ToList();

                // Group chunks by similar formatting within this line
                var formattedSegments = GroupChunksByFormatting(sortedChunks);

                // Determine if we should start a new paragraph
                bool isNewParagraph = ShouldStartNewParagraph(lines, i);

                if (isNewParagraph && currentParagraphSegments.Count > 0)
                {
                    // Create a block from the accumulated segments
                    var block = CreateBlockFromSegments(
                        currentParagraphSegments,
                        currentParagraphX,
                        currentParagraphY,
                        currentParagraphFontSize
                    );

                    if (block != null)
                    {
                        paragraphs.Add(block);
                    }

                    currentParagraphSegments.Clear();
                }

                // Add segments from this line to the current paragraph
                if (currentParagraphSegments.Count == 0)
                {
                    // Starting new paragraph
                    currentParagraphY = line.Y;
                    currentParagraphX = sortedChunks[0].X;
                    currentParagraphFontSize = sortedChunks[0].FontSize;
                }

                currentParagraphSegments.AddRange(formattedSegments);
            }

            // Don't forget the last paragraph
            if (currentParagraphSegments.Count > 0)
            {
                var block = CreateBlockFromSegments(
                    currentParagraphSegments,
                    currentParagraphX,
                    currentParagraphY,
                    currentParagraphFontSize
                );

                if (block != null)
                {
                    paragraphs.Add(block);
                }
            }

            return paragraphs;
        }

        private FormattedTextBlock CreateBlockFromSegments(
            List<FormattedSegment> segments,
            float x,
            float y,
            float fontSize)
        {
            if (segments.Count == 0)
                return null;

            // Combine all segments into a single text with markdown formatting
            StringBuilder combinedText = new StringBuilder();

            foreach (var segment in segments)
            {
                string segmentText = CombineChunksInSegment(segment.Chunks);

                if (string.IsNullOrWhiteSpace(segmentText))
                    continue;

                // Apply inline formatting to this segment
                string formattedSegment = ApplyInlineFormattingToSegment(
                    segmentText,
                    segment.IsBold,
                    segment.IsItalic,
                    segment.IsUnderlined,
                    segment.IsStrikethrough,
                    false //segment.IsMonospace
                );

                combinedText.Append(formattedSegment);
            }

            string fullText = combinedText.ToString().Trim();

            if (string.IsNullOrWhiteSpace(fullText))
                return null;

            // Get the dominant formatting from first segment (for block-level decisions)
            var firstSegment = segments[0];
            var firstChunk = firstSegment.Chunks[0];

            // Detect block-level properties
            bool isBulletPoint = fullText.TrimStart().StartsWith('o') ||
                                 fullText.TrimStart().StartsWith('•') ||
                                 fullText.TrimStart().StartsWith('·') ||
                                 fullText.TrimStart().StartsWith('-');
            if (isBulletPoint)
            {
                // The string starts with a bullet. Only remove the first character.
                fullText = fullText[1..];
            }

            bool isNumberedList = System.Text.RegularExpressions.Regex.IsMatch(
                fullText.TrimStart(), @"^\d+[\.\)]\s");

            // Collect all chunks for code detection
            // !!! This does not work so ignore code blocks.
            var allChunks = segments.SelectMany(s => s.Chunks).ToList();
            bool isCodeBlock = IsCodeBlock(allChunks, fullText);

            bool isHeading = fontSize > 12 && firstSegment.IsBold &&
                            !isBulletPoint && !isNumberedList && !isCodeBlock;

            // For block properties, use the dominant (first segment) formatting
            return new FormattedTextBlock
            {
                Text = fullText,
                X = x,
                Y = y,
                FontSize = fontSize,
                FontName = firstChunk.FontName,
                IsBold = firstSegment.IsBold, // Dominant formatting
                IsItalic = firstSegment.IsItalic,
                IsUnderlined = firstSegment.IsUnderlined,
                IsStrikethrough = firstSegment.IsStrikethrough,
                IsHeading = isHeading,
                IsBulletPoint = isBulletPoint,
                IsNumberedList = isNumberedList,
                IsCode = false, // Inline code is now embedded in the text
                IsCodeBlock = false, // isCodeBlock,
                IsBlockQuote = fullText.TrimStart().StartsWith(">"),
                IsLink = false
            };
        }

        private string ApplyInlineFormattingToSegment(
            string text,
            bool isBold,
            bool isItalic,
            bool isUnderlined,
            bool isStrikethrough,
            bool isMonospace)
        {
            // Don't add formatting if text is already empty or whitespace
            if (string.IsNullOrWhiteSpace(text))
                return text;

            // *** If there is not formatting for this segment, return the original text
            // *** to maintain spacing.
            if (!isBold && !isItalic && !isUnderlined && !isStrikethrough && !isMonospace)
            {
                return text;
            }

            string formatted = text.Trim();

            //// Apply inline code first (if monospace and short)
            //if (isMonospace && text.Length < 100)
            //{
            //    formatted = $"`{formatted}`";
            //}

            // Apply bold and italic
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

            // Apply underline
            if (isUnderlined)
            {
                formatted = $"<u>{formatted}</u>";
            }

            // Apply strikethrough
            if (isStrikethrough)
            {
                formatted = $"~~{formatted}~~";
            }

            return formatted;
        }

        private bool ShouldStartNewParagraph(List<LineInfo> lines, int currentIndex)
        {
            if (currentIndex == 0)
                return true;

            var currentLine = lines[currentIndex];
            var prevLine = lines[currentIndex - 1];

            float lineSpacing = Math.Abs(prevLine.Y - currentLine.Y);

            // Get first chunk positions to check for indentation changes
            var currentFirstChunk = currentLine.Chunks.OrderBy(c => c.X).First();
            var prevFirstChunk = prevLine.Chunks.OrderBy(c => c.X).First();

            // Significant indentation change
            if (Math.Abs(currentFirstChunk.X - prevFirstChunk.X) > 20)
                return true;

            // Font size change (heading detection)
            if (Math.Abs(currentFirstChunk.FontSize - prevFirstChunk.FontSize) > 1)
                return true;

            // Large vertical spacing
            return lineSpacing > LINE_SPACING_THRESHOLD * 1.5;
        }

        // --------------------------------------------------------------------------------

        private bool IsCodeBlock(List<PreservedTextRenderInfo> chunks, string text)
        {
            if (chunks.Count == 0)
                return false;

            // Check multiple heuristics for code detection
            int codeIndicators = 0;

            //// 1. Monospace font (original check)
            //if (chunks.Any(c => c.IsMonospace))
            //{
            //    codeIndicators += 1; // Strong indicator
            //}

            // 2. Check for common code patterns
            if (HasCodePatterns(text))
            {
                codeIndicators += 2;
            }

            // 3. Check for syntax characters common in code
            if (HasCodeSyntaxCharacters(text))
            {
                codeIndicators += 1;
            }

            // 4. Check indentation (code often has consistent indentation)
            if (chunks[0].X > 50) // Indented more than typical paragraph
            {
                codeIndicators += 1;
            }

            // 5. Check for all caps (not typical in code)
            if (text.All(c => !char.IsLetter(c) || char.IsUpper(c)))
            {
                codeIndicators -= 1;
            }

            // 6. Background color (code blocks sometimes have grey background)
            // This would require extracting background color from PDF (advanced)

            // Return true if we have strong evidence (3+ indicators)
            return codeIndicators >= 3;
        }

        private bool HasCodePatterns(string text)
        {
            // Common code patterns
            var codePatterns = new[]
            {
        @"^(public|private|protected|internal|static|void|class|interface|struct|enum)\s",
        @"^(function|const|let|var|def|import|export|return)\s",
        @"^\s*(if|else|for|while|switch|case|try|catch|finally)\s*[\(\{]",
        @"[{}()\[\];].*[{}()\[\];]", // Multiple brackets/braces
        @"^\s*//", // Comment
        @"^\s*/\*", // Block comment
        @"^\s*#.*", // Python/shell comment or preprocessor
        @"=>\s*{", // Arrow function
        @"\w+\s*\(.*\)\s*{", // Function definition
        @"^\s*<\w+.*>.*</\w+>", // XML/HTML tag
        @"SELECT|FROM|WHERE|INSERT|UPDATE|DELETE", // SQL (case insensitive check separately)
    };

            foreach (var pattern in codePatterns)
            {
                if (System.Text.RegularExpressions.Regex.IsMatch(text, pattern,
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private bool HasCodeSyntaxCharacters(string text)
        {
            // Count programming-specific characters
            int syntaxCharCount = 0;
            var syntaxChars = new[] { '{', '}', '[', ']', ';', '(', ')', '<', '>', '=' };

            foreach (char c in text)
            {
                if (syntaxChars.Contains(c))
                    syntaxCharCount++;
            }

            // If more than 10% of characters are syntax characters, likely code
            return text.Length > 0 && (syntaxCharCount / (float)text.Length) > 0.10;
        }

        private List<FormattedSegment> GroupChunksByFormatting(List<PreservedTextRenderInfo> chunks)
        {
            var segments = new List<FormattedSegment>();
            var currentSegment = new FormattedSegment
            {
                Chunks = new List<PreservedTextRenderInfo>()
            };

            foreach (var chunk in chunks)
            {
                if (currentSegment.Chunks.Count == 0)
                {
                    // Start new segment
                    currentSegment.Chunks.Add(chunk);
                    currentSegment.IsBold = chunk.IsBold;
                    currentSegment.IsItalic = chunk.IsItalic;
                    currentSegment.IsUnderlined = chunk.IsUnderlined;
                    currentSegment.IsStrikethrough = chunk.IsStrikethrough;
                    currentSegment.IsMonospace = false; // chunk.IsMonospace;
                }
                else
                {
                    // Check if formatting matches current segment
                    if (chunk.IsBold == currentSegment.IsBold &&
                        chunk.IsItalic == currentSegment.IsItalic &&
                        chunk.IsUnderlined == currentSegment.IsUnderlined &&
                        chunk.IsStrikethrough == currentSegment.IsStrikethrough &&
                        chunk.IsMonospace == false && // currentSegment.IsMonospace &&
                        Math.Abs(chunk.FontSize - currentSegment.Chunks[0].FontSize) < 0.5f)
                    {
                        currentSegment.Chunks.Add(chunk);
                    }
                    else
                    {
                        // Save current segment and start new one
                        segments.Add(currentSegment);
                        currentSegment = new FormattedSegment
                        {
                            Chunks = new List<PreservedTextRenderInfo> { chunk },
                            IsBold = chunk.IsBold,
                            IsItalic = chunk.IsItalic,
                            IsUnderlined = chunk.IsUnderlined,
                            IsStrikethrough = chunk.IsStrikethrough,
                            IsMonospace = false // chunk.IsMonospace
                        };
                    }
                }
            }

            if (currentSegment.Chunks.Count > 0)
            {
                segments.Add(currentSegment);
            }

            return segments;
        }

        //private bool ShouldStartNewParagraph(List<LineInfo> lines, int currentIndex)
        //{
        //    if (currentIndex == 0)
        //        return true;

        //    var currentLine = lines[currentIndex];
        //    var prevLine = lines[currentIndex - 1];

        //    float lineSpacing = Math.Abs(prevLine.Y - currentLine.Y);

        //    return lineSpacing > LINE_SPACING_THRESHOLD * 1.5;
        //}

        private string CombineChunksInSegment(List<PreservedTextRenderInfo> chunks)
        {
            if (chunks.Count == 0)
                return "";

            StringBuilder text = new StringBuilder();

            for (int i = 0; i < chunks.Count; i++)
            {
                var chunk = chunks[i];
                text.Append(chunk.Text);

                if (i < chunks.Count - 1)
                {
                    var nextChunk = chunks[i + 1];
                    float gap = nextChunk.X - (chunk.X + chunk.Width);

                    if (gap > HORIZONTAL_SPACING_THRESHOLD)
                    {
                        text.Append(" ");
                    }
                }
            }

            return text.ToString();
        }

        private bool IsBoldFont(string fontName)
        {
            if (string.IsNullOrEmpty(fontName))
                return false;

            fontName = fontName.ToLower();
            return fontName.Contains("bold") ||
                   fontName.Contains("heavy") ||
                   fontName.Contains("black") ||
                   fontName.Contains("semibold") ||
                   fontName.Contains("demibold");
        }

        private bool IsItalicFont(string fontName)
        {
            if (string.IsNullOrEmpty(fontName))
                return false;

            fontName = fontName.ToLower();
            return fontName.Contains("italic") ||
                   fontName.Contains("oblique") ||
                   fontName.Contains("slanted");
        }

        private bool IsMonospaceFont(string fontName)
        {
            return false;

            if (string.IsNullOrEmpty(fontName))
                return false;

            fontName = fontName.ToLower();
            return fontName.Contains("courier") ||
                   fontName.Contains("mono") ||
                   fontName.Contains("console") ||
                   fontName.Contains("code") ||
                   fontName.Contains("fixed");
        }

        private bool IsBoldFromFontWeight(PdfFont font)
        {
            if (font == null)
                return false;

            try
            {
                var fontProgram = font.GetFontProgram();
                if (fontProgram != null)
                {
                    var fontNames = fontProgram.GetFontNames();
                    int weight = fontNames.GetFontWeight();
                    // Font weights: 400 is normal, 700+ is bold
                    return weight >= 700;
                }
            }
            catch
            {
                // If we can't get weight, fall back to name detection
            }

            return false;
        }

        private bool IsItalicFromFontStyle(PdfFont font)
        {
            if (font == null)
                return false;

            try
            {
                var fontProgram = font.GetFontProgram();
                if (fontProgram != null)
                {
                    var fontNames = fontProgram.GetFontNames();
                    // Check if font is marked as italic
                    return fontNames.GetStyle()?.ToLower().Contains("italic") ?? false;
                }
            }
            catch
            {
                // If we can't get style, fall back to name detection
            }

            return false;
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
            public bool IsBold { get; set; }
            public bool IsItalic { get; set; }
            public bool IsUnderlined { get; set; }
            public bool IsStrikethrough { get; set; }
            public bool IsMonospace { get; set; }
            public float Rise { get; set; }
        }

        private class LineInfo
        {
            public float Y { get; set; }
            public List<PreservedTextRenderInfo> Chunks { get; set; }
        }

        private class FormattedSegment
        {
            public List<PreservedTextRenderInfo> Chunks { get; set; }
            public bool IsBold { get; set; }
            public bool IsItalic { get; set; }
            public bool IsUnderlined { get; set; }
            public bool IsStrikethrough { get; set; }
            public bool IsMonospace { get; set; }
        }
    }
}