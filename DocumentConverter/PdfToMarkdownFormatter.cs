using ClickUpDocumentImporter.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClickUpDocumentImporter.DocumentConverter
{
    /// <summary>
    /// Converts PDF formatted text blocks into proper markdown format compatible with ClickUp
    /// </summary>
    internal class PdfToMarkdownFormatter
    {
        private readonly ClickUpDocumentBuilder _builder;
        private readonly string _listId;

        public PdfToMarkdownFormatter(ClickUpDocumentBuilder builder, string listId)
        {
            _builder = builder;
            _listId = listId;
        }

        /// <summary>
        /// Merges text blocks and images based on their position and converts to markdown
        /// </summary>
        public async Task FormatAndAddContent(List<FormattedTextBlock> textBlocks, List<ImageData> images)
        {
            // Combine text blocks and images into a unified list with positions
            var contentItems = new List<PositionedContent>();

            // Add text blocks
            foreach (var block in textBlocks)
            {
                contentItems.Add(new PositionedContent
                {
                    Type = ContentType.Text,
                    Position = block.Y,
                    TextBlock = block
                });
            }

            // Add images
            foreach (var image in images)
            {
                contentItems.Add(new PositionedContent
                {
                    Type = ContentType.Image,
                    Position = image.Y,
                    ImageData = image
                });
            }

            // Sort by vertical position (top to bottom)
            contentItems = contentItems.OrderByDescending(c => c.Position).ToList();

            // Process content items with context awareness
            await ProcessContentItems(contentItems);
        }

        private async Task ProcessContentItems(List<PositionedContent> items)
        {
            var listContext = new ListContext();

            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i];

                if (item.Type == ContentType.Image)
                {
                    // Close any open lists before adding image
                    CloseActiveLists(listContext);

                    await _builder.AddImage(
                        item.ImageData.Data,
                        item.ImageData.FileName,
                        _listId
                    );
                    Console.WriteLine($"Added image at position {item.Position:F2}: {item.ImageData.FileName}");
                }
                else if (item.Type == ContentType.Text)
                {
                    await ProcessTextBlock(item.TextBlock, listContext);
                }
            }

            // Close any remaining open lists
            CloseActiveLists(listContext);
        }

        // --------------------------------------------------------------
        private async Task ProcessTextBlock(FormattedTextBlock block, ListContext listContext)
        {
            if (string.IsNullOrWhiteSpace(block.Text))
                return;

            string text = block.Text.Trim();

            // Text already has inline markdown formatting applied, so we don't need to apply it again

            // Handle different block types
            if (block.IsHeading)
            {
                CloseActiveLists(listContext);
                int headingLevel = DetermineHeadingLevel(block);
                // Remove any markdown formatting from heading text for cleaner display
                string cleanHeading = RemoveInlineMarkdown(text);
                _builder.AddHeading(cleanHeading, headingLevel);
            }
            else if (block.IsCodeBlock)
            {
                CloseActiveLists(listContext);
                // Remove any markdown formatting from code blocks
                string cleanCode = RemoveInlineMarkdown(text);
                _builder.AddCodeBlock(cleanCode, block.CodeLanguage ?? "");
            }
            else if (block.IsBlockQuote)
            {
                CloseActiveLists(listContext);
                string quoteText = text.TrimStart('>', ' ');
                _builder.AddBlockQuote(quoteText); // Keep inline formatting in quotes
            }
            else if (block.IsBulletPoint)
            {
                int indentLevel = GetIndentLevel(block);

                if (listContext.InNumberedList)
                {
                    listContext.InNumberedList = false;
                }

                listContext.InBulletList = true;
                string bulletText = ExtractListItemText(text, isBullet: true);

                string indent = new string(' ', indentLevel * 2);
                _builder.AddMarkdown($"{indent}- {bulletText}"); // Keep inline formatting in list items
            }
            else if (block.IsNumberedList)
            {
                int indentLevel = GetIndentLevel(block);

                if (listContext.InBulletList)
                {
                    listContext.InBulletList = false;
                }

                listContext.InNumberedList = true;
                string numberedText = ExtractListItemText(text, isBullet: false);

                string indent = new string(' ', indentLevel * 2);
                _builder.AddMarkdown($"{indent}1. {numberedText}"); // Keep inline formatting in list items
            }
            else
            {
                // Regular paragraph - text already has inline formatting
                CloseActiveLists(listContext);
                _builder.AddParagraph(text);
            }
        }

        private string RemoveInlineMarkdown(string text)
        {
            // Remove markdown formatting for headings and code blocks
            text = System.Text.RegularExpressions.Regex.Replace(text, @"\*\*\*(.*?)\*\*\*", "$1"); // Bold+Italic
            text = System.Text.RegularExpressions.Regex.Replace(text, @"\*\*(.*?)\*\*", "$1"); // Bold
            text = System.Text.RegularExpressions.Regex.Replace(text, @"\*(.*?)\*", "$1"); // Italic
            text = System.Text.RegularExpressions.Regex.Replace(text, @"~~(.*?)~~", "$1"); // Strikethrough
            text = System.Text.RegularExpressions.Regex.Replace(text, @"`(.*?)`", "$1"); // Inline code
            text = System.Text.RegularExpressions.Regex.Replace(text, @"<u>(.*?)</u>", "$1"); // Underline

            return text;
        }

        // --------------------------------------------------------------

        //private async Task ProcessTextBlock(FormattedTextBlock block, ListContext listContext)
        //{
        //    if (string.IsNullOrWhiteSpace(block.Text))
        //        return;

        //    string text = block.Text.Trim();

        //    // Handle different block types
        //    if (block.IsHeading)
        //    {
        //        CloseActiveLists(listContext);
        //        int headingLevel = DetermineHeadingLevel(block);
        //        _builder.AddHeading(text, headingLevel);
        //    }
        //    else if (block.IsCodeBlock)
        //    {
        //        CloseActiveLists(listContext);
        //        _builder.AddCodeBlock(block.Text, block.CodeLanguage ?? "");
        //    }
        //    else if (block.IsBlockQuote)
        //    {
        //        CloseActiveLists(listContext);
        //        string quoteText = text.TrimStart('>', ' ');
        //        string formattedText = ApplyInlineFormatting(block, quoteText);
        //        _builder.AddBlockQuote(formattedText);
        //    }
        //    else if (block.IsBulletPoint)
        //    {
        //        // Handle nested bullet points
        //        int indentLevel = GetIndentLevel(block);

        //        // Close numbered list if we were in one
        //        if (listContext.InNumberedList)
        //        {
        //            listContext.InNumberedList = false;
        //        }

        //        listContext.InBulletList = true;
        //        string bulletText = ExtractListItemText(text, isBullet: true);
        //        string formattedText = ApplyInlineFormatting(block, bulletText);

        //        // Add indentation for nested lists
        //        string indent = new string(' ', indentLevel * 2);
        //        _builder.AddMarkdown($"{indent}- {formattedText}");
        //    }
        //    else if (block.IsNumberedList)
        //    {
        //        // Handle nested numbered lists
        //        int indentLevel = GetIndentLevel(block);

        //        // Close bullet list if we were in one
        //        if (listContext.InBulletList)
        //        {
        //            listContext.InBulletList = false;
        //        }

        //        listContext.InNumberedList = true;
        //        string numberedText = ExtractListItemText(text, isBullet: false);
        //        string formattedText = ApplyInlineFormatting(block, numberedText);

        //        // Add indentation for nested lists
        //        string indent = new string(' ', indentLevel * 2);
        //        _builder.AddMarkdown($"{indent}1. {formattedText}");
        //    }
        //    else
        //    {
        //        // Regular paragraph
        //        CloseActiveLists(listContext);
        //        string formattedText = ApplyInlineFormatting(block, text);
        //        _builder.AddParagraph(formattedText);
        //    }
        //}

        //private string ApplyInlineFormatting(FormattedTextBlock block, string text)
        //{
        //    // Apply inline markdown formatting based on detected styles

        //    // Handle code spans (inline code)
        //    if (block.IsCode && !block.IsCodeBlock)
        //    {
        //        text = $"`{text}`";
        //    }

        //    // Handle bold and italic combinations
        //    if (block.IsBold && block.IsItalic)
        //    {
        //        text = $"***{text}***";
        //    }
        //    else if (block.IsBold)
        //    {
        //        text = $"**{text}**";
        //    }
        //    else if (block.IsItalic)
        //    {
        //        text = $"*{text}*";
        //    }

        //    // Handle underline (using HTML since markdown doesn't have native underline)
        //    if (block.IsUnderlined)
        //    {
        //        text = $"<u>{text}</u>";
        //    }

        //    // Handle strikethrough
        //    if (block.IsStrikethrough)
        //    {
        //        text = $"~~{text}~~";
        //    }

        //    // Handle links
        //    if (block.IsLink && !string.IsNullOrEmpty(block.LinkUrl))
        //    {
        //        text = $"[{text}]({block.LinkUrl})";
        //    }

        //    // Handle text color if needed (using HTML)
        //    if (!string.IsNullOrEmpty(block.Color) && block.Color.ToLower() != "#000000" && block.Color.ToLower() != "black")
        //    {
        //        text = $"<span style=\"color: {block.Color}\">{text}</span>";
        //    }

        //    return text;
        //}

        private string ExtractListItemText(string text, bool isBullet)
        {
            if (isBullet)
            {
                // Remove bullet markers: o, •, ·, -, *
                text = text.TrimStart();
                if (text.StartsWith("o ") || text.StartsWith("• ") ||
                    text.StartsWith("· ") || text.StartsWith("- ") ||
                    text.StartsWith("* "))
                {
                    text = text.Substring(2);
                }
            }
            else
            {
                // Remove numbered list markers: 1., 2), etc.
                var match = System.Text.RegularExpressions.Regex.Match(text.TrimStart(), @"^\d+[\.\)]\s*");
                if (match.Success)
                {
                    text = text.Substring(match.Length);
                }
            }

            return text.Trim();
        }

        private int GetIndentLevel(FormattedTextBlock block)
        {
            // Determine indent level based on X position
            // Assuming standard indent is around 36 points (0.5 inch)
            const float INDENT_SIZE = 36f;

            if (block.X < INDENT_SIZE)
                return 0;
            else if (block.X < INDENT_SIZE * 2)
                return 1;
            else if (block.X < INDENT_SIZE * 3)
                return 2;
            else
                return 3;
        }

        private int DetermineHeadingLevel(FormattedTextBlock block)
        {
            // Determine heading level based on font size
            // These thresholds can be adjusted based on your PDF documents
            if (block.FontSize >= 28)
                return 1;
            else if (block.FontSize >= 24)
                return 2;
            else if (block.FontSize >= 20)
                return 3;
            else if (block.FontSize >= 16)
                return 4;
            else if (block.FontSize >= 14)
                return 5;
            else
                return 6;
        }

        private void CloseActiveLists(ListContext context)
        {
            // Add blank line to properly close lists
            if (context.InBulletList || context.InNumberedList)
            {
                _builder.AddMarkdown(""); // Empty line to close list
                context.InBulletList = false;
                context.InNumberedList = false;
            }
        }

        // Helper classes
        private enum ContentType
        {
            Text,
            Image
        }

        private class PositionedContent
        {
            public ContentType Type { get; set; }
            public double Position { get; set; }
            public FormattedTextBlock TextBlock { get; set; }
            public ImageData ImageData { get; set; }
        }

        private class ListContext
        {
            public bool InBulletList { get; set; }
            public bool InNumberedList { get; set; }
        }
    }
}
