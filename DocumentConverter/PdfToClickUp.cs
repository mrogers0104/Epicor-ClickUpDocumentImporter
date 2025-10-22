using ClickUpDocumentImporter.Helpers;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Data;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Pdf.Xobject;
using iTextSharp.text;
using Microsoft.VisualBasic;
using Org.BouncyCastle.Tsp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;
using static System.Runtime.InteropServices.JavaScript.JSType;
using Path = System.IO.Path;

namespace ClickUpDocumentImporter.DocumentConverter
{
    /// <summary>
    /// Convert PDF document to a ClickUp page.
    ///
    /// Key Points:
    /// 1. <b>Position-aware extraction</b>: Uses LocationTextExtractionStrategy to extract text with positional information(X, Y coordinates)
    /// 2. <b>Content merging</b>: Combines text chunks and images based on their vertical position(Y coordinate) on the page
    /// 3. <b>Proper ordering</b>: Sorts content by position to maintain the original document layout
    /// 4. <b></b>Text grouping</b>: Groups consecutive text chunks together while respecting position breaks(where images should appear)
    ///
    /// </b>Important notes</b>:
    ///
    /// Your PdfImageData class needs to include position information(X and Y properties). You'll need to modify PdfImageExtractor.ExtractImagesFromPdf() to capture image positions
    /// The threshold value(50) in GroupConsecutiveText may need adjustment based on your PDF structure
    /// PDF coordinates start at bottom-left, so higher Y values mean higher on the page
    ///
    /// If your PdfImageExtractor doesn't currently capture position data, you'll need to enhance it to extract image transformation matrices and calculate positions.
    /// </summary>
    internal static class PdfToClickUp
    {

        //public static async Task ConvertPdfToClickUpAsync(
        //    string pdfFilePath,
        //    HttpClient clickupClient,
        //    string workspaceId,
        //    string wikiId,
        //    string listId,
        //    string parentPageId = null
        //)
        //{
        //    var builder = new ClickUpDocumentBuilder(clickupClient);

        //    // Extract images from PDF with position information
        //    //var images = PdfImageExtractor.ExtractImagesFromPdf(pdfFilePath);
        //    var images = ExtractImagesFromPdf(pdfFilePath);
        //    Console.WriteLine($"Found {images.Count} images in PDF");

        //    using (PdfDocument pdfDoc = new PdfDocument(new PdfReader(pdfFilePath)))
        //    {
        //        builder.AddHeading(Path.GetFileNameWithoutExtension(pdfFilePath), 1);

        //        for (int pageNum = 1; pageNum <= pdfDoc.GetNumberOfPages(); pageNum++)
        //        {
        //            var page = pdfDoc.GetPage(pageNum);

        //            // Extract text with positional information
        //            var locationStrategy = new LocationTextExtractionStrategy();
        //            PdfTextExtractor.GetTextFromPage(page, locationStrategy);

        //            // Get text chunks with their positions
        //            var textChunks = GetTextChunksWithPositions(page);

        //            // Get images for this page with positions
        //            var pageImages = images.Where(img => img.PageNumber == pageNum)
        //                                   .OrderBy(img => img.Y) // Sort by Y position (top to bottom)
        //                                   .ToList();

        //            if (textChunks.Count == 0 && pageImages.Count == 0)
        //                continue;

        //            builder.AddHeading($"Page {pageNum}", 2);

        //            // Merge text and images based on vertical position
        //            await MergeContentByPosition(builder, textChunks, pageImages, listId);
        //        }
        //    }

        //    // Create the ClickUp page
        //    string pageName = System.IO.Path.GetFileNameWithoutExtension(pdfFilePath);
        //    string pageId = await builder.CreateAndPopulatePageAsync(
        //        workspaceId,
        //        wikiId,
        //        pageName,
        //        parentPageId,
        //        uploadMethod: "task",
        //        listIdForTaskUpload: listId
        //    );

        //    Console.WriteLine($"\n✓ Created ClickUp page: {pageName} (ID: {pageId})");
        //}

        //private static List<TextChunk> GetTextChunksWithPositions(PdfPage page)
        //{
        //    var strategy = new CustomLocationTextExtractionStrategy();
        //    PdfTextExtractor.GetTextFromPage(page, strategy);
        //    return strategy.GetTextChunks();
        //}

        //private static async Task MergeContentByPosition(
        //    ClickUpDocumentBuilder builder,
        //    List<TextChunk> textChunks,
        //    List<ImageData> images,
        //    string listId)
        //{
        //    // Combine text and images into a unified list with positions
        //    var contentItems = new List<ContentItem>();

        //    // Add text chunks
        //    foreach (var chunk in textChunks)
        //    {
        //        contentItems.Add(new ContentItem
        //        {
        //            Type = ContentType.Text,
        //            Position = chunk.Y,
        //            TextContent = chunk.Text,
        //            TextChunk = chunk
        //        });
        //    }

        //    // Add images
        //    foreach (var image in images)
        //    {
        //        contentItems.Add(new ContentItem
        //        {
        //            Type = ContentType.Image,
        //            Position = image.Y,
        //            ImageData = image
        //        });
        //    }

        //    // Sort by vertical position (top to bottom)
        //    // Note: PDF coordinates start at bottom-left, so higher Y = higher on page
        //    contentItems = contentItems.OrderByDescending(c => c.Position).ToList();

        //    // Group consecutive text chunks together
        //    var groupedContent = GroupConsecutiveText(contentItems);

        //    // Add content in order
        //    foreach (var item in groupedContent)
        //    {
        //        if (item.Type == ContentType.Text)
        //        {
        //            if (!string.IsNullOrWhiteSpace(item.TextContent))
        //            {
        //                builder.AddParagraph(item.TextContent);
        //            }
        //        }
        //        else if (item.Type == ContentType.Image)
        //        {
        //            await builder.AddImage(
        //                item.ImageData.Data,
        //                item.ImageData.FileName,
        //                listId
        //            );
        //            Console.WriteLine($"Added image at position {item.Position:F2}: {item.ImageData.FileName}");
        //        }
        //    }
        //}

        //private static List<ContentItem> GroupConsecutiveText(List<ContentItem> items)
        //{
        //    var grouped = new List<ContentItem>();
        //    StringBuilder currentText = null;
        //    double lastPosition = double.MaxValue;

        //    foreach (var item in items)
        //    {
        //        if (item.Type == ContentType.Text)
        //        {
        //            if (currentText == null)
        //            {
        //                currentText = new StringBuilder(item.TextContent);
        //                lastPosition = item.Position;
        //            }
        //            else
        //            {
        //                // If text is close together (within reasonable line spacing), combine it
        //                if (Math.Abs(lastPosition - item.Position) < 50) // Adjust threshold as needed
        //                {
        //                    currentText.Append(" ").Append(item.TextContent);
        //                }
        //                else
        //                {
        //                    // Position jump detected, flush current text
        //                    grouped.Add(new ContentItem
        //                    {
        //                        Type = ContentType.Text,
        //                        Position = lastPosition,
        //                        TextContent = currentText.ToString()
        //                    });
        //                    currentText = new StringBuilder(item.TextContent);
        //                    lastPosition = item.Position;
        //                }
        //            }
        //        }
        //        else
        //        {
        //            // Flush any accumulated text before adding image
        //            if (currentText != null)
        //            {
        //                grouped.Add(new ContentItem
        //                {
        //                    Type = ContentType.Text,
        //                    Position = lastPosition,
        //                    TextContent = currentText.ToString()
        //                });
        //                currentText = null;
        //            }
        //            grouped.Add(item);
        //        }
        //    }

        //    // Flush remaining text
        //    if (currentText != null)
        //    {
        //        grouped.Add(new ContentItem
        //        {
        //            Type = ContentType.Text,
        //            Position = lastPosition,
        //            TextContent = currentText.ToString()
        //        });
        //    }

        //    return grouped;
        //}

        //// Helper classes
        //private class TextChunk
        //{
        //    public string Text { get; set; }
        //    public float X { get; set; }
        //    public float Y { get; set; }
        //}

        //private enum ContentType
        //{
        //    Text,
        //    Image
        //}

        //private class ContentItem
        //{
        //    public ContentType Type { get; set; }
        //    public double Position { get; set; }
        //    public string TextContent { get; set; }
        //    public TextChunk TextChunk { get; set; }
        //    public ImageData ImageData { get; set; }
        //}

        //// Custom extraction strategy to get text with positions
        //private class CustomLocationTextExtractionStrategy : LocationTextExtractionStrategy
        //{
        //    private List<TextChunk> chunks = new List<TextChunk>();

        //    public override void EventOccurred(IEventData data, EventType type)
        //    {
        //        if (type == EventType.RENDER_TEXT)
        //        {
        //            var renderInfo = (TextRenderInfo)data;
        //            var baseline = renderInfo.GetBaseline();

        //            chunks.Add(new TextChunk
        //            {
        //                Text = renderInfo.GetText(),
        //                X = baseline.GetStartPoint().Get(0),
        //                Y = baseline.GetStartPoint().Get(1)
        //            });
        //        }

        //        base.EventOccurred(data, type);
        //    }

        //    public List<TextChunk> GetTextChunks()
        //    {
        //        return chunks;
        //    }
        //}




        // -----------------------------------------------------------------------------

        /// <summary>
        ///
        ///
        /// Key features of this solution:
        /// 1. Font-based formatting detection: Detects bold, italic, monospace fonts
        /// 2. Heading detection: Based on font size and bold formatting
        /// 3. List detection: Recognizes bullet points and numbered lists
        /// 4. Inline markdown: Converts to **bold**, *italic*, ~~strikethrough~~, `code`
        /// 5. Block-level elements: Handles headings, lists, code blocks, block quotes
        /// 6. Position-aware: Maintains proper ordering with images
        /// 7. ClickUp compatible: Uses markdown and HTML that ClickUp supports
        ///
        /// The solution groups text chunks by formatting, detects structural elements, and converts them to appropriate markdown that ClickUp can render.
        ///
        /// </summary>
        /// <param name="pdfFilePath"></param>
        /// <param name="clickupClient"></param>
        /// <param name="workspaceId"></param>
        /// <param name="wikiId"></param>
        /// <param name="listId"></param>
        /// <param name="parentPageId"></param>
        /// <returns></returns>
        public static async Task ConvertPdfToClickUpAsync(
            string pdfFilePath,
            HttpClient clickupClient,
            string workspaceId,
            string wikiId,
            string listId,
            string parentPageId = null
        )
        {
            var builder = new ClickUpDocumentBuilder(clickupClient);

            // Extract images from PDF with position information
            var images = ExtractImagesFromPdf(pdfFilePath);
            //var images = PdfImageExtractor.ExtractImagesFromPdf(pdfFilePath);
            Console.WriteLine($"Found {images.Count} images in PDF");

            using (PdfDocument pdfDoc = new PdfDocument(new PdfReader(pdfFilePath)))
            {
                //builder.AddHeading(Path.GetFileNameWithoutExtension(pdfFilePath), 1);

                for (int pageNum = 1; pageNum <= pdfDoc.GetNumberOfPages(); pageNum++)
                {
                    var page = pdfDoc.GetPage(pageNum);

                    // Extract formatted text blocks with positional information
                    var formattedBlocks = ExtractFormattedTextBlocks(page);

                    // Get images for this page with positions
                    var pageImages = images.Where(img => img.PageNumber == pageNum)
                                           .OrderBy(img => img.Y)
                                           .ToList();

                    if (formattedBlocks.Count == 0 && pageImages.Count == 0)
                        continue;

                    //builder.AddHeading($"Page {pageNum}", 2);

                    // Use the new formatter to convert to markdown
                    var formatter = new PdfToMarkdownFormatter(builder, listId);
                    await formatter.FormatAndAddContent(formattedBlocks, pageImages);

                    //// Merge text and images based on vertical position
                    //await MergeFormattedContentByPosition(builder, formattedBlocks, pageImages, listId);

                }
            }

            // Create the ClickUp page
            string pageName = Path.GetFileNameWithoutExtension(pdfFilePath);
            string pageId = await builder.CreateAndPopulatePageAsync(
                workspaceId,
                wikiId,
                pageName,
                parentPageId,
                uploadMethod: "task",
                listIdForTaskUpload: listId
            );

            Console.WriteLine($"\n✓ Created ClickUp page: {pageName} (ID: {pageId})");
        }


        //public static async Task ConvertPdfToClickUpAsync(
        //    string pdfFilePath,
        //    HttpClient clickupClient,
        //    string workspaceId,
        //    string wikiId,
        //    string listId,
        //    string parentPageId = null
        //)
        //{
        //    var builder = new ClickUpDocumentBuilder(clickupClient);

        //    // Extract images from PDF with position information
        //    //var images = PdfImageExtractor.ExtractImagesFromPdf(pdfFilePath);
        //    var images = ExtractImagesFromPdf(pdfFilePath);
        //    Console.WriteLine($"Found {images.Count} images in PDF");

        //    using (PdfDocument pdfDoc = new PdfDocument(new PdfReader(pdfFilePath)))
        //    {
        //        //builder.AddHeading(Path.GetFileNameWithoutExtension(pdfFilePath), 1);

        //        for (int pageNum = 1; pageNum <= pdfDoc.GetNumberOfPages(); pageNum++)
        //        {
        //            var page = pdfDoc.GetPage(pageNum);

        //            // Extract formatted text blocks with positional information
        //            var formattedBlocks = ExtractFormattedTextBlocks(page);

        //            // Get images for this page with positions
        //            var pageImages = images.Where(img => img.PageNumber == pageNum)
        //                                   .OrderBy(img => img.Y)
        //                                   .ToList();

        //            if (formattedBlocks.Count == 0 && pageImages.Count == 0)
        //                continue;

        //            //builder.AddHeading($"Page {pageNum}", 2);

        //            // Merge text and images based on vertical position
        //            await MergeFormattedContentByPosition(builder, formattedBlocks, pageImages, listId);
        //        }
        //    }

        //    // Create the ClickUp page
        //    string pageName = Path.GetFileNameWithoutExtension(pdfFilePath);
        //    string pageId = await builder.CreateAndPopulatePageAsync(
        //        workspaceId,
        //        wikiId,
        //        pageName,
        //        parentPageId,
        //        uploadMethod: "task",
        //        listIdForTaskUpload: listId
        //    );

        //    Console.WriteLine($"\n✓ Created ClickUp page: {pageName} (ID: {pageId})");
        //}

        private static List<FormattedTextBlock> ExtractFormattedTextBlocks(PdfPage page)
        {
            var strategy = new FormattedTextExtractionStrategy();
            PdfTextExtractor.GetTextFromPage(page, strategy);
            return strategy.GetFormattedBlocks();
        }

        private static async Task MergeFormattedContentByPosition(
            ClickUpDocumentBuilder builder,
            List<FormattedTextBlock> textBlocks,
            List<ImageData> images,
            string listId)
        {
            // Combine text blocks and images into a unified list with positions
            var contentItems = new List<FormattedContentItem>();

            // Add text blocks
            foreach (var block in textBlocks)
            {
                contentItems.Add(new FormattedContentItem
                {
                    Type = FormattedContentType.Text,
                    Position = block.Y,
                    TextBlock = block
                });
            }

            // Add images
            foreach (var image in images)
            {
                contentItems.Add(new FormattedContentItem
                {
                    Type = FormattedContentType.Image,
                    Position = image.Y,
                    ImageData = image
                });
            }

            // Sort by vertical position (top to bottom)
            contentItems = contentItems.OrderByDescending(c => c.Position).ToList();

            // Add content in order with proper formatting
            foreach (var item in contentItems)
            {
                if (item.Type == FormattedContentType.Text)
                {
                    var block = item.TextBlock;

                    if (string.IsNullOrWhiteSpace(block.Text))
                        continue;

                    // Convert to markdown based on formatting
                    string markdownText = ConvertToMarkdown(block);

                    // Determine block type and add accordingly
                    if (block.IsHeading)
                    {
                        int headingLevel = DetermineHeadingLevel(block);
                        builder.AddHeading(block.Text.Trim(), Math.Min(headingLevel + 2, 6)); // +2 because page is already h2
                    }
                    else if (block.IsBulletPoint)
                    {
                        builder.AddBulletPoint(markdownText);
                    }
                    else if (block.IsNumberedList)
                    {
                        builder.AddNumberedListItem("1.", markdownText);
                    }
                    else if (block.IsCodeBlock)
                    {
                        builder.AddCodeBlock(block.Text, block.CodeLanguage ?? "");
                    }
                    else if (block.IsBlockQuote)
                    {
                        builder.AddBlockQuote(markdownText);
                    }
                    else
                    {
                        builder.AddParagraph(markdownText);
                    }
                }
                else if (item.Type == FormattedContentType.Image)
                {
                    await builder.AddImage(
                        item.ImageData.Data,
                        item.ImageData.FileName,
                        listId
                    );
                    Console.WriteLine($"Added image at position {item.Position:F2}: {item.ImageData.FileName}");
                }
            }
        }

        private static string ConvertToMarkdown(FormattedTextBlock block)
        {
            string text = block.Text;

            // Apply inline formatting
            if (block.IsBold && !block.IsItalic)
            {
                text = $"**{text}**";
            }
            else if (block.IsItalic && !block.IsBold)
            {
                text = $"*{text}*";
            }
            else if (block.IsBold && block.IsItalic)
            {
                text = $"***{text}***";
            }

            if (block.IsUnderlined)
            {
                // ClickUp supports underline with HTML
                text = $"<u>{text}</u>";
            }

            if (block.IsStrikethrough)
            {
                text = $"~~{text}~~";
            }

            if (block.IsCode)
            {
                text = $"`{text}`";
            }

            if (block.IsLink)
            {
                text = $"[{text}]({block.LinkUrl})";
            }

            return text;
        }

        private static int DetermineHeadingLevel(FormattedTextBlock block)
        {
            // Determine heading level based on font size
            if (block.FontSize >= 24)
                return 1;
            else if (block.FontSize >= 20)
                return 2;
            else if (block.FontSize >= 16)
                return 3;
            else if (block.FontSize >= 14)
                return 4;
            else
                return 5;
        }

        // Helper classes
        //private class FormattedTextBlock
        //{
        //    public string Text { get; set; }
        //    public float X { get; set; }
        //    public float Y { get; set; }
        //    public float FontSize { get; set; }
        //    public string FontName { get; set; }
        //    public bool IsBold { get; set; }
        //    public bool IsItalic { get; set; }
        //    public bool IsUnderlined { get; set; }
        //    public bool IsStrikethrough { get; set; }
        //    public bool IsCode { get; set; }
        //    public bool IsHeading { get; set; }
        //    public bool IsBulletPoint { get; set; }
        //    public bool IsNumberedList { get; set; }
        //    public bool IsCodeBlock { get; set; }
        //    public bool IsBlockQuote { get; set; }
        //    public bool IsLink { get; set; }
        //    public string LinkUrl { get; set; }
        //    public string CodeLanguage { get; set; }
        //    public string Color { get; set; }
        //}

        private enum FormattedContentType
        {
            Text,
            Image
        }

        private class FormattedContentItem
        {
            public FormattedContentType Type { get; set; }
            public double Position { get; set; }
            public FormattedTextBlock TextBlock { get; set; }
            public ImageData ImageData { get; set; }
        }

        public static List<ImageData> ExtractImagesFromPdf(string pdfFilePath)
        {
            var images = new List<ImageData>();
            int imageIndex = 0;

            string uniqueId = Globals.CreateUniqueImageId(pdfFilePath);

            try
            {
                using (PdfDocument pdfDoc = new PdfDocument(new PdfReader(pdfFilePath)))
                {
                    int numberOfPages = pdfDoc.GetNumberOfPages();

                    for (int pageNum = 1; pageNum <= numberOfPages; pageNum++)
                    {
                        var page = pdfDoc.GetPage(pageNum);

                        // Use a custom event listener to capture image positions
                        var imageEventListener = new ImageRenderListener();
                        var processor = new PdfCanvasProcessor(imageEventListener);
                        processor.ProcessPageContent(page);

                        // Get images with their positions from the listener
                        var pageImages = imageEventListener.GetImages();

                        foreach (var imgInfo in pageImages)
                        {
                            try
                            {
                                byte[] imageBytes = imgInfo.Image.GetImageBytes();
                                string extension = DetermineImageExtension(imgInfo.Image);

                                images.Add(new ImageData
                                {
                                    Data = imageBytes,
                                    FileName = $"pdf_image_{imageIndex}_{uniqueId}{extension}",
                                    Index = imageIndex,
                                    PageNumber = pageNum,
                                    X = imgInfo.X,
                                    Y = imgInfo.Y,
                                    Width = imgInfo.Width,
                                    Height = imgInfo.Height
                                });

                                imageIndex++;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error extracting image on page {pageNum}: {ex.Message}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred processing the PDF: {ex.Message}");
            }

            Console.WriteLine($"Extracted {images.Count} images from PDF");
            return images;
        }

        private static string DetermineImageExtension(PdfImageXObject image)
        {
            // Try to determine from filter
            var filter = image.GetPdfObject().GetAsName(PdfName.Filter);

            if (filter != null)
            {
                if (filter.Equals(PdfName.DCTDecode)) return ".jpg";
                if (filter.Equals(PdfName.JPXDecode)) return ".jp2";
                if (filter.Equals(PdfName.FlateDecode)) return ".png";
            }

            return ".png"; // Default
        }

        // Custom event listener to capture images with their positions
        private class ImageRenderListener : IEventListener
        {
            private List<ImageInfo> images = new List<ImageInfo>();

            public void EventOccurred(IEventData data, EventType type)
            {
                if (type == EventType.RENDER_IMAGE)
                {
                    var renderInfo = (ImageRenderInfo)data;
                    var image = renderInfo.GetImage();

                    if (image != null)
                    {
                        // Get the transformation matrix to determine position
                        var ctm = renderInfo.GetImageCtm();

                        // Extract position from the matrix
                        // The matrix gives us the bottom-left corner of the image
                        float x = ctm.Get(Matrix.I31);
                        float y = ctm.Get(Matrix.I32);

                        // Calculate width and height from the transformation matrix
                        float width = Math.Abs(ctm.Get(Matrix.I11));
                        float height = Math.Abs(ctm.Get(Matrix.I22));

                        images.Add(new ImageInfo
                        {
                            Image = image,
                            X = x,
                            Y = y,
                            Width = width,
                            Height = height
                        });
                    }
                }
            }

            public ICollection<EventType> GetSupportedEvents()
            {
                return new List<EventType> { EventType.RENDER_IMAGE };
            }

            public List<ImageInfo> GetImages()
            {
                return images;
            }
        }

        // Helper class to store image information
        private class ImageInfo
        {
            public PdfImageXObject Image { get; set; }
            public float X { get; set; }
            public float Y { get; set; }
            public float Width { get; set; }
            public float Height { get; set; }
        }

        // Custom extraction strategy to get formatted text
        //private class FormattedTextExtractionStrategy : ITextExtractionStrategy
        //{
        //    private List<FormattedTextBlock> blocks = new List<FormattedTextBlock>();
        //    private List<TextRenderInfo> currentLineChunks = new List<TextRenderInfo>();
        //    private float lastY = float.MaxValue;
        //    private const float LINE_SPACING_THRESHOLD = 5f;

        //    public void EventOccurred(IEventData data, EventType type)
        //    {
        //        if (type == EventType.RENDER_TEXT)
        //        {
        //            var renderInfo = (TextRenderInfo)data;
        //            ProcessTextRenderInfo(renderInfo);
        //        }
        //    }

        //    private void ProcessTextRenderInfo(TextRenderInfo renderInfo)
        //    {
        //        var baseline = renderInfo.GetBaseline();
        //        float y = baseline.GetStartPoint().Get(1);

        //        // Check if we're on a new line
        //        if (Math.Abs(lastY - y) > LINE_SPACING_THRESHOLD && currentLineChunks.Count > 0)
        //        {
        //            // Process the completed line
        //            ProcessLine();
        //            currentLineChunks.Clear();
        //        }

        //        currentLineChunks.Add(renderInfo);
        //        lastY = y;
        //    }

        //    private void ProcessLine()
        //    {
        //        if (currentLineChunks.Count == 0)
        //            return;

        //        // Combine chunks that have similar formatting
        //        var groupedChunks = GroupChunksByFormatting(currentLineChunks);

        //        foreach (var group in groupedChunks)
        //        {
        //            var firstChunk = group[0];
        //            var baseline = firstChunk.GetBaseline();
        //            string text = string.Join("", group.Select(c => c.GetText()));

        //            // Skip empty text
        //            if (string.IsNullOrWhiteSpace(text))
        //                continue;

        //            var font = firstChunk.GetFont();
        //            float fontSize = firstChunk.GetFontSize();
        //            string fontName = font?.GetFontProgram()?.GetFontNames()?.GetFontName() ?? "";

        //            // Detect formatting
        //            bool isBold = IsBoldFont(fontName);
        //            bool isItalic = IsItalicFont(fontName);
        //            bool isHeading = fontSize > 12 && isBold; // Heuristic for headings
        //            bool isBulletPoint = text.TrimStart().StartsWith("•") ||
        //                                text.TrimStart().StartsWith("·") ||
        //                                text.TrimStart().StartsWith("-");
        //            bool isNumberedList = System.Text.RegularExpressions.Regex.IsMatch(
        //                text.TrimStart(), @"^\d+[\.\)]\s");

        //            var block = new FormattedTextBlock
        //            {
        //                Text = text,
        //                X = baseline.GetStartPoint().Get(0),
        //                Y = baseline.GetStartPoint().Get(1),
        //                FontSize = fontSize,
        //                FontName = fontName,
        //                IsBold = isBold,
        //                IsItalic = isItalic,
        //                IsHeading = isHeading,
        //                IsBulletPoint = isBulletPoint,
        //                IsNumberedList = isNumberedList,
        //                IsUnderlined = false, // PDF doesn't easily expose underline
        //                IsStrikethrough = false, // PDF doesn't easily expose strikethrough
        //                IsCode = IsMonospaceFont(fontName),
        //                IsCodeBlock = false,
        //                IsBlockQuote = text.TrimStart().StartsWith(">"),
        //                IsLink = false // Would need to process annotations for links
        //            };

        //            blocks.Add(block);
        //        }
        //    }

        //    private List<List<TextRenderInfo>> GroupChunksByFormatting(List<TextRenderInfo> chunks)
        //    {
        //        var groups = new List<List<TextRenderInfo>>();
        //        var currentGroup = new List<TextRenderInfo>();

        //        foreach (var chunk in chunks)
        //        {
        //            if (currentGroup.Count == 0)
        //            {
        //                currentGroup.Add(chunk);
        //            }
        //            else
        //            {
        //                var lastChunk = currentGroup[currentGroup.Count - 1];

        //                // Check if formatting is similar
        //                if (IsSimilarFormatting(lastChunk, chunk))
        //                {
        //                    currentGroup.Add(chunk);
        //                }
        //                else
        //                {
        //                    groups.Add(new List<TextRenderInfo>(currentGroup));
        //                    currentGroup.Clear();
        //                    currentGroup.Add(chunk);
        //                }
        //            }
        //        }

        //        if (currentGroup.Count > 0)
        //        {
        //            groups.Add(currentGroup);
        //        }

        //        return groups;
        //    }

        //    private bool IsSimilarFormatting(TextRenderInfo chunk1, TextRenderInfo chunk2)
        //    {
        //        var font1 = chunk1.GetFont()?.GetFontProgram()?.GetFontNames()?.GetFontName() ?? "";
        //        var font2 = chunk2.GetFont()?.GetFontProgram()?.GetFontNames()?.GetFontName() ?? "";

        //        return Math.Abs(chunk1.GetFontSize() - chunk2.GetFontSize()) < 0.5f &&
        //               font1.Equals(font2, StringComparison.OrdinalIgnoreCase);
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

        //    public ICollection<EventType> GetSupportedEvents()
        //    {
        //        return new List<EventType> { EventType.RENDER_TEXT };
        //    }

        //    public string GetResultantText()
        //    {
        //        // Process any remaining chunks
        //        if (currentLineChunks.Count > 0)
        //        {
        //            ProcessLine();
        //        }

        //        return string.Join("\n", blocks.Select(b => b.Text));
        //    }

        //    public List<FormattedTextBlock> GetFormattedBlocks()
        //    {
        //        // Process any remaining chunks
        //        if (currentLineChunks.Count > 0)
        //        {
        //            ProcessLine();
        //        }

        //        return blocks;
        //    }
        //}


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
        /// </remarks>
        //private class FormattedTextExtractionStrategy : ITextExtractionStrategy
        //{
        //    private List<FormattedTextBlock> blocks = new List<FormattedTextBlock>();
        //    private List<PreservedTextRenderInfo> currentLineChunks = new List<PreservedTextRenderInfo>();
        //    private float lastY = float.MaxValue;
        //    private const float LINE_SPACING_THRESHOLD = 5f;

        //    public void EventOccurred(IEventData data, EventType type)
        //    {
        //        if (type == EventType.RENDER_TEXT)
        //        {
        //            var renderInfo = (TextRenderInfo)data;

        //            // Preserve the graphics state before processing
        //            renderInfo.PreserveGraphicsState();

        //            ProcessTextRenderInfo(renderInfo);
        //        }
        //    }

        //    private void ProcessTextRenderInfo(TextRenderInfo renderInfo)
        //    {
        //        var baseline = renderInfo.GetBaseline();
        //        float y = baseline.GetStartPoint().Get(1);

        //        // Check if we're on a new line
        //        if (Math.Abs(lastY - y) > LINE_SPACING_THRESHOLD && currentLineChunks.Count > 0)
        //        {
        //            // Process the completed line
        //            ProcessLine();
        //            currentLineChunks.Clear();
        //        }

        //        // Store preserved information
        //        var preservedInfo = new PreservedTextRenderInfo
        //        {
        //            Text = renderInfo.GetText(),
        //            X = baseline.GetStartPoint().Get(0),
        //            Y = baseline.GetStartPoint().Get(1),
        //            FontSize = renderInfo.GetFontSize(),
        //            FontName = renderInfo.GetFont()?.GetFontProgram()?.GetFontNames()?.GetFontName() ?? ""
        //        };

        //        currentLineChunks.Add(preservedInfo);
        //        lastY = y;
        //    }

        //    private void ProcessLine()
        //    {
        //        if (currentLineChunks.Count == 0)
        //            return;

        //        // Group chunks that have similar formatting
        //        var groupedChunks = GroupChunksByFormatting(currentLineChunks);

        //        foreach (var group in groupedChunks)
        //        {
        //            var firstChunk = group[0];
        //            string text = string.Join("", group.Select(c => c.Text));

        //            // Skip empty text
        //            if (string.IsNullOrWhiteSpace(text))
        //                continue;

        //            float fontSize = firstChunk.FontSize;
        //            string fontName = firstChunk.FontName;

        //            // Detect formatting
        //            bool isBold = IsBoldFont(fontName);
        //            bool isItalic = IsItalicFont(fontName);
        //            bool isHeading = fontSize >= 12 && isBold;
        //            bool isBulletPoint = text.TrimStart().StartsWith("•") ||
        //                                text.TrimStart().StartsWith("·") ||
        //                                text.TrimStart().StartsWith("o") ||
        //                                text.TrimStart().StartsWith("-");
        //            bool isNumberedList = System.Text.RegularExpressions.Regex.IsMatch(
        //                text.TrimStart(), @"^\d+[\.\)]\s");

        //            if (isHeading)
        //            {
        //                text = text.Trim();
        //            }

        //            var block = new FormattedTextBlock
        //            {
        //                Text = text,
        //                X = firstChunk.X,
        //                Y = firstChunk.Y,
        //                FontSize = fontSize,
        //                FontName = fontName,
        //                IsBold = isBold,
        //                IsItalic = isItalic,
        //                IsHeading = isHeading,
        //                IsBulletPoint = isBulletPoint,
        //                IsNumberedList = isNumberedList,
        //                IsUnderlined = false,
        //                IsStrikethrough = false,
        //                IsCode = IsMonospaceFont(fontName),
        //                IsCodeBlock = false,
        //                IsBlockQuote = text.TrimStart().StartsWith(">"),
        //                IsLink = false
        //            };

        //            blocks.Add(block);
        //        }
        //    }

        //    private List<List<PreservedTextRenderInfo>> GroupChunksByFormatting(List<PreservedTextRenderInfo> chunks)
        //    {
        //        var groups = new List<List<PreservedTextRenderInfo>>();
        //        var currentGroup = new List<PreservedTextRenderInfo>();

        //        foreach (var chunk in chunks)
        //        {
        //            if (currentGroup.Count == 0)
        //            {
        //                currentGroup.Add(chunk);
        //            }
        //            else
        //            {
        //                var lastChunk = currentGroup[currentGroup.Count - 1];

        //                // Check if formatting is similar
        //                if (IsSimilarFormatting(lastChunk, chunk))
        //                {
        //                    currentGroup.Add(chunk);
        //                }
        //                else
        //                {
        //                    groups.Add(new List<PreservedTextRenderInfo>(currentGroup));
        //                    currentGroup.Clear();
        //                    currentGroup.Add(chunk);
        //                }
        //            }
        //        }

        //        if (currentGroup.Count > 0)
        //        {
        //            groups.Add(currentGroup);
        //        }

        //        return groups;
        //    }

        //    private bool IsSimilarFormatting(PreservedTextRenderInfo chunk1, PreservedTextRenderInfo chunk2)
        //    {
        //        return Math.Abs(chunk1.FontSize - chunk2.FontSize) < 0.5f &&
        //               chunk1.FontName.Equals(chunk2.FontName, StringComparison.OrdinalIgnoreCase);
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

        //    public ICollection<EventType> GetSupportedEvents()
        //    {
        //        return new List<EventType> { EventType.RENDER_TEXT };
        //    }

        //    public string GetResultantText()
        //    {
        //        // Process any remaining chunks
        //        if (currentLineChunks.Count > 0)
        //        {
        //            ProcessLine();
        //        }

        //        return string.Join("\n", blocks.Select(b => b.Text));
        //    }

        //    public List<FormattedTextBlock> GetFormattedBlocks()
        //    {
        //        // Process any remaining chunks
        //        if (currentLineChunks.Count > 0)
        //        {
        //            ProcessLine();
        //        }

        //        return blocks;
        //    }

        //    // Helper class to store preserved render info
        //    private class PreservedTextRenderInfo
        //    {
        //        public string Text { get; set; }
        //        public float X { get; set; }
        //        public float Y { get; set; }
        //        public float FontSize { get; set; }
        //        public string FontName { get; set; }
        //    }
        //}
    }
}
