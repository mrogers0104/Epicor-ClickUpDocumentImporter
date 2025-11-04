using ClickUpDocumentImporter.Helpers;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Data;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Pdf.Xobject;

//using iTextSharp.text;
//using Org.BouncyCastle.Tsp;
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

            ConsoleHelper.WriteSeparator();
            ConsoleHelper.WriteInfo($"~~~~~ PDF: {Path.GetFileName(pdfFilePath)} ~~~~~");
            ConsoleHelper.WriteInfo($"Found {images.Count} images in PDF");

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

        private static List<FormattedTextBlock> ExtractFormattedTextBlocks(PdfPage page)
        {
            var strategy = new FormattedTextExtractionStrategy();
            PdfTextExtractor.GetTextFromPage(page, strategy);
            return strategy.GetFormattedBlocks();
        }

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
                                ConsoleHelper.WriteError($"Error extracting image on page {pageNum}: {ex.Message}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ConsoleHelper.WriteError($"An error occurred processing the PDF: {ex.Message}");
            }

            ConsoleHelper.WriteInfo($"Extracted {images.Count} images from PDF");
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
    }
}