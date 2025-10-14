using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Pdf.Xobject;
using ClickUpDocumentImporter.Helpers;


namespace ClickUpDocumentImporter.DocumentConverter
{
    /// <summary>
    ///  COMPLETE EXAMPLE: Extract and Upload to ClickUp
    /// </summary>
    public class CompleteDocumentConverter
    {
        public static async Task ConvertWordToClickUpAsync(
            string wordFilePath,
            string apiToken,
            string workspaceId,
            string parentPageId = null)
        {
            var builder = new ClickUpDocumentBuilder(apiToken);

            // Extract images first
            var images = WordImageExtractor.ExtractImagesFromWord(wordFilePath);
            Console.WriteLine($"Found {images.Count} images in Word document");

            // Build content with text and images
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(wordFilePath, false))
            {
                var body = wordDoc.MainDocumentPart.Document.Body;
                int currentImageIndex = 0;

                foreach (var element in body.Elements())
                {
                    if (element is Paragraph para)
                    {
                        // Check if paragraph contains an image
                        var drawing = para.Descendants<Drawing>().FirstOrDefault();

                        if (drawing != null)
                        {
                            // This paragraph contains an image
                            if (currentImageIndex < images.Count)
                            {
                                var imageData = images[currentImageIndex];
                                await builder.AddImageAsync(
                                    imageData.Data,
                                    imageData.FileName,
                                    workspaceId
                                );
                                Console.WriteLine($"Added image: {imageData.FileName}");
                                currentImageIndex++;
                            }
                        }
                        else
                        {
                            // Regular text paragraph
                            string text = para.InnerText;
                            if (!string.IsNullOrWhiteSpace(text))
                            {
                                // Check if it's a heading
                                var paragraphProperties = para.ParagraphProperties;
                                var styleId = paragraphProperties?.ParagraphStyleId?.Val?.Value;

                                if (styleId != null && styleId.StartsWith("Heading"))
                                {
                                    int level = int.TryParse(styleId.Replace("Heading", ""), out int l) ? l : 1;
                                    builder.AddHeading(text, level);
                                    Console.WriteLine($"Added heading (level {level}): {text}");
                                }
                                else
                                {
                                    builder.AddParagraph(text);
                                    Console.WriteLine($"Added paragraph: {text.Substring(0, Math.Min(50, text.Length))}...");
                                }
                            }
                        }
                    }
                }
            }

            // Create the ClickUp page
            string pageName = Path.GetFileNameWithoutExtension(wordFilePath);
            string pageId = await builder.CreateAndPopulatePageAsync(
                workspaceId,
                pageName,
                parentPageId
            );

            Console.WriteLine($"\n✓ Created ClickUp page: {pageName} (ID: {pageId})");
        }

        public static async Task ConvertPdfToClickUpAsync(
            string pdfFilePath,
            string apiToken,
            string workspaceId,
            string parentPageId = null)
        {
            var builder = new ClickUpDocumentBuilder(apiToken);

            // Extract images from PDF
            var images = PdfImageExtractor.ExtractImagesFromPdf(pdfFilePath);
            Console.WriteLine($"Found {images.Count} images in PDF");

            // Extract text and structure from PDF
            using (PdfDocument pdfDoc = new PdfDocument(new PdfReader(pdfFilePath)))
            {
                builder.AddHeading(Path.GetFileNameWithoutExtension(pdfFilePath), 1);

                for (int pageNum = 1; pageNum <= pdfDoc.GetNumberOfPages(); pageNum++)
                {
                    var page = pdfDoc.GetPage(pageNum);

                    // Extract text
                    var strategy = new SimpleTextExtractionStrategy();
                    string pageText = PdfTextExtractor.GetTextFromPage(page, strategy);

                    if (!string.IsNullOrWhiteSpace(pageText))
                    {
                        builder.AddHeading($"Page {pageNum}", 2);
                        builder.AddParagraph(pageText);
                    }

                    // Add images from this page
                    var pageImages = images.Where(img => img.PageNumber == pageNum).ToList();
                    foreach (var imageData in pageImages)
                    {
                        await builder.AddImageAsync(
                            imageData.Data,
                            imageData.FileName,
                            workspaceId
                        );
                        Console.WriteLine($"Added image from page {pageNum}: {imageData.FileName}");
                    }
                }
            }

            // Create the ClickUp page
            string pageName = Path.GetFileNameWithoutExtension(pdfFilePath);
            string pageId = await builder.CreateAndPopulatePageAsync(
                workspaceId,
                pageName,
                parentPageId
            );

            Console.WriteLine($"\n✓ Created ClickUp page: {pageName} (ID: {pageId})");
        }
    }

    //// ===== SIMPLE USAGE EXAMPLES =====
    //class Program
    //{
    //    static async Task Main(string[] args)
    //    {
    //        string apiToken = "your_api_token";
    //        string workspaceId = "your_workspace_id";

    //        // Example 1: Convert Word document
    //        await CompleteDocumentConverter.ConvertWordToClickUpAsync(
    //            "my_document.docx",
    //            apiToken,
    //            workspaceId
    //        );

    //        // Example 2: Convert PDF document
    //        await CompleteDocumentConverter.ConvertPdfToClickUpAsync(
    //            "my_document.pdf",
    //            apiToken,
    //            workspaceId
    //        );

    //        // Example 3: Just extract images (if you need them separately)
    //        var wordImages = WordImageExtractor.ExtractImagesFromWord("document.docx");
    //        foreach (var img in wordImages)
    //        {
    //            // Save locally if needed
    //            File.WriteAllBytes($"extracted_{img.FileName}", img.Data);
    //            Console.WriteLine($"Saved: {img.FileName}");
    //        }
    //    }
    //}
}
