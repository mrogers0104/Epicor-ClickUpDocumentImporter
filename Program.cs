using System;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing;
using A = DocumentFormat.OpenXml.Drawing; // Alias needed for Drawing
//using A = DocumentFormat.OpenXml.Wordprocessing; // Alias needed for Drawing
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Pdf.Xobject;
using Path = System.IO.Path;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using ClickUpDocumentImporter.Helpers;
using System.Diagnostics;
using ClickUpDocumentImporter.DocumentConverter;


namespace ClickUpDocumentImporter
{
    class ImageInfo
    {
        public byte[] ImageData { get; set; }
        public string FileName { get; set; }
        public int Position { get; set; } // Position in text to insert image
        public string ContentType { get; set; }
    }


    /// <summary>
    /// Import a Word or PDF document into ClickUp as a page,
    ///
    ///  * Key Features:
    ///     * - Extracts images from Word documents(embedded images)
    ///     * - Extracts images from PDF documents
    ///     * - Uploads images to ClickUp via API
    ///     * - Maintains image position in the document
    ///     * - Replaces placeholders with actual image URLs
    ///     * - Handles multiple images per document
    ///     * - Preserves text formatting(headers, tables, paragraphs)
    /// </summary>
    /// <remarks>
    ///  NOTES
    ///     * - Images are uploaded to ClickUp's attachment API
    ///     * - The image URLs returned by ClickUp are then embedded in markdown
    ///     * - Rate limiting is important when uploading many images
    ///     * - Some PDF images may need format conversion depending on PDF encoding
    ///
    ///  REQUIRED NUGET PACKAGES
    ///     * Install-Package DocumentFormat.OpenXml
    ///     * Install-Package itext
    ///     * Install-Package System.Text.Json
    /// </remarks>
    internal class Program
    {
        // *** Wiki URL: https://app.clickup.com/9010105092/docs/8cgpjr4-40131/8cgpjr4-23231
        // *** User Documentation URL: https://app.clickup.com/9010105092/docs/8cgpjr4-40131/8cgpjr4-25471
        // *** Developer Documentation URL: https://app.clickup.com/9010105092/docs/8cgpjr4-40131/8cgpjr4-25571
        // *** Screen Logic Customization page: https://app.clickup.com/9010105092/docs/8cgpjr4-40131
        private static readonly string CLICKUP_API_TOKEN = Globals.CLICKUP_API_KEY;
        private static readonly string WORKSPACE_ID = Globals.CLICKUP_WORKSPACE_ID;
        private static readonly string SPACE_ID = "8cgpjr4-40131";
        private static readonly string WIKI_ID = "8cgpjr4-40131";  // Wiki
        private static readonly string PARENT_PAGE_ID = "8cgpjr4-25571"; // Optional: for nesting pages
        private static readonly string LIST_ID = Globals.CLICKUP_LIST_ID; // List to add images

        private static HttpClient clickupClient = new HttpClient();
        private static List<PageInfo> allPages = new List<PageInfo>();
        private static int imageCounter = 0;
        private static readonly string addPagesToDoc = "Screen Logic Customization"; // Page to add Documents

#if DEBUG
        private static string documentsFolder = @"C:\temp\CustomizedScreenLogic";
#else
        private static string documentsFolder = @"C:\temp\";
#endif

        static async Task Main(string[] args)
        {
            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            ConsoleHelper.WriteHeader($"S T A R T: Import Documents to ClickUp  [v{version}]");
            ConsoleHelper.WriteBlank();

            // *** User must enter the directory to the documents to be imported to ClickUp
            if (args.Length <= 0)
            {
                var found = EnterDocumentDirectory();
                if (!found)
                {
                    return;
                }
            }

            // *** Configure HTTP clickupClient
            var client = new ClickUpClient();
            clickupClient = client.ClickUpHttpClient;

            // *** List all pages in your space to select parent page
            var selectionList = await ListPagesInSpace();

            var selectedPageName = GetPageSelection(selectionList);

            //var page = PageExtractor.FindPageByName(allPages, addPagesToDoc, caseSensitive: false);
            var page = PageExtractor.FindPageByName(allPages, selectedPageName, caseSensitive: false);

            if (page == null || page.Id == null)
            {
                //Console.WriteLine($"Page not found for Document Title: {addPagesToDoc}");
                ConsoleHelper.WriteError($"Page not found for Document Title: {addPagesToDoc}");
                return;
            }

            var files = Directory.GetFiles(documentsFolder, "*.*")
                .Where(f => f.EndsWith(".docx") || f.EndsWith(".pdf"))
                .ToArray();

            //Console.WriteLine($"Found {files.Length} documents to import");
            ConsoleHelper.WriteInfo($"Found {files.Length} documents to import");

            string apiToken = CLICKUP_API_TOKEN;
            string workspaceId = WORKSPACE_ID;
            string wikiId = WIKI_ID;
            string parentPageId = page?.Id;
            string listId = LIST_ID;
            foreach (var file in files)
            {
                string ext = Path.GetExtension(file).ToLower();
                if (ext.Equals(".docx"))
                {
                    // *** Convert Word document
                    await CompleteDocumentConverter.ConvertWordToClickUpAsync(
                        file,
                        clickupClient,
                        workspaceId,
                        wikiId,
                        listId: listId,
                        parentPageId: parentPageId
                    );
                }
                else if (ext.Equals(".pdf"))
                {
                    // *** Convert PDF document
                    await PdfToClickUp.ConvertPdfToClickUpAsync( // CompleteDocumentConverter.ConvertPdfToClickUpAsync(
                        file,
                        clickupClient,
                        workspaceId,
                        wikiId,
                        listId: listId,
                        parentPageId: parentPageId
                    );

                }

            }

            ConsoleHelper.WriteLogPath();
            ConsoleHelper.WriteSeparator();
            ConsoleHelper.WriteSuccess("\nImport complete!");
            ConsoleHelper.Pause();
         }

        static bool EnterDocumentDirectory()
        {
            string directoryFolder = string.Empty;

            while (string.IsNullOrEmpty(directoryFolder))
            {
                string input = ConsoleHelper.AskQuestion("Enter document directory path", documentsFolder);

                if (Directory.Exists(input))
                {
                    directoryFolder = input;
                    ConsoleHelper.WriteSuccess($"Directory found: {directoryFolder}");
                }
                else
                {
                    ConsoleHelper.WriteWarning($"Directory does not exist: {input}");

                    if (!ConsoleHelper.AskYesNo("Would you like to enter the directory?", true))
                    {
                        ConsoleHelper.WriteWarning("Run terminated.");
                        return false;
                    }
                }
            }
            documentsFolder = directoryFolder;
            return true;
        }

        static string GetPageSelection(List<SelectionItem> items)
        {
            ConsoleHelper.WriteHeader("Select top level (parent) page");
            //ConsoleHelper.WriteSeparator();

            ConsoleHelper.WriteInfo("Use Arrow Keys to navigate, <ENTER> to select, <ESC> to cancel");
            //ConsoleHelper.WriteSeparator();

            var selected = ConsoleHelper.SelectFromList(items, "Please select parent page:", 2);

            if (selected != null)
            {
                ConsoleHelper.WriteSeparator();
                ConsoleHelper.WriteSuccess($"You selected: {selected.Value}");
                return (string) selected.Value;
            }
            else
            {
                ConsoleHelper.WriteWarning("Selection was cancelled");
                return string.Empty;
            }

        }

        //static (string content, List<ImageInfo> images) ExtractContentWithImages(string filePath)
        //{
        //    string extension = Path.GetExtension(filePath).ToLower();

        //    if (extension == ".docx")
        //    {
        //        return ExtractFromWord(filePath);
        //    }
        //    else if (extension == ".pdf")
        //    {
        //        return ExtractFromPdf(filePath);
        //    }

        //    throw new NotSupportedException($"File type {extension} not supported");
        //}

        //static (string content, List<ImageInfo> images) ExtractFromWord(string filePath)
        //{
        //    var sb = new StringBuilder();
        //    var images = new List<ImageInfo>();
        //    int charPosition = 0;

        //    using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
        //    {
        //        Body body = doc.MainDocumentPart.Document.Body;

        //        foreach (var element in body.Elements())
        //        {
        //            if (element is Paragraph para)
        //            {
        //                // Check for images in paragraph
        //                var drawings = para.Descendants<Drawing>().ToList();

        //                if (drawings.Any())
        //                {
        //                    // Process text before image
        //                    string textBeforeImage = GetParagraphText(para, beforeDrawing: true);
        //                    if (!string.IsNullOrEmpty(textBeforeImage))
        //                    {
        //                        sb.Append(textBeforeImage);
        //                        charPosition += textBeforeImage.Length;
        //                    }

        //                    // Process each image in the paragraph
        //                    foreach (var drawing in drawings)
        //                    {
        //                        var imageInfo = ExtractImageFromDrawing(drawing, doc.MainDocumentPart, charPosition);
        //                        if (imageInfo != null)
        //                        {
        //                            images.Add(imageInfo);
        //                            string placeholder = $"\n![Image_{images.Count}](IMAGE_PLACEHOLDER_{images.Count})\n";
        //                            sb.Append(placeholder);
        //                            charPosition += placeholder.Length;
        //                        }
        //                    }

        //                    sb.AppendLine();
        //                    charPosition += Environment.NewLine.Length;
        //                }
        //                else
        //                {
        //                    // Regular paragraph without images
        //                    string text = para.InnerText;
        //                    var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;

        //                    if (!string.IsNullOrEmpty(text))
        //                    {
        //                        string line = "";
        //                        if (styleId?.StartsWith("Heading") == true)
        //                        {
        //                            int level = int.TryParse(styleId.Replace("Heading", ""), out int l) ? l : 1;
        //                            line = $"{new string('#', level)} {text}\n\n";
        //                        }
        //                        else
        //                        {
        //                            line = $"{text}\n\n";
        //                        }

        //                        sb.Append(line);
        //                        charPosition += line.Length;
        //                    }
        //                }
        //            }
        //            else if (element is Table table)
        //            {
        //                string tableMarkdown = ConvertTableToMarkdown(table);
        //                sb.Append(tableMarkdown);
        //                charPosition += tableMarkdown.Length;
        //            }
        //        }
        //    }

        //    return (sb.ToString(), images);
        //}

        //static string GetParagraphText(Paragraph para, bool beforeDrawing)
        //{
        //    var text = new StringBuilder();
        //    foreach (var run in para.Elements<Run>())
        //    {
        //        if (beforeDrawing && run.Descendants<Drawing>().Any())
        //            break;

        //        text.Append(run.InnerText);
        //    }
        //    return text.ToString();
        //}

        //static ImageInfo ExtractImageFromDrawing(Drawing drawing, MainDocumentPart mainPart, int position)
        //{
        //    try
        //    {
        //        var blip = drawing.Descendants<A.Blip>().FirstOrDefault();
        //        if (blip?.Embed?.Value == null)
        //            return null;

        //        string imageId = blip.Embed.Value;
        //        var imagePart = mainPart.GetPartById(imageId) as ImagePart;

        //        if (imagePart == null)
        //            return null;

        //        using (var stream = imagePart.GetStream())
        //        using (var memoryStream = new MemoryStream())
        //        {
        //            stream.CopyTo(memoryStream);
        //            imageCounter++;

        //            return new ImageInfo
        //            {
        //                ImageData = memoryStream.ToArray(),
        //                FileName = $"image_{imageCounter}{GetImageExtension(imagePart.ContentType)}",
        //                Position = position,
        //                ContentType = imagePart.ContentType
        //            };
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine($"  Warning: Could not extract image - {ex.Message}");
        //        return null;
        //    }
        //}

        //static (string content, List<ImageInfo> images) ExtractFromPdf(string filePath)
        //{
        //    var sb = new StringBuilder();
        //    var images = new List<ImageInfo>();
        //    int charPosition = 0;

        //    using (PdfReader reader = new PdfReader(filePath))
        //    using (PdfDocument pdf = new PdfDocument(reader))
        //    {
        //        for (int pageNum = 1; pageNum <= pdf.GetNumberOfPages(); pageNum++)
        //        {
        //            var page = pdf.GetPage(pageNum);

        //            // Extract text
        //            string pageText = PdfTextExtractor.GetTextFromPage(page);
        //            sb.AppendLine(pageText);
        //            sb.AppendLine();
        //            charPosition = sb.Length;

        //            // Extract images from page
        //            var pageImages = ExtractImagesFromPdfPage(page, pageNum);
        //            foreach (var img in pageImages)
        //            {
        //                img.Position = charPosition;
        //                images.Add(img);

        //                string placeholder = $"\n![Image_{images.Count}](IMAGE_PLACEHOLDER_{images.Count})\n";
        //                sb.Append(placeholder);
        //                charPosition += placeholder.Length;
        //            }
        //        }
        //    }

        //    return (sb.ToString(), images);
        //}

        //static List<ImageInfo> ExtractImagesFromPdfPage(iText.Kernel.Pdf.PdfPage page, int pageNum)
        //{
        //    var images = new List<ImageInfo>();
        //    var resources = page.GetResources();
        //    var xObjects = resources.GetResource(iText.Kernel.Pdf.PdfName.XObject);

        //    if (xObjects == null || !xObjects.IsDictionary())
        //        return images;

        //    var xObjDict = (iText.Kernel.Pdf.PdfDictionary)xObjects;

        //    foreach (var key in xObjDict.KeySet())
        //    {
        //        try
        //        {
        //            var xObj = xObjDict.Get(key);
        //            if (xObj == null) continue;

        //            var stream = (iText.Kernel.Pdf.PdfStream)xObj;
        //            if (stream == null) continue;

        //            var subtype = stream.GetAsName(iText.Kernel.Pdf.PdfName.Subtype);
        //            if (subtype == null || !subtype.Equals(iText.Kernel.Pdf.PdfName.Image))
        //                continue;

        //            var pdfImage = new PdfImageXObject(stream);
        //            byte[] imageBytes = pdfImage.GetImageBytes();

        //            imageCounter++;

        //            images.Add(new ImageInfo
        //            {
        //                ImageData = imageBytes,
        //                FileName = $"pdf_page{pageNum}_image_{imageCounter}.png",
        //                Position = 0,
        //                ContentType = "image/png"
        //            });
        //        }
        //        catch (Exception ex)
        //        {
        //            Console.WriteLine($"  Warning: Could not extract image from PDF page {pageNum} - {ex.Message}");
        //        }
        //    }

        //    return images;
        //}

        //static string GetImageExtension(string contentType)
        //{
        //    return contentType.ToLower() switch
        //    {
        //        "image/jpeg" => ".jpg",
        //        "image/jpg" => ".jpg",
        //        "image/png" => ".png",
        //        "image/gif" => ".gif",
        //        "image/bmp" => ".bmp",
        //        "image/tiff" => ".tiff",
        //        _ => ".png"
        //    };
        //}

        //static string ConvertTableToMarkdown(Table table)
        //{
        //    var sb = new StringBuilder();
        //    var rows = table.Elements<TableRow>().ToList();

        //    if (rows.Count == 0) return "";

        //    // Header row
        //    var headerCells = rows[0].Elements<TableCell>().Select(c => c.InnerText).ToList();
        //    sb.AppendLine("| " + string.Join(" | ", headerCells) + " |");
        //    sb.AppendLine("| " + string.Join(" | ", headerCells.Select(c => "---")) + " |");

        //    // Data rows
        //    for (int i = 1; i < rows.Count; i++)
        //    {
        //        var cells = rows[i].Elements<TableCell>().Select(c => c.InnerText).ToList();
        //        sb.AppendLine("| " + string.Join(" | ", cells) + " |");
        //    }

        //    sb.AppendLine();
        //    return sb.ToString();
        //}

        //static async Task CreateClickUpPageWithImages(string title, string content, List<ImageInfo> images)
        //{
        //    // Step 1: Upload all images to ClickUp and get their URLs
        //    var imageUrls = new Dictionary<int, string>();

        //    for (int i = 0; i < images.Count; i++)
        //    {
        //        var image = images[i];
        //        try
        //        {
        //            string imageUrl = await UploadImageToClickUp(image);
        //            imageUrls[i + 1] = imageUrl; // 1-indexed for placeholders
        //            Console.WriteLine($"  ✓ Uploaded image {i + 1}/{images.Count}");
        //        }
        //        catch (Exception ex)
        //        {
        //            Console.WriteLine($"  ✗ Failed to upload image {i + 1}: {ex.Message}");
        //        }
        //    }

        //    // Step 2: Replace placeholders with actual image URLs
        //    string finalContent = content;
        //    foreach (var kvp in imageUrls)
        //    {
        //        string placeholder = $"(IMAGE_PLACEHOLDER_{kvp.Key})";
        //        string imageMarkdown = $"({kvp.Value})";
        //        finalContent = finalContent.Replace(placeholder, imageMarkdown);
        //    }

        //    // Step 3: Create the page with images embedded
        //    await CreateClickUpPage(title, finalContent);
        //}

        //static async Task<string> UploadImageToClickUp(ImageInfo image)
        //{
        //    using (var formData = new MultipartFormDataContent())
        //    {
        //        var imageContent = new ByteArrayContent(image.ImageData);
        //        imageContent.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue(image.ContentType);

        //        formData.Add(imageContent, "attachment", image.FileName);

        //        // Upload to team/workspace (adjust endpoint based on your needs)
        //        // You may need to use task attachment or a different endpoint
        //        var response = await clickupClient.PostAsync(
        //            $"https://api.clickup.com/api/v2/space/{SPACE_ID}/attachment",
        //            formData
        //        );

        //        if (!response.IsSuccessStatusCode)
        //        {
        //            string error = await response.Content.ReadAsStringAsync();
        //            throw new Exception($"Failed to upload image: {response.StatusCode} - {error}");
        //        }

        //        string responseJson = await response.Content.ReadAsStringAsync();
        //        var doc = JsonDocument.Parse(responseJson);

        //        // Extract the URL from response
        //        if (doc.RootElement.TryGetProperty("url", out JsonElement urlElement))
        //        {
        //            return urlElement.GetString();
        //        }

        //        throw new Exception("Could not extract image URL from response");
        //    }
        //}

        //static async Task CreateClickUpPage(string title, string content)
        //{
        //    var payload = new
        //    {
        //        name = title,
        //        content = content,
        //        content_type = "markdown",
        //        parent_page_id = string.IsNullOrEmpty(PARENT_PAGE_ID) ? null : PARENT_PAGE_ID
        //    };

        //    string json = JsonSerializer.Serialize(payload);
        //    var httpContent = new StringContent(json, Encoding.UTF8, "application/json");

        //    var response = await clickupClient.PostAsync(
        //        $"https://api.clickup.com/api/v2/space/{SPACE_ID}/page",
        //        httpContent
        //    );

        //    if (!response.IsSuccessStatusCode)
        //    {
        //        string error = await response.Content.ReadAsStringAsync();
        //        throw new Exception($"ClickUp API error: {response.StatusCode} - {error}");
        //    }
        //}

        static async Task<List<SelectionItem>> ListPagesInSpace()
        {
            //Console.WriteLine("Fetching pages in Wiki...\n");
            ConsoleHelper.LogInformation("Fetching pages in Wiki...\n");

            var response = await clickupClient.GetAsync(
                $"https://api.clickup.com/api/v3/workspaces/{WORKSPACE_ID}/docs/{WIKI_ID}/pages"
            );

            if (response.IsSuccessStatusCode)
            {
                string json = await response.Content.ReadAsStringAsync();
                // Extract pages maintaining hierarchy
                var pages = PageExtractor.ExtractPages(json);
                allPages = pages;

                //// Print hierarchical structure
                //Console.WriteLine("Hierarchical Structure:");
                //PageExtractor.PrintHierarchy(pages);

                var selectionItems = new List<SelectionItem>();
                PageExtractor.ExtractPageHierarchy(pages, selectionItems);

                //Console.WriteLine("\n---\n");

                //// Or flatten if you need a simple list
                //var flatPages = PageExtractor.FlattenPages(pages);
                //Console.WriteLine("Flattened List:");
                //foreach (var page in flatPages)
                //{
                //    Console.WriteLine($"- {page.Name} (ID: {page.Id}, Parent: {page.ParentPageId ?? "null"})");
                //}

                //Console.WriteLine();

                return selectionItems;
            }

            return [];
        }
    }
}
