using ClickUpDocumentImporter.Helpers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Pdf.Xobject;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;
using A = DocumentFormat.OpenXml.Drawing;


namespace ClickUpDocumentImporter.DocumentConverter
{
    /// <summary>
    /// COMPLETE EXAMPLE: Extract and Upload to ClickUp
    /// </summary>
    /// <remarks>
    /// By using the Relationship ID mapping and iterating over all relevant content parts (Body,
    /// Headers, and Footers), we ensure that all images are correctly identified, located, and
    /// processed without relying on unreliable sequential indexing.
    /// </remarks>
    public class CompleteDocumentConverter
    {
        public static async Task ConvertWordToClickUpAsync(
            string wordFilePath,
            HttpClient clickupClient,
            string workspaceId,
            string wikiId,
            string listId,
            string parentPageId = null
            )
        {
            var builder = new ClickUpDocumentBuilder(clickupClient);

            // Extract images first
            var images = WordImageExtractor.ExtractImagesFromWord(wordFilePath);

            // 1. 🔥 CREATE LOOKUP MAP: Create a Dictionary for fast lookup by RelationshipId
            var imageLookup = images.ToDictionary(i => i.RelationshipId, i => i);

            ConsoleHelper.WriteInfo($"Found {images.Count} images in Word document");

            ConsoleHelper.WriteSeparator();
            ConsoleHelper.WriteInfo($"~~~~~ Document: {Path.GetFileName(wordFilePath)} ~~~~~");
            ConsoleHelper.WriteInfo($"Found {images.Count} images in Word document");

            //// Build content with text and images
            //using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(wordFilePath, false))
            //{
            //    var formatter = new WordToMarkdownFormatter(wordDoc);

            //    var body = wordDoc.MainDocumentPart.Document.Body;
            //    int currentImageIndex = 0;

            //    foreach (var element in body.Elements())
            //    {

            //        if (element is Paragraph para)
            //        {
            //            formatter.ProcessParagraph(para, builder);

            //            // Check if paragraph contains an image
            //            var drawing = para.Descendants<Drawing>().FirstOrDefault();

            //            if (drawing != null)
            //            {
            //                // This paragraph contains an image
            //                if (currentImageIndex < images.Count)
            //                {
            //                    var imageData = images[currentImageIndex];
            //                    await builder.AddImage(
            //                        imageData.Data,
            //                        imageData.FileName,
            //                        listId
            //                    );

            //                    Console.WriteLine($"Added image #{currentImageIndex}: {imageData.FileName}");
            //                    currentImageIndex++;
            //                }
            //            }
            //        }
            //        else if (element is Table table)
            //        {
            //            formatter.ProcessTable(table, builder);
            //        }
            //        // Handle images separately as before
            //        else if (element is Drawing drawing)
            //        {
            //            // This paragraph contains an image
            //            if (currentImageIndex < images.Count)
            //            {
            //                var imageData = images[currentImageIndex];
            //                await builder.AddImage(
            //                    imageData.Data,
            //                    imageData.FileName,
            //                    listId
            //                );

            //                ConsoleHelper.WriteInfo($"Added image: {imageData.FileName}");
            //                currentImageIndex++;
            //            }
            //        } else
            //        {
            //            continue;
            //        }
            //    }

            //    formatter.ResetListCounters();
            //}

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(wordFilePath, false))
            {
                var formatter = new WordToMarkdownFormatter(wordDoc);
                var mainPart = wordDoc.MainDocumentPart;

                // Track images already processed to avoid double-counting (crucial!)
                var processedRIds = new HashSet<string>();

                // 2. 🔥 IDENTIFY ALL PARTS TO SEARCH (Body, Headers, Footers)
                var contentContainers = new List<OpenXmlPart>() { mainPart };
                //contentContainers.AddRange(mainPart.HeaderParts.Cast<OpenXmlPart>());
                //contentContainers.AddRange(mainPart.FooterParts.Cast<OpenXmlPart>());

                // Add all existing HeaderParts and FooterParts
                // GetPartsOfType is safer than iterating over properties which might be null.
                contentContainers.AddRange(mainPart.GetPartsOfType<HeaderPart>().Cast<OpenXmlPart>());
                contentContainers.AddRange(mainPart.GetPartsOfType<FooterPart>().Cast<OpenXmlPart>());

                // The contentRoot is the Body, Header, or Footer element.
                // We iterate over its direct children to maintain the document flow.
                foreach (var container in contentContainers) //   OpenXmlElement element in contentContainers) // contentRoot.Elements())
                {
                    //var contentRoot = container.RootElement;
                    //if (contentRoot == null) continue;

                    OpenXmlElement contentToProcess = null;

                    if (container is MainDocumentPart mPart)
                    {
                        // 🔥 Get the Body element, which holds the content.
                        contentToProcess = mPart.Document.Body;
                    }
                    else
                    {
                        // For Header/Footer parts, the RootElement IS the correct content container.
                        contentToProcess = container.RootElement;
                    }

                    if (contentToProcess == null) continue;

                    await ProcessContentElementsAsync(contentToProcess,
                                                      formatter,
                                                      builder,
                                                      imageLookup,
                                                      processedRIds,
                                                      listId);

                    //    if (element is Paragraph para)
                    //    {
                    //        // 1. Process the sequential paragraph text
                    //        formatter.ProcessParagraph(para, builder);

                    //        // 2. Process all *nested* drawings within this single paragraph
                    //        foreach (var drawing in para.Descendants<Drawing>())
                    //        {
                    //            await ProcessDrawingElementAsync(drawing, imageLookup, processedRIds, builder, listId);
                    //        }
                    //    }
                    //    else if (element is Table table)
                    //    {
                    //        // 1. Process the sequential table content
                    //        formatter.ProcessTable(table, builder);

                    //        // 2. Process all *nested* drawings within the table
                    //        foreach (var drawingInTable in table.Descendants<Drawing>())
                    //        {
                    //            await ProcessDrawingElementAsync(drawingInTable, imageLookup, processedRIds, builder, listId);
                    //        }
                    //    }
                    //    // Handle other block-level elements like SdtBlock if necessary
                    //    else if (element is SdtBlock sdtBlock)
                    //    {
                    //        // Recursively process content inside the content control
                    //        // Find Paragraphs and Tables inside the SdtBlock
                    //        foreach (var blockElement in sdtBlock.Descendants<OpenXmlElement>())
                    //        {
                    //            if (blockElement is Paragraph nestedPara)
                    //            {
                    //                // Process content and images recursively
                    //                formatter.ProcessParagraph(nestedPara, builder);
                    //                foreach (var drawing in nestedPara.Descendants<Drawing>())
                    //                {
                    //                    await ProcessDrawingElementAsync(drawing, imageLookup, processedRIds, builder, listId);
                    //                }
                    //            }
                    //            else if (blockElement is Table nestedTable)
                    //            {
                    //                // Process content and images recursively
                    //                formatter.ProcessTable(nestedTable, builder);
                    //                foreach (var drawingInTable in nestedTable.Descendants<Drawing>())
                    //                {
                    //                    await ProcessDrawingElementAsync(drawingInTable, imageLookup, processedRIds, builder, listId);
                    //                }
                    //            }
                    //        }
                    //    }
                    //    // Ignore other elements that don't produce output (like SectionProperties, etc.)
                    //}

                    //// 2. ITERATE OVER EACH CONTENT PART
                    //foreach (var container in contentContainers)
                    //{
                    //    // The RootElement gives you the content (Body, Header, or Footer element)
                    //    var contentRoot = container.RootElement;
                    //    if (contentRoot == null) continue;

                    //    // 3. PROCESS PARAGRAPHS AND TABLES WITHIN THIS PART (RECURSIVELY)
                    //    ProcessContentElements(contentRoot, formatter, builder, imageLookup, processedRIds, listId);
                    //}

                    //// 3. ITERATE OVER ALL CONTENT CONTAINERS (Body, then Headers, then Footers)
                    //foreach (var container in contentContainers)
                    //{
                    //    // Get the content element (Document.Body for MainPart, Header for HeaderPart, etc.)
                    //    var contentRoot = container.RootElement;
                    //    if (contentRoot == null) continue;

                    //    // Iterate over all elements in this container (Paragraphs, Tables, etc.)
                    //    foreach (var element in contentRoot.Elements())
                    //    {
                    //        if (element is Paragraph para)
                    //        {
                    //            // Process Paragraph text content first
                    //            formatter.ProcessParagraph(para, builder);

                    //            // CHECK FOR IMAGES WITHIN THE PARAGRAPH (Inline or Anchor)
                    //            var drawing = para.Descendants<Drawing>().FirstOrDefault();
                    //            if (drawing != null)
                    //            {
                    //                await ProcessDrawingElement(drawing, imageLookup, processedRIds, builder, listId);
                    //            }
                    //        }
                    //        else if (element is Table table)
                    //        {
                    //            // Process Table text/structure
                    //            formatter.ProcessTable(table, builder);

                    //            // You must also search for images *inside* the table cells
                    //            foreach (var drawingInTable in table.Descendants<Drawing>())
                    //            {
                    //                await ProcessDrawingElement(drawingInTable, imageLookup, processedRIds, builder, listId);
                    //            }
                    //        }
                    //        // No need for 'else if (element is Drawing drawing)' anymore,
                    //        // as drawings are usually descendants of Paragraph or Table.
                    //    }
                    //}
                }
            }

                // Create the ClickUp page
                string pageName = Path.GetFileNameWithoutExtension(wordFilePath);
            string pageId = await builder.CreateAndPopulatePageAsync(
                workspaceId,
                wikiId,
                pageName,
                parentPageId,
                uploadMethod: "task",
                listIdForTaskUpload: listId
            );

            ConsoleHelper.WriteInfo($"\n✓ Created ClickUp page: {pageName} (ID: {pageId})");
        }

        //private async static void ProcessContentElements(
        //    OpenXmlElement contentRoot,
        //    WordToMarkdownFormatter formatter,
        //    ClickUpDocumentBuilder builder,
        //    Dictionary<string, ImageData> imageLookup,
        //    HashSet<string> processedRIds,
        //    string listId)
        //{
        //    // A. Process all Paragraphs in this part
        //    foreach (var para in contentRoot.Descendants<Paragraph>())
        //    {
        //        // 1. Process Text Content (Applies only to Body content, but safe to run)
        //        // You might need a more sophisticated check here if you want to skip
        //        // processing the text from headers/footers for the main document content.
        //        formatter.ProcessParagraph(para, builder);

        //        // 2. Check for image drawings inside the paragraph
        //        //var drawing = para.Descendants<Drawing>().FirstOrDefault();
        //        //if (drawing != null)
        //        //{
        //        //    // Process Drawing (finds relationship ID, uploads image)
        //        //    // Note: Since this is synchronous, image uploading must be handled inside
        //        //    // the main async method, or this helper needs to be async.
        //        //    // Assuming ProcessDrawingElement is async based on your earlier request:
        //        //    ProcessDrawingElementAsync(drawing, imageLookup, processedRIds, builder, listId).Wait();
        //        //}

        //        // 🔥 FIX: Iterate over ALL drawings within the paragraph (or any descendant)
        //        foreach (var drawing in para.Descendants<Drawing>())
        //        {
        //            // Process Drawing (finds relationship ID, uploads image)
        //            await ProcessDrawingElementAsync(drawing, imageLookup, processedRIds, builder, listId);
        //        }
        //    }

        //    // B. Process all Tables in this part
        //    foreach (var table in contentRoot.Descendants<Table>())
        //    {
        //        formatter.ProcessTable(table, builder);

        //        // Check for drawings *inside* the table cells
        //        foreach (var drawingInTable in table.Descendants<Drawing>())
        //        {
        //            ProcessDrawingElementAsync(drawingInTable, imageLookup, processedRIds, builder, listId).Wait();
        //        }
        //    }
        //}

        private static async Task ProcessContentElementsAsync(OpenXmlElement contentRoot,
                                                              WordToMarkdownFormatter formatter,
                                                              ClickUpDocumentBuilder builder,
                                                              Dictionary<string, ImageData> imageLookup,
                                                              HashSet<string> processedRIds,
                                                              string listId)
        {
            // Iterate over the direct children to maintain correct document flow/sequence
            foreach (OpenXmlElement element in contentRoot.Elements())
            {
                // --- Case 1: Paragraph (Most common text container) ---
                if (element is Paragraph para)
                {
                    formatter.ProcessParagraph(para, builder);

                    // Use Descendants to find ALL nested images within this paragraph
                    foreach (var drawing in para.Descendants<Drawing>())
                    {
                        await ProcessDrawingElementAsync(drawing, imageLookup, processedRIds, builder, listId);
                    }
                }
                // --- Case 2: Table ---
                else if (element is Table table)
                {
                    formatter.ProcessTable(table, builder);

                    // Use Descendants to find ALL nested images within all cells of this table
                    foreach (var drawingInTable in table.Descendants<Drawing>())
                    {
                        await ProcessDrawingElementAsync(drawingInTable, imageLookup, processedRIds, builder, listId);
                    }
                }
                // --- Case 3: SdtBlock (Content Control Wrapper) ---
                else if (element is SdtBlock sdtBlock)
                {
                    // Recursively process the content inside the SdtBlock's Content element
                    var content = sdtBlock.Descendants<SdtContentBlock>().FirstOrDefault();
                    if (content != null)
                    {
                        // Call this method recursively on the content block
                        await ProcessContentElementsAsync(content, formatter, builder, imageLookup, processedRIds, listId);
                    }
                }
                // Other top-level elements (like SectionProperties) are implicitly skipped or handled by the formatter.
            }
        }

        // NOTE: You must update your main loop to call the async version:
        // await ProcessContentElementsAsync(contentRoot, ...);

        // ⚠️ Ensure your image processing logic is made synchronous or called asynchronously
        // if you are keeping your main conversion method async. (Used .Wait() for simplicity here.)
        private static async Task ProcessDrawingElementAsync(
            Drawing drawing,
            Dictionary<string, ImageData> imageLookup,
            HashSet<string> processedRIds,
            ClickUpDocumentBuilder builder,
            string listId)
        {
            var blip = drawing.Descendants<A.Blip>().FirstOrDefault();
            if (blip?.Embed != null)
            {
                string relationshipId = blip.Embed.Value;

                // Use the HashSet to ensure we only process the image once, even if referenced multiple times
                if (imageLookup.TryGetValue(relationshipId, out var imageData) &&
                    !processedRIds.Contains(relationshipId))
                {
                    await builder.AddImage(imageData.Data, imageData.FileName, listId);
                    processedRIds.Add(relationshipId);
                }
            }
        }

        ///// <summary>
        ///// helper method uses the Relationship ID to safely look up and process the image, preventing duplicates.
        ///// </summary>
        ///// <param name="drawing"></param>
        ///// <param name="imageLookup"></param>
        ///// <param name="processedRIds"></param>
        ///// <param name="builder"></param>
        ///// <param name="listId"></param>
        ///// <returns></returns>
        //private static async Task ProcessDrawingElement(Drawing drawing,
        //                                                Dictionary<string, ImageData> imageLookup,
        //                                                HashSet<string> processedRIds,
        //                                                ClickUpDocumentBuilder builder,
        //                                                string listId)
        //{
        //    // The Blip element holds the relationship ID (Embed attribute)
        //    var blip = drawing.Descendants<A.Blip>().FirstOrDefault();

        //    if (blip?.Embed != null)
        //    {
        //        string relationshipId = blip.Embed.Value;

        //        if (imageLookup.TryGetValue(relationshipId, out var imageData) &&
        //            !processedRIds.Contains(relationshipId))
        //        {
        //            // Found the image and haven't processed it yet!
        //            await builder.AddImage(
        //                imageData.Data,
        //                imageData.FileName,
        //                listId
        //            );

        //            ConsoleHelper.WriteInfo($"Added image (RID: {relationshipId}): {imageData.FileName}");

        //            // Mark as processed
        //            processedRIds.Add(relationshipId);
        //        }
        //    }
        //}
    }
}
