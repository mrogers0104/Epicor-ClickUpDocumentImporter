using ClickUpDocumentImporter.Helpers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
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
            // *** Use the first image found for each RelationshipId to avoid duplicates ***
            var imageLookup = images
                .Where(i => !string.IsNullOrEmpty(i.RelationshipId)) // Optional: filter nulls
                .GroupBy(i => i.RelationshipId)
                .ToDictionary(g => g.Key, g => g.First());

            ConsoleHelper.WriteInfo($"Found {images.Count} images in Word document");

            ConsoleHelper.WriteSeparator();
            ConsoleHelper.WriteInfo($"~~~~~ Document: {Path.GetFileName(wordFilePath)} ~~~~~");
            ConsoleHelper.WriteInfo($"Found {images.Count} images in Word document");

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(wordFilePath, false))
            {
                var formatter = new WordToMarkdownFormatter(wordDoc);
                var mainPart = wordDoc.MainDocumentPart;

                // Track images already processed to avoid double-counting (crucial!)
                var processedRIds = new HashSet<string>();

                // 2. 🔥 IDENTIFY ALL PARTS TO SEARCH (Body, Headers, Footers)
                var contentContainers = new List<OpenXmlPart>() { mainPart };

                // Add all existing HeaderParts and FooterParts
                // GetPartsOfType is safer than iterating over properties which might be null.
                contentContainers.AddRange(mainPart.GetPartsOfType<HeaderPart>().Cast<OpenXmlPart>());
                contentContainers.AddRange(mainPart.GetPartsOfType<FooterPart>().Cast<OpenXmlPart>());

                // The contentRoot is the Body, Header, or Footer element.
                // We iterate over its direct children to maintain the document flow.
                foreach (var container in contentContainers) //   OpenXmlElement element in contentContainers) // contentRoot.Elements())
                {
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
    }
}