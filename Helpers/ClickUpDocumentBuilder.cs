using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.Intrinsics.X86;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading.Tasks;

namespace ClickUpDocumentImporter.Helpers
{

    // Helper classes for document structure
    public enum ElementType { Heading, Paragraph, Image }

    public class DocumentElement
    {
        public ElementType Type { get; set; }
        public string Text { get; set; }
        public int Level { get; set; }
        public byte[] ImageData { get; set; }
        public string ImageName { get; set; }
    }

    /// <summary>
    /// ClickUpDocumentBuilder helps build and update ClickUp pages with mixed content (text,
    /// headings, images).
    ///
    /// Key Point: Sequential Building: The ClickUpDocumentBuilder maintains a list of content
    /// blocks that you add in the order they appear in your source document.
    ///
    /// Two Approaches:
    ///
    /// 1. Sequential(simpler) : As you parse each element(text, image, heading), immediately upload
    ///    images and add blocks in order
    /// 2. Batch Upload(faster) : First upload all images, then build the content structure with the URLs
    ///
    /// The Critical Part: When you parse your Word/PDF document, you must:
    ///
    /// Iterate through elements in document order
    /// * Call AddParagraph(), AddImage(), AddHeading() etc. in the same sequence as they appear
    /// * Only call UpdatePageAsync() once at the end with the complete, ordered content
    ///
    /// Example Flow:
    /// Document: [Title] → [Paragraph] → [Image] → [Paragraph] → [Image]
    /// ClickUp:  Block 0 → Block 1 → Block 2 → Block 3 → Block 4
    ///
    /// This ensures images appear exactly where they were in the original document.
    /// </summary>
    /// <remarks>
    /// Creating a New Page
    ///         var builder = new ClickUpDocumentBuilder(apiToken);
    ///
    ///         // Add all your content
    ///         builder.AddHeading("My Document", 1);
    ///         builder.AddParagraph("Some text...");
    ///         await builder.AddImageAsync(imageData, "image.png", workspaceId);
    ///
    ///             // Create new page with content (optionally under a parent page)
    ///             string newPageId = await builder.CreateAndPopulatePageAsync(
    ///                 workspaceId,
    ///                 "My Document Name",
    ///                 parentPageId  // optional - null for top-level page
    ///             );
    /// Updating an Existing Page
    ///     var builder = new ClickUpDocumentBuilder(apiToken);
    ///
    ///         // Add all your content
    ///         builder.AddHeading("My Document", 1);
    ///     builder.AddParagraph("Some text...");
    ///     await builder.AddImageAsync(imageData, "image.png", workspaceId);
    ///
    ///         // Update existing page
    ///         bool success = await builder.UpdatePageAsync(existingPageId);
    /// Key Parameters:
    ///
    /// * workspaceId: Required for creating pages and uploading images
    /// * parentPageId: Optional - use this to create the page as a child of another page (for organizing documents hierarchically)
    /// * pageName: The title/name of the new page
    ///
    /// The content is populated during creation, so images will be in the correct positions right from the start!
    /// </remarks>
    public class ClickUpDocumentBuilder
    {
        private readonly HttpClient _httpClient;
        private readonly string _apiToken;
        private readonly List<JsonObject> _contentBlocks;

        public ClickUpDocumentBuilder(string apiToken)
        {
            _apiToken = apiToken;
            _httpClient = new HttpClient();
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _apiToken);
            _contentBlocks = new List<JsonObject>();
        }

        // Upload image and return URL
        public async Task<string> UploadImageAsync(byte[] imageData, string fileName, string workspaceId)
        {
            var content = new MultipartFormDataContent();
            var fileContent = new ByteArrayContent(imageData);
            fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse("image/png");
            content.Add(fileContent, "attachment", fileName);

            var response = await _httpClient.PostAsync(
                $"https://api.clickup.com/api/v2/workspace/{workspaceId}/attachment",
                content
            );

            response.EnsureSuccessStatusCode();
            var responseJson = await response.Content.ReadAsStringAsync();
            var result = JsonSerializer.Deserialize<JsonNode>(responseJson);

            return result["url"].GetValue<string>();
        }

        // Add a paragraph block
        public void AddParagraph(string text)
        {
            var block = new JsonObject
            {
                ["type"] = "paragraph",
                ["content"] = new JsonArray
            {
                new JsonObject
                {
                    ["type"] = "text",
                    ["text"] = text
                }
            }
            };
            _contentBlocks.Add(block);
        }

        // Add a heading block
        public void AddHeading(string text, int level = 1)
        {
            var block = new JsonObject
            {
                ["type"] = "heading",
                ["attrs"] = new JsonObject { ["level"] = level },
                ["content"] = new JsonArray
            {
                new JsonObject
                {
                    ["type"] = "text",
                    ["text"] = text
                }
            }
            };
            _contentBlocks.Add(block);
        }

        // Add an image block at current position
        public async Task AddImageAsync(byte[] imageData, string fileName, string workspaceId)
        {
            // Upload the image first
            var imageUrl = await UploadImageAsync(imageData, fileName, workspaceId);

            // Create image block
            var block = new JsonObject
            {
                ["type"] = "image",
                ["attrs"] = new JsonObject
                {
                    ["src"] = imageUrl,
                    ["alt"] = fileName
                }
            };

            _contentBlocks.Add(block);
            Console.WriteLine($"Image added at position {_contentBlocks.Count - 1}: {fileName}");
        }

        // Add a placeholder for an image (to be uploaded later)
        public void AddImagePlaceholder(int imageIndex)
        {
            var block = new JsonObject
            {
                ["type"] = "image_placeholder",
                ["attrs"] = new JsonObject
                {
                    ["imageIndex"] = imageIndex
                }
            };
            _contentBlocks.Add(block);
        }

        // Replace placeholder with actual uploaded image
        public async Task ReplaceImagePlaceholderAsync(int placeholderIndex, byte[] imageData,
            string fileName, string workspaceId)
        {
            var imageUrl = await UploadImageAsync(imageData, fileName, workspaceId);

            _contentBlocks[placeholderIndex] = new JsonObject
            {
                ["type"] = "image",
                ["attrs"] = new JsonObject
                {
                    ["src"] = imageUrl,
                    ["alt"] = fileName
                }
            };
        }

        // Create a new ClickUp page
        public async Task<string> CreatePageAsync(string workspaceId, string pageName, string parentPageId = null)
        {
            var contentArray = new JsonArray();
            foreach (var block in _contentBlocks)
            {
                contentArray.Add(block);
            }

            var createPayload = new JsonObject
            {
                ["name"] = pageName,
                ["content"] = contentArray
            };

            // If there's a parent page, add it to the payload
            if (!string.IsNullOrEmpty(parentPageId))
            {
                createPayload["parent_page_id"] = parentPageId;
            }

            var jsonString = createPayload.ToJsonString();
            var content = new StringContent(
                jsonString,
                Encoding.UTF8,
                "application/json"
            );

            var response = await _httpClient.PostAsync(
                $"https://api.clickup.com/api/v3/workspaces/{workspaceId}/pages",
                content
            );

            response.EnsureSuccessStatusCode();
            var responseJson = await response.Content.ReadAsStringAsync();
            var result = JsonSerializer.Deserialize<JsonNode>(responseJson);

            var pageId = result["id"].GetValue<string>();
            Console.WriteLine($"Created new page: {pageName} (ID: {pageId})");

            return pageId;
        }

        // Update an existing ClickUp page with all content blocks
        public async Task<bool> UpdatePageAsync(string pageId)
        {
            var contentArray = new JsonArray();
            foreach (var block in _contentBlocks)
            {
                contentArray.Add(block);
            }

            var updatePayload = new JsonObject
            {
                ["content"] = contentArray
            };

            var jsonString = updatePayload.ToJsonString();
            var content = new StringContent(
                jsonString,
                Encoding.UTF8,
                "application/json"
            );

            var response = await _httpClient.PutAsync(
                $"https://api.clickup.com/api/v3/page/{pageId}",
                content
            );

            return response.IsSuccessStatusCode;
        }

        // Create a new page and populate it with content in one call
        public async Task<string> CreateAndPopulatePageAsync(string workspaceId, string pageName, string parentPageId = null)
        {
            // Content is already added via AddParagraph, AddImage, etc.
            return await CreatePageAsync(workspaceId, pageName, parentPageId);
        }

        // Clear all content blocks
        public void Clear()
        {
            _contentBlocks.Clear();
        }

        // Make content blocks accessible for batch upload scenario
        public List<JsonObject> ContentBlocks => _contentBlocks;
    }

    //// Example: Converting a document with mixed content
    //class DocumentConverter
    //{
    //    // Create a NEW page with content
    //    public static async Task<string> ConvertDocumentToNewPageAsync(
    //        string apiToken,
    //        string workspaceId,
    //        string pageName,
    //        string parentPageId = null)
    //    {
    //        var builder = new ClickUpDocumentBuilder(apiToken);

    //        // Build content blocks in order
    //        builder.AddHeading("Document Title", 1);

    //        builder.AddParagraph("This is the first paragraph of the document.");

    //        // Image appears here in the original document
    //        byte[] image1Data = /* extracted image bytes */;
    //        await builder.AddImageAsync(image1Data, "diagram1.png", workspaceId);

    //        builder.AddParagraph("This paragraph comes after the first image.");

    //        builder.AddHeading("Section 2", 2);

    //        builder.AddParagraph("Another paragraph with more content.");

    //        // Another image in sequence
    //        byte[] image2Data = /* extracted image bytes */;
    //        await builder.AddImageAsync(image2Data, "chart1.png", workspaceId);

    //        builder.AddParagraph("Final paragraph after the second image.");

    //        // Create the new page with all content
    //        var newPageId = await builder.CreateAndPopulatePageAsync(workspaceId, pageName, parentPageId);
    //        Console.WriteLine($"Document conversion completed. New page ID: {newPageId}");

    //        return newPageId;
    //    }

    //    // Update an EXISTING page with content
    //    public static async Task ConvertDocumentToExistingPageAsync(
    //        string apiToken,
    //        string workspaceId,
    //        string pageId)
    //    {
    //        var builder = new ClickUpDocumentBuilder(apiToken);

    //        // Simulate parsing your Word/PDF document in order
    //        builder.AddHeading("Document Title", 1);

    //        builder.AddParagraph("This is the first paragraph of the document.");

    //        byte[] image1Data = /* extracted image bytes */;
    //        await builder.AddImageAsync(image1Data, "diagram1.png", workspaceId);

    //        builder.AddParagraph("This paragraph comes after the first image.");

    //        // Update the existing ClickUp page with all content in correct order
    //        var success = await builder.UpdatePageAsync(pageId);
    //        Console.WriteLine($"Document conversion completed: {success}");
    //    }

    //    // Alternative approach: Batch upload images first, then build content
    //    public static async Task ConvertWithBatchUploadAsync(
    //        string apiToken,
    //        string workspaceId,
    //        string pageId,
    //        List<DocumentElement> elements) // Your parsed document elements
    //    {
    //        var builder = new ClickUpDocumentBuilder(apiToken);

    //        // First pass: Upload all images and store URLs
    //        var imageUrls = new Dictionary<int, string>();

    //        for (int i = 0; i < elements.Count; i++)
    //        {
    //            if (elements[i].Type == ElementType.Image)
    //            {
    //                var url = await builder.UploadImageAsync(
    //                    elements[i].ImageData,
    //                    elements[i].ImageName,
    //                    workspaceId
    //                );
    //                imageUrls[i] = url;
    //                Console.WriteLine($"Uploaded image {i}: {elements[i].ImageName}");
    //            }
    //        }

    //        // Second pass: Build content in correct order with uploaded image URLs
    //        for (int i = 0; i < elements.Count; i++)
    //        {
    //            switch (elements[i].Type)
    //            {
    //                case ElementType.Heading:
    //                    builder.AddHeading(elements[i].Text, elements[i].Level);
    //                    break;
    //                case ElementType.Paragraph:
    //                    builder.AddParagraph(elements[i].Text);
    //                    break;
    //                case ElementType.Image:
    //                    // Add image block with pre-uploaded URL
    //                    builder.ContentBlocks.Add(new JsonObject
    //                    {
    //                        ["type"] = "image",
    //                        ["attrs"] = new JsonObject
    //                        {
    //                            ["src"] = imageUrls[i],
    //                            ["alt"] = elements[i].ImageName
    //                        }
    //                    });
    //                    break;
    //            }
    //        }

    //        // Update page once with all content
    //        var success = await builder.UpdatePageAsync(pageId);
    //        Console.WriteLine($"Batch conversion completed: {success}");
    //    }
    //}
}
