using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using DocumentFormat.OpenXml.Office2019.Presentation;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using HashidsNet;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Mail;
using System.Runtime.Intrinsics.X86;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading.Tasks;

namespace ClickUpDocumentImporter.Helpers
{
    // Helper classes for document structure
    public enum ElementType
    { Heading, Paragraph, Image }

    public class DocumentElement
    {
        public ElementType Type { get; set; }
        public string Text { get; set; }
        public int Level { get; set; }
        public byte[] ImageData { get; set; }
        public string ImageName { get; set; }
    }

    // Internal class to track content blocks with pending images
    internal class ContentBlock
    {
        public JsonObject JsonBlock { get; set; }
        public byte[] ImageData { get; set; }
        public string FileName { get; set; }
        public bool IsImageBlock { get; set; }
    }

    // Internal class to track pending images
    internal class ImageBlock
    {
        public byte[] ImageData { get; set; }
        public string FileName { get; set; }
        public string Placeholder { get; set; }
    }

    /// <summary>
    /// ClickUpDocumentBuilder helps build and update ClickUp pages with mixed content (text,
    /// headings, images).
    ///
    /// Key Fetures:
    /// 1. Correct Flow:
    ///
    /// Create the page first with text content
    /// Upload images to the page(not workspace) using /api/v3/page/{pageId
    /// }/ attachment
    /// Update the page with the complete content including image URLs
    ///
    /// 2. Internal Storage:
    ///
    /// Uses a ContentBlock class to track pending image data
    /// Stores byte[] image data until the page is created
    ///
    /// 3. Single Method:
    ///
    /// CreateAndPopulatePageAsync() now handles the entire workflow:
    ///
    /// Creates page with text + placeholders
    /// Uploads all images to the created page
    /// Updates page with final content
    ///
    /// 4. Usage Example:
    /// C#
    ///     var builder = new ClickUpDocumentBuilder(apiToken);
    ///
    ///     // Add content in document order
    ///     builder.AddHeading("My Document Title", 1);
    ///     builder.AddParagraph("Introduction text here...");
    ///     builder.AddImage(imageBytes1, "diagram1.png");
    ///     builder.AddParagraph("More text after the image...");
    ///     builder.AddImage(imageBytes2, "screenshot.png");
    ///
    ///     // Create and populate in one call
    ///     string pageId = await builder.CreateAndPopulatePageAsync(
    ///         workspaceId: "your-workspace-id",
    ///         pageName: "Imported Document",
    ///         parentPageId: "parent-page-id"
    ///     );
    ///
    /// The corrected flow ensures images are uploaded to the correct page context and maintains the document's content order.
    /// </remarks>
    public class ClickUpDocumentBuilder
    {
        private readonly HttpClient _httpClient;

        //private readonly string _apiToken;
        private readonly StringBuilder _markdownContent;

        private readonly List<ImageBlock> _pendingImages;
        //private List<object> contentBlocks; // This should be defined at class level

        public ClickUpDocumentBuilder(HttpClient httpClient)
        {
            _httpClient = httpClient;
            _markdownContent = new StringBuilder();
            _pendingImages = new List<ImageBlock>();
            //contentBlocks = new List<object>(); // Initialize in constructor
        }

        // Add a paragraph
        public void AddParagraph(string text)
        {
            _markdownContent.AppendLine(text);
            _markdownContent.AppendLine(); // Add blank line after paragraph
        }

        // Add a heading
        public void AddHeading(string text, int level = 1)
        {
            string prefix = new string('#', level);
            _markdownContent.AppendLine($"{prefix} {text}");
            _markdownContent.AppendLine();
        }

        public void AddMarkdown(string markdown)
        {
            _markdownContent.AppendLine(markdown);
            _markdownContent.AppendLine();
        }

        // Add an image (stores image data, uploads later)
        public async Task AddImage(byte[] imageData, string fileName, string listId)
        {
            //string placeholder = $"[[IMAGE_PLACEHOLDER_{_pendingImages.Count}]]";
            //_pendingImages.Add(new ImageBlock
            //{
            //    ImageData = imageData,
            //    FileName = fileName,
            //    Placeholder = placeholder
            //});

            if (string.IsNullOrEmpty(listId))
            {
                throw new ArgumentException("listId is required to upload images");
            }

            string imageUrl = await UploadImageViaTaskAsync(imageData, fileName, listId);
            //string imageUrl = await UploadImageViaTaskAsync(imageData, fileName, listId);

            //// Replace placeholder with markdown image syntax
            //updatedMarkdown = updatedMarkdown.Replace(
            //    imageBlock.Placeholder,
            //    $"![{imageBlock.FileName}]({imageUrl})"
            //);

            //_markdownContent.AppendLine(placeholder);
            _markdownContent.AppendLine($"![{fileName}]({imageUrl})");
            //_markdownContent.AppendLine();
        }

        public void AddBulletPoint(string text)
        {
            // Add bullet point formatting
            //contentBlocks.Add(new { type = "bullet", text = text });
            string txt = text.Replace("o", ""); // Remove leading "o " if present

            _markdownContent.AppendLine($"* {txt.Trim()}");
            _markdownContent.AppendLine();
        }

        public void AddListItem(string text)
        {
            // Add numbered list item
            //contentBlocks.Add(new { type = "numbered", text = text });
            _markdownContent.AppendLine(text);
            _markdownContent.AppendLine();
        }

        public void AddNumberedListItem(string text, string number)
        {
            // Add numbered list item
            //contentBlocks.Add(new { type = "numbered", text = text });
            _markdownContent.AppendLine($"{number} {text}");
            _markdownContent.AppendLine();
        }

        public void AddCodeBlock(string text)
        {
            // Add code block with language
            //contentBlocks.Add(new { type = "code", text = code, language = language });
            _markdownContent.AppendLine($"`{text}`");
            _markdownContent.AppendLine();
        }

        public void AddBlockQuote(string text)
        {
            // Add block quote
            //contentBlocks.Add(new { type = "quote", text = text });
            _markdownContent.AppendLine($"> {text}");
            _markdownContent.AppendLine();
        }

        // WORKAROUND: Upload image to a task first, then reference in page
        // ClickUp doesn't have direct page attachment API, but task attachments work
        private async Task<string> UploadImageViaTaskAsync(byte[] imageData, string fileName, string listId)
        {
            // Step 1: Create a temporary task to hold the image
            var taskPayload = new JsonObject
            {
                ["name"] = $"[Image Upload] {fileName}"
            };

            var taskResponse = await _httpClient.PostAsync(
                $"https://api.clickup.com/api/v2/list/{listId}/task",
                new StringContent(taskPayload.ToJsonString(), Encoding.UTF8, "application/json")
            );

            if (!taskResponse.IsSuccessStatusCode)
            {
                throw new HttpRequestException($"Failed to create temporary task for image upload");
            }

            var taskJson = await taskResponse.Content.ReadAsStringAsync();
            var taskResult = JsonSerializer.Deserialize<JsonNode>(taskJson);
            var taskId = taskResult["id"].GetValue<string>();

            // Step 2: Upload image to the task
            var content = new MultipartFormDataContent();
            var fileContent = new ByteArrayContent(imageData);
            fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse(GetContentType(fileName));
            content.Add(fileContent, "attachment", fileName);

            var uploadResponse = await _httpClient.PostAsync(
                $"https://api.clickup.com/api/v2/task/{taskId}/attachment",
                content
            );

            if (!uploadResponse.IsSuccessStatusCode)
            {
                var errorContent = await uploadResponse.Content.ReadAsStringAsync();
                throw new HttpRequestException($"Failed to upload image to task. Status: {uploadResponse.StatusCode}, Error: {errorContent}");
            }

            var uploadJson = await uploadResponse.Content.ReadAsStringAsync();
            var uploadResult = JsonSerializer.Deserialize<JsonNode>(uploadJson);
            var imageUrl = uploadResult["url"].GetValue<string>();

            // Step 3: Optionally delete the temporary task (or keep it as an image library)
            // await _httpClient.DeleteAsync($"https://api.clickup.com/api/v2/task/{taskId}");

            Console.WriteLine($"Image uploaded via task: {fileName}");
            return imageUrl;
        }

        // Helper method to determine content type
        private string GetContentType(string fileName)
        {
            string ext = Path.GetExtension(fileName).ToLower();
            return ext switch
            {
                ".jpg" or ".jpeg" => "image/jpeg",
                ".png" => "image/png",
                ".gif" => "image/gif",
                ".webp" => "image/webp",
                ".svg" => "image/svg+xml",
                _ => "application/octet-stream"
            };
        }

        // Create page with markdown content, then upload images and update
        // Set uploadMethod: "base64" (default), "task" (uses task attachment workaround), or "external" (needs implementation)
        public async Task<string> CreateAndPopulatePageAsync(string workspaceId, string docId, string pageName,
            string parentPageId = null, string uploadMethod = "base64", string listIdForTaskUpload = null)
        {
            // Step 1: Create page with markdown content (with placeholders for images)
            var markdownText = _markdownContent.ToString();

            var createPayload = new JsonObject
            {
                ["name"] = pageName,
                ["content"] = markdownText
            };

            if (!string.IsNullOrEmpty(parentPageId))
            {
                createPayload["parent_page_id"] = parentPageId;
            }

            var jsonString = createPayload.ToJsonString();
            ConsoleHelper.LogInformation($"Creating Page - Request URL: https://api.clickup.com/api/v3/workspaces/{workspaceId}/docs/{docId}/pages");
            ConsoleHelper.LogInformation($"Request Body (first 500 chars): {jsonString.Substring(0, Math.Min(500, jsonString.Length))}...");

            var content = new StringContent(jsonString, Encoding.UTF8, "application/json");

            var response = await _httpClient.PostAsync(
                $"https://api.clickup.com/api/v3/workspaces/{workspaceId}/docs/{docId}/pages",
                content
            );

            if (!response.IsSuccessStatusCode)
            {
                var errorContent = await response.Content.ReadAsStringAsync();
                Console.WriteLine($"Error Response: {errorContent}");
                throw new HttpRequestException($"Failed to create Page. Status: {response.StatusCode}, Error: {errorContent}");
            }

            var responseJson = await response.Content.ReadAsStringAsync();
            var result = JsonSerializer.Deserialize<JsonNode>(responseJson);
            var pageId = result["id"].GetValue<string>();

            ConsoleHelper.WriteInfo($"Created page: {pageName} (ID: {pageId})");

            return pageId;
        }
    }
}