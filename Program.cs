//using A = DocumentFormat.OpenXml.Wordprocessing; // Alias needed for Drawing
using ClickUpDocumentImporter.DocumentConverter;
using ClickUpDocumentImporter.Helpers;
using Path = System.IO.Path;

namespace ClickUpDocumentImporter
{
    internal class ImageInfo
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

        private static ClickUpClient client;
        private static HttpClient clickupClient = new HttpClient();
        private static List<PageInfo> allPages = new List<PageInfo>();
        private static int imageCounter = 0;
        private static readonly string addPagesToDoc = "Screen Logic Customization"; // Page to add Documents

#if DEBUG
        private static string documentsFolder = @"C:\temp\CustomizedScreenLogic";
#else
        private static string documentsFolder = @"C:\temp\";
#endif

        private static async Task Main(string[] args)
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
            client = new ClickUpClient();
            clickupClient = client.ClickUpHttpClient;

            // *** List all pages in your space to select parent page
            var selectionList = await ListPagesInSpace();

            if (selectionList.Count == 0)
            {
                ConsoleHelper.WriteSeparator();
                ConsoleHelper.WriteError("No pages found in the ClickUp Wiki");
                ConsoleHelper.WriteError("Run terminated.");
                ConsoleHelper.Pause();
                return;
            }

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

        private static bool EnterDocumentDirectory()
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

        private static string GetPageSelection(List<SelectionItem> items)
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
                return (string)selected.Value;
            }
            else
            {
                ConsoleHelper.WriteWarning("Selection was cancelled");
                return string.Empty;
            }
        }

        private static async Task<List<SelectionItem>> ListPagesInSpace()
        {
            ConsoleHelper.LogInformation("Fetching pages in Wiki...\n");

            var response = await clickupClient.GetAsync(
                $"workspaces/{WORKSPACE_ID}/docs/{WIKI_ID}/pages"
                //$"https://api.clickup.com/api/v3/workspaces/{WORKSPACE_ID}/docs/{WIKI_ID}/pages"
            );

            //var response = await client.GetWithRetryAsync(
            //    $"https://api.clickup.com/api/v3/workspaces/{WORKSPACE_ID}/docs/{WIKI_ID}/pages"
            //);

            var content = await response.Content.ReadAsStringAsync();

            if (response.IsSuccessStatusCode)
            {
                string json = content;

                // Extract pages maintaining hierarchy
                var pages = PageExtractor.ExtractPages(json);
                allPages = pages;

                //// Print hierarchical structure
                //Console.WriteLine("Hierarchical Structure:");
                //PageExtractor.PrintHierarchy(pages);

                var selectionItems = new List<SelectionItem>();
                PageExtractor.ExtractPageHierarchy(pages, selectionItems);

                return selectionItems;
            }
            else
            {
                var errorContent = content;

                ConsoleHelper.WriteSeparator();
                ConsoleHelper.WriteError($"Could not retrieve ClickUp Wiki page titles. Status Code: {response.StatusCode}");
                ConsoleHelper.WriteError($"Error: {errorContent}");
            }

            return [];
        }
    }
}