using System.Text.Json;

namespace ClickUpDocumentImporter.Helpers
{
    public class PageInfo
    {
        public string Id { get; set; }
        public string DocId { get; set; }
        public string ParentPageId { get; set; }
        public string Name { get; set; }
        public List<PageInfo> Pages { get; set; } = new List<PageInfo>();
    }

    /// <summary>
    /// This class provides:
    /// 1. PageInfo class: A simple model containing only the 4 properties you need (id, doc_id,
    ///    parent_page_id, name) plus a nested Pages collection to maintain hierarchy
    /// 2. ExtractPages method: The main entry point that parses your JSON and returns a
    ///    hierarchical list of pages
    /// 3. ParsePages method: A recursive helper that processes each page and its nested pages,
    ///    maintaining the parent-child relationships
    /// 4. FindPageByName method: A recursive helper that find a single page by name
    /// 5. FlattenPages method: An optional utility if you need all pages in a single flat list
    ///    instead of nested structure
    /// 6. PrintHierarchy method: A helper to visualize the hierarchy
    ///
    /// The recursive approach automatically handles any depth of nesting, so whether you have 2
    /// levels or 10 levels of nested pages, it will extract them all while preserving the structure.
    /// </summary>
    public class PageExtractor
    {
        public static List<PageInfo> ExtractPages(string json)
        {
            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            };

            using JsonDocument doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            // Check if root is an array or object
            if (root.ValueKind == JsonValueKind.Array)
            {
                // Root is already an array of pages
                return ParsePages(root);
            }
            else if (root.ValueKind == JsonValueKind.Object)
            {
                // Root is an object, check for pages property
                if (root.TryGetProperty("pages", out JsonElement pagesElement))
                {
                    return ParsePages(pagesElement);
                }
            }

            return new List<PageInfo>();
        }

        private static List<PageInfo> ParsePages(JsonElement pagesElement)
        {
            var pages = new List<PageInfo>();

            if (pagesElement.ValueKind != JsonValueKind.Array)
                return pages;

            foreach (JsonElement pageElement in pagesElement.EnumerateArray())
            {
                var page = new PageInfo();

                if (pageElement.TryGetProperty("id", out JsonElement idElement))
                    page.Id = idElement.GetString();

                if (pageElement.TryGetProperty("doc_id", out JsonElement docIdElement))
                    page.DocId = docIdElement.GetString();

                if (pageElement.TryGetProperty("parent_page_id", out JsonElement parentIdElement))
                    page.ParentPageId = parentIdElement.GetString();

                if (pageElement.TryGetProperty("name", out JsonElement nameElement))
                    page.Name = nameElement.GetString();

                // Recursively process nested pages
                if (pageElement.TryGetProperty("pages", out JsonElement nestedPagesElement))
                {
                    page.Pages = ParsePages(nestedPagesElement);
                }

                pages.Add(page);
            }

            return pages;
        }

        // Find a single page by name (searches recursively)
        public static PageInfo FindPageByName(List<PageInfo> pages, string name, bool caseSensitive = false)
        {
            foreach (var page in pages)
            {
                bool isMatch = caseSensitive
                    ? page.Name == name
                    : page.Name?.Equals(name, StringComparison.OrdinalIgnoreCase) ?? false;

                if (isMatch)
                    return page;

                // Recursively search nested pages
                if (page.Pages.Count > 0)
                {
                    var found = FindPageByName(page.Pages, name, caseSensitive);
                    if (found != null)
                        return found;
                }
            }

            return null;
        }

        // Helper method to flatten the hierarchy if needed
        public static List<PageInfo> FlattenPages(List<PageInfo> pages)
        {
            var result = new List<PageInfo>();

            foreach (var page in pages)
            {
                result.Add(page);
                if (page.Pages.Count > 0)
                {
                    result.AddRange(FlattenPages(page.Pages));
                }
            }

            return result;
        }

        // Helper method to print hierarchy
        public static void PrintHierarchy(List<PageInfo> pages, int indent = 0)
        {
            foreach (var page in pages)
            {
                Console.WriteLine($"{new string(' ', indent * 2)}- {page.Name} (ID: {page.Id}, Parent: {page.ParentPageId ?? "null"})");
                if (page.Pages.Count > 0)
                {
                    PrintHierarchy(page.Pages, indent + 1);
                }
            }
        }
    }
}

//// Usage example
//class Program
//{
//    static void Main()
//    {
//        string json = @"{ your JSON here }";

//        // Extract pages maintaining hierarchy
//        var pages = PageExtractor.ExtractPages(json);

//        // Print hierarchical structure
//        Console.WriteLine("Hierarchical Structure:");
//        PageExtractor.PrintHierarchy(pages);

//        Console.WriteLine("\n---\n");

//        // Or flatten if you need a simple list
//        var flatPages = PageExtractor.FlattenPages(pages);
//        Console.WriteLine("Flattened List:");
//        foreach (var page in flatPages)
//        {
//            Console.WriteLine($"- {page.Name} (ID: {page.Id}, Parent: {page.ParentPageId ?? "null"})");
//        }
//    }
//}