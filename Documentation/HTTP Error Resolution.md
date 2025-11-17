# HTTP Error Resolution

I have a C# console application that accesses a list of pages under a ClickUp Wiki.  It was working but is not return a list of page now.  I get the following response from the API request: context deadline exceeded.

Here is the C# statement to the API.  How do I correct this issue to move forward?

# Solutions

The "context deadline exceeded" error typically means your HTTP request is timing out before the API can respond. This could happen if the Wiki has many pages or if there are network/API performance issues.

Here are several ways to fix this:

## 1. **Increase the HTTP Client Timeout**

```csharp
// Set a longer timeout (default is usually 100 seconds)
clickupClient.Timeout = TimeSpan.FromMinutes(5); // or whatever duration you need

var response = await clickupClient.GetAsync(
    $"https://api.clickup.com/api/v3/workspaces/{WORKSPACE_ID}/docs/{WIKI_ID}/pages"
);
```

## 2. **Use Pagination (Recommended)**

If the Wiki has many pages, you should use pagination to retrieve them in smaller batches:

```csharp
int page = 0;
int pageSize = 50; // Adjust based on API limits
var allPages = new List<YourPageType>();

while (true)
{
    var response = await clickupClient.GetAsync(
        $"https://api.clickup.com/api/v3/workspaces/{WORKSPACE_ID}/docs/{WIKI_ID}/pages?page={page}&page_size={pageSize}"
    );
    
    response.EnsureSuccessStatusCode();
    var content = await response.Content.ReadAsStringAsync();
    var result = JsonSerializer.Deserialize<YourResponseType>(content);
    
    if (result.Pages == null || result.Pages.Count == 0)
        break;
        
    allPages.AddRange(result.Pages);
    page++;
}
```

## 3. **Add Retry Logic with Exponential Backoff**

```csharp
int maxRetries = 3;
int retryCount = 0;

while (retryCount < maxRetries)
{
    try
    {
        var response = await clickupClient.GetAsync(
            $"https://api.clickup.com/api/v3/workspaces/{WORKSPACE_ID}/docs/{WIKI_ID}/pages"
        );
        
        if (response.IsSuccessStatusCode)
        {
            // Process response
            break;
        }
    }
    catch (TaskCanceledException)
    {
        retryCount++;
        if (retryCount >= maxRetries)
            throw;
            
        await Task.Delay(TimeSpan.FromSeconds(Math.Pow(2, retryCount))); // Exponential backoff
    }
}
```

## 4. **Check ClickUp API Documentation**

The ClickUp API v3 might have specific requirements or limits. Check if:
- There are query parameters to limit results
- The endpoint supports filtering or pagination
- There's a rate limit you're hitting

## Quick Fix to Try First

The Quick Fix did not work (November 17, 2025)

```csharp
// Increase timeout first
clickupClient.Timeout = TimeSpan.FromMinutes(3);

var response = await clickupClient.GetAsync(
    $"https://api.clickup.com/api/v3/workspaces/{WORKSPACE_ID}/docs/{WIKI_ID}/pages"
);

if (!response.IsSuccessStatusCode)
{
    var error = await response.Content.ReadAsStringAsync();
    Console.WriteLine($"Error: {response.StatusCode} - {error}");
}
```

**Most likely solution**: Start by increasing the timeout, then investigate if the API supports pagination or if there's been a change in the API that requires different parameters.