using System.Net.Http.Headers;

namespace ClickUpDocumentImporter.Helpers
{
    public class ClickUpClient
    {
        private readonly HttpClient _httpClient;
        private readonly string _apiKey = Globals.CLICKUP_API_KEY;

        public ClickUpClient() // string apiKey, string workspaceId, string spaceId) //, string folderId, string listId)
        {
            // Create HttpClient instance
            _httpClient = new HttpClient();
            _httpClient.BaseAddress = new Uri("https://api.clickup.com/api/v3/");

            //_httpClient.Timeout = TimeSpan.FromSeconds(120); // Increase from default 100s
            _httpClient.DefaultRequestHeaders.Clear();
            _httpClient.DefaultRequestHeaders.Add("Authorization", _apiKey);

            _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            ClickUpHttpClient = _httpClient;
        }

        public HttpClient ClickUpHttpClient { get; set; }

        /// <summary>
        /// Sends an HTTP GET request to the specified URL with retry logic for transient failures.
        /// </summary>
        /// <remarks>This method implements an exponential backoff strategy for retries, with delays of 1,
        /// 2, and 4 seconds  between attempts by default. Retries are triggered for server errors (HTTP 5xx) or when a 
        /// <see cref="TaskCanceledException"/> occurs due to a timeout. If the maximum number of retries is exceeded, 
        /// an exception is thrown.</remarks>
        /// <param name="url">The URL to which the GET request is sent. Cannot be null or empty.</param>
        /// <param name="maxRetries">The maximum number of retry attempts in case of transient failures, such as server errors (HTTP 5xx)  or
        /// request timeouts. Must be greater than or equal to 1. Defaults to 3.</param>
        /// <returns>An <see cref="HttpResponseMessage"/> representing the HTTP response received from the server.  The response
        /// may indicate success or failure depending on the server's status.</returns>
        /// <exception cref="Exception">Thrown if the maximum number of retries is exceeded without a successful response.</exception>
        public async Task<HttpResponseMessage> GetWithRetryAsync(string url, int maxRetries = 3)
        {
            for (int i = 0; i < maxRetries; i++)
            {
                try
                {
                    var response = await _httpClient.GetAsync(url);

                    if (response.IsSuccessStatusCode)
                    {
                        return response;
                    }

                    // If it's a server error (5xx), retry
                    if ((int)response.StatusCode >= 500)
                    {
                        if (i < maxRetries - 1)
                        {
                            var delay = TimeSpan.FromSeconds(Math.Pow(2, i)); // 1s, 2s, 4s
                            ConsoleHelper.WriteError($"Server error. Retrying in {delay.TotalSeconds}s...");
                            await Task.Delay(delay);
                            continue;
                        }
                    }

                    return response;
                }
                catch (TaskCanceledException ex)
                {
                    if (i < maxRetries - 1)
                    {
                        ConsoleHelper.WriteError($"Request timeout. Retrying... ({i + 1}/{maxRetries})");
                        await Task.Delay(TimeSpan.FromSeconds(Math.Pow(2, i)));
                        continue;
                    }
                    throw;
                }
            }

            throw new Exception("Max retries exceeded");
        }
    }
}