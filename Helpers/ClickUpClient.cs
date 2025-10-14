using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace ClickUpDocumentImporter.Helpers
{
    public class ClickUpClient
    {

        private readonly HttpClient _httpClient;
        private readonly string _apiKey = Globals.CLICKUP_API_KEY;
        //private readonly string _workspaceId;
        //private readonly string _spaceId;
        //private readonly string _folderId;
        //private readonly string _listId;


        public ClickUpClient() // string apiKey, string workspaceId, string spaceId) //, string folderId, string listId)
        {
            //_apiKey = apiKey;
            //_workspaceId = workspaceId;
            //_spaceId = spaceId;
            //_folderId = folderId;
            //_listId = listId;

            // Create HttpClient instance
            _httpClient = new HttpClient();
            _httpClient.BaseAddress = new Uri("https://api.clickup.com/api/v2/");

            _httpClient.DefaultRequestHeaders.Clear();
            _httpClient.DefaultRequestHeaders.Add("Authorization", _apiKey);

            _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            ClickUpHttpClient = _httpClient;
        }


        public HttpClient ClickUpHttpClient { get; set; }
    }
}
