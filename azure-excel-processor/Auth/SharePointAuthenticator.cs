using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Identity.Client;

namespace ExcelProcessor.Auth
{
    public class SharePointAuthenticator
    {
        private readonly IConfidentialClientApplication _app;
        private readonly string _sharePointUrl;

        public SharePointAuthenticator(string clientId, string clientSecret, string tenantId, string sharePointUrl)
        {
            _sharePointUrl = sharePointUrl;
            _app = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithClientSecret(clientSecret)
                .WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantId}"))
                .Build();
        }

        public async Task<string> GetAccessTokenAsync()
        {
            string[] scopes = { $"{_sharePointUrl}/.default" };
            var result = await _app.AcquireTokenForClient(scopes).ExecuteAsync();
            return result.AccessToken;
        }

        public async Task<HttpClient> GetAuthenticatedHttpClientAsync()
        {
            var token = await GetAccessTokenAsync();
            var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.Add("Accept", "application/json;odata=verbose");
            return client;
        }
    }
}

