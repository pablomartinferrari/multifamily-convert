using System;
using System.Collections.Generic;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using ExcelProcessor.Auth;
using ExcelProcessor.Models;
using Newtonsoft.Json;

namespace ExcelProcessor.Services
{
    public class SharePointService
    {
        private readonly SharePointAuthenticator _authenticator;
        private readonly string _siteUrl;

        public SharePointService(SharePointAuthenticator authenticator, string siteUrl)
        {
            _authenticator = authenticator;
            _siteUrl = siteUrl;
        }

        public async Task<byte[]> DownloadFileAsync(string fileUrl)
        {
            using var client = await _authenticator.GetAuthenticatedHttpClientAsync();
            var apiUrl = fileUrl.Replace(_siteUrl, $"{_siteUrl}/_api/web").Replace("/Forms/AllItems.aspx", "") + "/$value";
            if (!apiUrl.Contains("_api/web"))
            {
                // If it's a relative URL or already an absolute URL without API
                apiUrl = $"{_siteUrl}/_api/web/GetFileByServerRelativeUrl('{new Uri(fileUrl).AbsolutePath}')/$value";
            }
            
            var response = await client.GetAsync(apiUrl);
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadAsByteArrayAsync();
        }

        public async Task<string> UploadFileAsync(byte[] content, string fileName, string folderUrl)
        {
            using var client = await _authenticator.GetAuthenticatedHttpClientAsync();
            var uploadUrl = $"{_siteUrl}/_api/web/GetFolderByServerRelativeUrl('{folderUrl}')/Files/add(url='{fileName}',overwrite=true)";
            
            using var contentStream = new ByteArrayContent(content);
            contentStream.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
            
            var response = await client.PostAsync(uploadUrl, contentStream);
            response.EnsureSuccessStatusCode();
            
            var result = JsonConvert.DeserializeObject<dynamic>(await response.Content.ReadAsStringAsync());
            return result.d.ServerRelativeUrl;
        }

        public async Task CreateFolderIfNotExistsAsync(string serverRelativeUrl)
        {
            using var client = await _authenticator.GetAuthenticatedHttpClientAsync();
            var folderUrl = $"{_siteUrl}/_api/web/folders";
            var body = new { __metadata = new { type = "SP.Folder" }, ServerRelativeUrl = serverRelativeUrl };
            var jsonBody = JsonConvert.SerializeObject(body);
            var content = new StringContent(jsonBody, Encoding.UTF8, "application/json");
            
            await client.PostAsync(folderUrl, content);
            // We ignore errors here as it might already exist
        }
    }
}

