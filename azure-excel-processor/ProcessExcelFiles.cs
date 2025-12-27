using System.Net;
using System.IO;
using System.Linq;
using ExcelProcessor.Auth;
using ExcelProcessor.Models;
using ExcelProcessor.Services;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

namespace ExcelProcessor
{
    public class ProcessExcelFiles
    {
        private readonly ILogger _logger;
        private readonly ExcelService _excelService;
        private readonly XrfProcessingService _processingService;

        public ProcessExcelFiles(ILoggerFactory loggerFactory)
        {
            _logger = loggerFactory.CreateLogger<ProcessExcelFiles>();
            _excelService = new ExcelService();
            _processingService = new XrfProcessingService();
        }

        [Function("ProcessExcelFiles")]
        public async Task<HttpResponseData> Run([HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequestData req)
        {
            _logger.LogInformation("C# HTTP trigger function processed a request.");

            var response = req.CreateResponse(HttpStatusCode.OK);
            var processingResponse = new ProcessingResponse();

            try
            {
                string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
                var data = JsonConvert.DeserializeObject<ProcessingRequest>(requestBody);

                if (data == null || string.IsNullOrEmpty(data.SiteUrl) || data.FileUrls.Count == 0)
                {
                    processingResponse.Success = false;
                    processingResponse.Message = "Invalid request parameters.";
                    await response.WriteAsJsonAsync(processingResponse);
                    return response;
                }

                var clientId = Environment.GetEnvironmentVariable("APP_CLIENT_ID");
                var clientSecret = Environment.GetEnvironmentVariable("APP_CLIENT_SECRET");
                var tenantId = Environment.GetEnvironmentVariable("APP_TENANT_ID");

                var authenticator = new SharePointAuthenticator(clientId, clientSecret, tenantId, data.SiteUrl);
                var spService = new SharePointService(authenticator, data.SiteUrl);

                var allShots = new List<XrfShot>();

                foreach (var fileUrl in data.FileUrls)
                {
                    _logger.LogInformation($"Processing file: {fileUrl}");
                    var content = await spService.DownloadFileAsync(fileUrl);
                    var shots = _excelService.ReadShotsFromExcel(content);
                    allShots.AddRange(shots);
                }

                var results = _processingService.ProcessShots(allShots);

                // Generate and upload reports
                var timeStamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                var resultsFolder = "/Shared Documents/Processed Reports"; // Default library path
                
                await spService.CreateFolderIfNotExistsAsync(resultsFolder);
                if (results.AveragedResults.Any())
                {
                    var title = $"AVERAGED DWELLING {data.FileType.ToUpper()} COMPONENT RESULTS";
                    var fileName = $"{data.JobNumber}_Averaged_{timeStamp}.xlsx";
                    var content = _excelService.CreateSummaryExcel(results.AveragedResults, title);
                    var serverRelativeUrl = await spService.UploadFileAsync(content, fileName, resultsFolder);
                    processingResponse.GeneratedReports.Add(new GeneratedReport { ReportType = "Averaged", FileName = fileName, Url = $"{data.SiteUrl}{serverRelativeUrl}" });
                }

                // 2. Uniform Results
                if (results.UniformResults.Any())
                {
                    var title = $"INDIVIDUALLY TESTED {data.FileType.ToUpper()} COMPONENTS (UNIFORM RESULTS)";
                    var fileName = $"{data.JobNumber}_Uniform_{timeStamp}.xlsx";
                    var content = _excelService.CreateSummaryExcel(results.UniformResults, title);
                    var serverRelativeUrl = await spService.UploadFileAsync(content, fileName, resultsFolder);
                    processingResponse.GeneratedReports.Add(new GeneratedReport { ReportType = "Uniform", FileName = fileName, Url = $"{data.SiteUrl}{serverRelativeUrl}" });
                }

                // 3. Conflicting Results
                if (results.ConflictingResults.Any())
                {
                    var title = $"INDIVIDUALLY TESTED {data.FileType.ToUpper()} COMPONENTS (CONFLICTING RESULTS)";
                    var fileName = $"{data.JobNumber}_Conflicting_{timeStamp}.xlsx";
                    var content = _excelService.CreateConflictingExcel(results.ConflictingResults, title);
                    var serverRelativeUrl = await spService.UploadFileAsync(content, fileName, resultsFolder);
                    processingResponse.GeneratedReports.Add(new GeneratedReport { ReportType = "Conflicting", FileName = fileName, Url = $"{data.SiteUrl}{serverRelativeUrl}" });
                }

                processingResponse.Success = true;
                processingResponse.Message = "Processing completed successfully.";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing XRF data");
                processingResponse.Success = false;
                processingResponse.Message = $"Error: {ex.Message}";
                processingResponse.Errors.Add(ex.ToString());
            }

            await response.WriteAsJsonAsync(processingResponse);
            return response;
        }
    }
}

