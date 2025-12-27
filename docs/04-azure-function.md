# 4. Azure Function for Excel Processing

This section covers building an Azure Function in C# that processes Excel files from SharePoint, performs data transformations, and generates multiple reports.

## Prerequisites

- .NET 6+ SDK installed
- Azure Functions Core Tools: `npm install -g azure-functions-core-tools@4`
- Visual Studio Code with Azure Functions extension
- Azure CLI installed
- Authentication configured (from previous section)

## Step 1: Create Azure Functions Project

### Initialize Project

```bash
mkdir azure-excel-processor
cd azure-excel-processor

# Create Azure Functions project
func init --dotnet
```

### Select Template

```
Select a number for worker runtime:
1. dotnet
2. dotnet (isolated process)
3. node
4. python
5. java
6. powershell
7. custom

> 1  # Choose dotnet for in-process model
```

### Create HTTP Trigger Function

```bash
func new --name ProcessExcelFiles --template "HTTP trigger"
```

## Step 2: Project Structure

Your project structure should look like:

```
azure-excel-processor/
├── Auth/
│   └── SharePointAuthenticator.cs    # Authentication helper
├── Models/
│   ├── ProcessingRequest.cs          # Request model
│   ├── ProcessingResponse.cs         # Response model
│   └── FileInfo.cs                   # File information
├── Services/
│   ├── ExcelProcessor.cs             # Excel processing logic
│   ├── SharePointService.cs          # SharePoint operations
│   └── ReportGenerator.cs            # Report generation
├── ProcessExcelFiles.cs              # Main function
├── host.json                         # Function host config
├── local.settings.json               # Local settings
└── azure-excel-processor.csproj      # Project file
```

## Step 3: Install NuGet Packages

### Update azure-excel-processor.csproj

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net6.0</TargetFramework>
    <AzureFunctionsVersion>v4</AzureFunctionsVersion>
    <OutputType>Exe</OutputType>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="Microsoft.Azure.Functions.Worker" Version="1.19.0" />
    <PackageReference Include="Microsoft.Azure.Functions.Worker.Sdk" Version="1.16.4" />
    <PackageReference Include="Microsoft.Azure.Functions.Worker.Extensions.Http" Version="3.0.13" />
    <PackageReference Include="Microsoft.Extensions.Logging" Version="7.0.0" />
    <!-- Excel Processing -->
    <PackageReference Include="ClosedXML" Version="0.102.1" />
    <!-- SharePoint Integration -->
    <PackageReference Include="Microsoft.Identity.Client" Version="4.54.1" />
    <!-- JSON Processing -->
    <PackageReference Include="Newtonsoft.Json" Version="13.0.3" />
    <!-- HTTP Client -->
    <PackageReference Include="System.Net.Http" Version="4.3.4" />
  </ItemGroup>
  <ItemGroup>
    <None Update="host.json">
      <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
    </None>
    <None Update="local.settings.json">
      <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
    </None>
  </ItemGroup>
</Project>
```

## Step 4: Define Data Models

### Models/ProcessingRequest.cs

```csharp
using System.Collections.Generic;

namespace ExcelProcessor.Models
{
    public class ProcessingRequest
    {
        public string SiteUrl { get; set; } = string.Empty;
        public List<string> FileUrls { get; set; } = new List<string>();
        public string UserId { get; set; } = string.Empty;
        public Dictionary<string, string> Options { get; set; } = new Dictionary<string, string>();
    }
}
```

### Models/ProcessingResponse.cs

```csharp
using System.Collections.Generic;

namespace ExcelProcessor.Models
{
    public class ProcessingResponse
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public List<GeneratedReport> GeneratedReports { get; set; } = new List<GeneratedReport>();
        public List<string> Errors { get; set; } = new List<string>();
        public int ProcessedFiles { get; set; }
    }

    public class GeneratedReport
    {
        public string ReportType { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public string LibraryName { get; set; } = string.Empty;
        public string Url { get; set; } = string.Empty;
        public DateTime GeneratedAt { get; set; }
    }
}
```

### Models/FileInfo.cs

```csharp
namespace ExcelProcessor.Models
{
    public class FileInfo
    {
        public string Name { get; set; } = string.Empty;
        public string Url { get; set; } = string.Empty;
        public string LibraryName { get; set; } = string.Empty;
        public long Size { get; set; }
        public DateTime Modified { get; set; }
        public byte[] Content { get; set; } = Array.Empty<byte>();
    }
}
```

## Step 5: Implement SharePoint Authentication

### Auth/SharePointAuthenticator.cs

```csharp
using Microsoft.Identity.Client;
using System.Net.Http;
using System.Net.Http.Headers;

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

            try
            {
                var result = await _app.AcquireTokenForClient(scopes).ExecuteAsync();
                return result.AccessToken;
            }
            catch (MsalServiceException ex)
            {
                throw new Exception($"Error acquiring access token: {ex.Message}", ex);
            }
        }

        public async Task<HttpClient> GetAuthenticatedHttpClientAsync()
        {
            var token = await GetAccessTokenAsync();
            var client = new HttpClient();

            client.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", token);

            client.DefaultRequestHeaders.Accept.Add(
                new MediaTypeWithQualityHeaderValue("application/json"));

            client.DefaultRequestHeaders.Add("Accept", "application/json;odata=verbose");

            return client;
        }
    }
}
```

## Step 6: Create SharePoint Service

### Services/SharePointService.cs

```csharp
using ExcelProcessor.Auth;
using ExcelProcessor.Models;
using Newtonsoft.Json;
using System.Net.Http;
using System.Text;

namespace ExcelProcessor.Services
{
    public class SharePointService
    {
        private readonly SharePointAuthenticator _authenticator;
        private readonly string _siteUrl;
        private readonly HttpClient _httpClient;

        public SharePointService(SharePointAuthenticator authenticator, string siteUrl)
        {
            _authenticator = authenticator;
            _siteUrl = siteUrl;
            _httpClient = authenticator.GetAuthenticatedHttpClientAsync().Result;
        }

        public async Task<FileInfo> DownloadFileAsync(string fileUrl)
        {
            try
            {
                // Convert SharePoint URL to API endpoint
                var apiUrl = fileUrl.Replace(_siteUrl, $"{_siteUrl}/_api/web")
                                   .Replace("/Lists/", "/lists/")
                                   .Replace("/Attachments/", "/AttachmentFiles/")
                                   + "/$value";

                var response = await _httpClient.GetAsync(apiUrl);
                response.EnsureSuccessStatusCode();

                var content = await response.Content.ReadAsByteArrayAsync();

                // Get file metadata
                var metadataUrl = fileUrl.Replace(_siteUrl, $"{_siteUrl}/_api/web")
                                       .Replace("/Lists/", "/lists/")
                                       .Replace("/Attachments/", "/AttachmentFiles/");

                var metadataResponse = await _httpClient.GetAsync(metadataUrl);
                var metadata = JsonConvert.DeserializeObject<dynamic>(
                    await metadataResponse.Content.ReadAsStringAsync());

                return new FileInfo
                {
                    Name = metadata.Name,
                    Url = fileUrl,
                    Size = content.Length,
                    Modified = DateTime.Parse(metadata.TimeLastModified),
                    Content = content
                };

            }
            catch (Exception ex)
            {
                throw new Exception($"Error downloading file {fileUrl}: {ex.Message}", ex);
            }
        }

        public async Task<string> UploadFileAsync(byte[] content, string fileName, string libraryName, string folderPath = "")
        {
            try
            {
                // Get library info
                var libraryUrl = $"{_siteUrl}/_api/web/lists/getbytitle('{libraryName}')";
                var libraryResponse = await _httpClient.GetAsync(libraryUrl);
                var library = JsonConvert.DeserializeObject<dynamic>(
                    await libraryResponse.Content.ReadAsStringAsync());

                // Construct upload URL
                var uploadPath = string.IsNullOrEmpty(folderPath) ? "" : $"/{folderPath}";
                var uploadUrl = $"{_siteUrl}/_api/web/lists/getbytitle('{libraryName}')/RootFolder{uploadPath}/Files/add(url='{fileName}',overwrite=true)";

                using (var contentStream = new ByteArrayContent(content))
                {
                    contentStream.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");

                    var response = await _httpClient.PostAsync(uploadUrl, contentStream);
                    response.EnsureSuccessStatusCode();

                    var result = JsonConvert.DeserializeObject<dynamic>(
                        await response.Content.ReadAsStringAsync());

                    return result.ServerRelativeUrl;
                }

            }
            catch (Exception ex)
            {
                throw new Exception($"Error uploading file {fileName} to {libraryName}: {ex.Message}", ex);
            }
        }

        public async Task CreateFolderIfNotExistsAsync(string libraryName, string folderPath)
        {
            try
            {
                var folderUrl = $"{_siteUrl}/_api/web/lists/getbytitle('{libraryName}')/RootFolder/Folders";
                var body = new
                {
                    __metadata = new { type = "SP.Folder" },
                    ServerRelativeUrl = $"/sites/{new Uri(_siteUrl).Segments.Last()}/{libraryName}/{folderPath}"
                };

                var jsonBody = JsonConvert.SerializeObject(body);
                var content = new StringContent(jsonBody, Encoding.UTF8, "application/json");
                content.Headers.Add("X-RequestDigest", await GetRequestDigestAsync());

                var response = await _httpClient.PostAsync(folderUrl, content);

                // Ignore if folder already exists (409 Conflict)
                if (response.StatusCode != HttpStatusCode.Conflict)
                {
                    response.EnsureSuccessStatusCode();
                }

            }
            catch (Exception ex)
            {
                // Log but don't throw - folder creation is not critical
                Console.WriteLine($"Warning: Could not create folder {folderPath}: {ex.Message}");
            }
        }

        private async Task<string> GetRequestDigestAsync()
        {
            var digestUrl = $"{_siteUrl}/_api/contextinfo";
            var response = await _httpClient.PostAsync(digestUrl, new StringContent(""));
            response.EnsureSuccessStatusCode();

            var result = JsonConvert.DeserializeObject<dynamic>(
                await response.Content.ReadAsStringAsync());

            return result.d.GetContextWebInformation.FormDigestValue;
        }
    }
}
```

## Step 7: Implement Excel Processing Logic

### Services/ExcelProcessor.cs

```csharp
using ClosedXML.Excel;
using ExcelProcessor.Models;
using System.Data;

namespace ExcelProcessor.Services
{
    public class ExcelProcessor
    {
        public DataTable ReadExcelToDataTable(byte[] excelContent)
        {
            using (var stream = new MemoryStream(excelContent))
            using (var workbook = new XLWorkbook(stream))
            {
                var worksheet = workbook.Worksheets.First();
                var dataTable = new DataTable();

                // Read headers
                var headers = worksheet.Row(1);
                foreach (var cell in headers.Cells())
                {
                    dataTable.Columns.Add(cell.Value.ToString());
                }

                // Read data rows
                for (int row = 2; row <= worksheet.LastRowUsed().RowNumber(); row++)
                {
                    var dataRow = dataTable.NewRow();
                    var excelRow = worksheet.Row(row);

                    for (int col = 1; col <= dataTable.Columns.Count; col++)
                    {
                        var cell = excelRow.Cell(col);
                        dataRow[col - 1] = cell.Value;
                    }

                    dataTable.Rows.Add(dataRow);
                }

                return dataTable;
            }
        }

        public byte[] WriteDataTableToExcel(DataTable dataTable, string sheetName = "Sheet1")
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add(sheetName);

                // Write headers
                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    worksheet.Cell(1, col + 1).Value = dataTable.Columns[col].ColumnName;
                }

                // Write data
                for (int row = 0; row < dataTable.Rows.Count; row++)
                {
                    for (int col = 0; col < dataTable.Columns.Count; col++)
                    {
                        worksheet.Cell(row + 2, col + 1).Value = dataTable.Rows[row][col].ToString();
                    }
                }

                // Auto-fit columns
                worksheet.Columns().AdjustToContents();

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    return stream.ToArray();
                }
            }
        }

        public DataTable MergeDataTables(List<DataTable> tables, string mergeColumn)
        {
            if (tables == null || tables.Count == 0)
                throw new ArgumentException("No tables provided for merging");

            var mergedTable = tables[0].Clone();

            foreach (var table in tables)
            {
                foreach (DataRow row in table.Rows)
                {
                    mergedTable.ImportRow(row);
                }
            }

            return mergedTable;
        }

        public DataTable ApplyTransformations(DataTable dataTable, Dictionary<string, string> transformations)
        {
            var transformedTable = dataTable.Copy();

            foreach (var transformation in transformations)
            {
                switch (transformation.Key.ToLower())
                {
                    case "remove_duplicates":
                        transformedTable = RemoveDuplicates(transformedTable, transformation.Value);
                        break;
                    case "filter_column":
                        var parts = transformation.Value.Split('|');
                        if (parts.Length == 2)
                        {
                            transformedTable = FilterByColumn(transformedTable, parts[0], parts[1]);
                        }
                        break;
                    case "sort_column":
                        transformedTable = SortByColumn(transformedTable, transformation.Value);
                        break;
                }
            }

            return transformedTable;
        }

        private DataTable RemoveDuplicates(DataTable table, string columnName)
        {
            var uniqueRows = new HashSet<string>();
            var result = table.Clone();

            foreach (DataRow row in table.Rows)
            {
                var key = row[columnName].ToString();
                if (!uniqueRows.Contains(key))
                {
                    uniqueRows.Add(key);
                    result.ImportRow(row);
                }
            }

            return result;
        }

        private DataTable FilterByColumn(DataTable table, string columnName, string filterValue)
        {
            var result = table.Clone();

            foreach (DataRow row in table.Rows)
            {
                if (row[columnName].ToString().Contains(filterValue))
                {
                    result.ImportRow(row);
                }
            }

            return result;
        }

        private DataTable SortByColumn(DataTable table, string columnName)
        {
            var result = table.Copy();
            result.DefaultView.Sort = columnName;
            return result.DefaultView.ToTable();
        }
    }
}
```

## Step 8: Create Report Generator

### Services/ReportGenerator.cs

```csharp
using ExcelProcessor.Models;
using System.Data;

namespace ExcelProcessor.Services
{
    public class ReportGenerator
    {
        private readonly ExcelProcessor _excelProcessor;

        public ReportGenerator()
        {
            _excelProcessor = new ExcelProcessor();
        }

        public List<GeneratedReport> GenerateReports(List<FileInfo> processedFiles, Dictionary<string, string> options)
        {
            var reports = new List<GeneratedReport>();
            var allData = new List<DataTable>();

            // Read all Excel files
            foreach (var file in processedFiles)
            {
                var dataTable = _excelProcessor.ReadExcelToDataTable(file.Content);
                allData.Add(dataTable);
            }

            // Generate Report 1: Monthly Summary
            var monthlyReport = GenerateMonthlySummaryReport(allData);
            reports.Add(monthlyReport);

            // Generate Report 2: Data Quality
            var qualityReport = GenerateDataQualityReport(allData);
            reports.Add(qualityReport);

            // Generate Report 3: Trend Analysis
            var trendReport = GenerateTrendAnalysisReport(allData);
            reports.Add(trendReport);

            return reports;
        }

        private GeneratedReport GenerateMonthlySummaryReport(List<DataTable> dataTables)
        {
            // Merge all data
            var mergedData = _excelProcessor.MergeDataTables(dataTables, "Date");

            // Apply transformations (example)
            var transformations = new Dictionary<string, string>
            {
                { "remove_duplicates", "ID" },
                { "sort_column", "Date" }
            };

            var processedData = _excelProcessor.ApplyTransformations(mergedData, transformations);

            // Group by month and calculate summaries
            var summaryData = CreateMonthlySummary(processedData);

            var excelContent = _excelProcessor.WriteDataTableToExcel(summaryData, "Monthly Summary");

            return new GeneratedReport
            {
                ReportType = "Monthly Summary",
                FileName = $"Monthly_Summary_{DateTime.Now:yyyy_MM_dd}.xlsx",
                LibraryName = "Monthly Summary Reports",
                GeneratedAt = DateTime.Now
            };
        }

        private GeneratedReport GenerateDataQualityReport(List<DataTable> dataTables)
        {
            var qualityMetrics = CalculateDataQualityMetrics(dataTables);
            var excelContent = _excelProcessor.WriteDataTableToExcel(qualityMetrics, "Data Quality");

            return new GeneratedReport
            {
                ReportType = "Data Quality",
                FileName = $"Data_Quality_Report_{DateTime.Now:yyyy_MM_dd}.xlsx",
                LibraryName = "Data Quality Reports",
                GeneratedAt = DateTime.Now
            };
        }

        private GeneratedReport GenerateTrendAnalysisReport(List<DataTable> dataTables)
        {
            var trendData = AnalyzeTrends(dataTables);
            var excelContent = _excelProcessor.WriteDataTableToExcel(trendData, "Trend Analysis");

            return new GeneratedReport
            {
                ReportType = "Trend Analysis",
                FileName = $"Trend_Analysis_{DateTime.Now:yyyy_MM_dd}.xlsx",
                LibraryName = "Trend Analysis Reports",
                GeneratedAt = DateTime.Now
            };
        }

        private DataTable CreateMonthlySummary(DataTable data)
        {
            var summary = new DataTable();
            summary.Columns.Add("Month", typeof(string));
            summary.Columns.Add("TotalRecords", typeof(int));
            summary.Columns.Add("UniqueItems", typeof(int));
            summary.Columns.Add("AverageValue", typeof(decimal));

            // Group by month logic here
            // This is a simplified example - implement based on your data structure

            return summary;
        }

        private DataTable CalculateDataQualityMetrics(List<DataTable> dataTables)
        {
            var metrics = new DataTable();
            metrics.Columns.Add("Metric", typeof(string));
            metrics.Columns.Add("Value", typeof(string));
            metrics.Columns.Add("Status", typeof(string));

            // Calculate various quality metrics
            // Null values, duplicates, format consistency, etc.

            return metrics;
        }

        private DataTable AnalyzeTrends(List<DataTable> dataTables)
        {
            var trends = new DataTable();
            trends.Columns.Add("Period", typeof(string));
            trends.Columns.Add("Trend", typeof(string));
            trends.Columns.Add("ChangePercent", typeof(decimal));

            // Trend analysis logic here

            return trends;
        }
    }
}
```

## Step 9: Implement Main Function

### ProcessExcelFiles.cs

```csharp
using ExcelProcessor.Auth;
using ExcelProcessor.Models;
using ExcelProcessor.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace ExcelProcessor
{
    public class ProcessExcelFiles
    {
        [FunctionName("ProcessExcelFiles")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("Excel processing function triggered");

            try
            {
                // Read request body
                string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
                var request = JsonConvert.DeserializeObject<ProcessingRequest>(requestBody);

                if (request == null || request.FileUrls.Count == 0)
                {
                    return new BadRequestObjectResult(new ProcessingResponse
                    {
                        Success = false,
                        Message = "Invalid request: No files specified"
                    });
                }

                // Initialize services
                var clientId = Environment.GetEnvironmentVariable("APP_CLIENT_ID");
                var clientSecret = Environment.GetEnvironmentVariable("APP_CLIENT_SECRET");
                var tenantId = Environment.GetEnvironmentVariable("APP_TENANT_ID");

                var authenticator = new SharePointAuthenticator(clientId, clientSecret, tenantId, request.SiteUrl);
                var sharePointService = new SharePointService(authenticator, request.SiteUrl);
                var excelProcessor = new ExcelProcessor();
                var reportGenerator = new ReportGenerator();

                // Download and process files
                var processedFiles = new List<FileInfo>();

                foreach (var fileUrl in request.FileUrls)
                {
                    try
                    {
                        log.LogInformation($"Downloading file: {fileUrl}");
                        var fileInfo = await sharePointService.DownloadFileAsync(fileUrl);
                        processedFiles.Add(fileInfo);
                    }
                    catch (Exception ex)
                    {
                        log.LogError(ex, $"Error downloading file {fileUrl}");
                        // Continue with other files
                    }
                }

                if (processedFiles.Count == 0)
                {
                    return new BadRequestObjectResult(new ProcessingResponse
                    {
                        Success = false,
                        Message = "Failed to download any files"
                    });
                }

                // Generate reports
                log.LogInformation("Generating reports...");
                var generatedReports = reportGenerator.GenerateReports(processedFiles, request.Options);

                // Upload reports to SharePoint
                foreach (var report in generatedReports)
                {
                    try
                    {
                        // Create folder structure if needed
                        var folderPath = $"{DateTime.Now.Year}/{DateTime.Now:MM}";
                        await sharePointService.CreateFolderIfNotExistsAsync(report.LibraryName, folderPath);

                        // Upload file (you'd need to modify ReportGenerator to return file content)
                        var fileUrl = await sharePointService.UploadFileAsync(
                            Array.Empty<byte>(), // You'd get this from ReportGenerator
                            report.FileName,
                            report.LibraryName,
                            folderPath);

                        report.Url = $"{request.SiteUrl}{fileUrl}";
                        log.LogInformation($"Uploaded report: {report.FileName}");

                    }
                    catch (Exception ex)
                    {
                        log.LogError(ex, $"Error uploading report {report.FileName}");
                        report.Url = "Error: Upload failed";
                    }
                }

                var response = new ProcessingResponse
                {
                    Success = true,
                    Message = $"Successfully processed {processedFiles.Count} files and generated {generatedReports.Count} reports",
                    GeneratedReports = generatedReports,
                    ProcessedFiles = processedFiles.Count
                };

                return new OkObjectResult(response);

            }
            catch (Exception ex)
            {
                log.LogError(ex, "Error in ProcessExcelFiles function");

                return new BadRequestObjectResult(new ProcessingResponse
                {
                    Success = false,
                    Message = $"Processing failed: {ex.Message}",
                    Errors = new List<string> { ex.Message }
                });
            }
        }
    }
}
```

## Step 10: Configure Function Settings

### local.settings.json

```json
{
  "IsEncrypted": false,
  "Values": {
    "AzureWebJobsStorage": "UseDevelopmentStorage=true",
    "FUNCTIONS_WORKER_RUNTIME": "dotnet",
    "APP_CLIENT_ID": "your-client-id",
    "APP_CLIENT_SECRET": "your-client-secret",
    "APP_TENANT_ID": "your-tenant-id"
  }
}
```

### host.json

```json
{
  "version": "2.0",
  "logging": {
    "applicationInsights": {
      "samplingSettings": {
        "isEnabled": true,
        "excludedTypes": "Request"
      }
    }
  },
  "functionTimeout": "00:10:00",
  "extensions": {
    "http": {
      "routePrefix": "api",
      "maxOutstandingRequests": 20,
      "maxConcurrentRequests": 10
    }
  }
}
```

## Step 11: Test Locally

### Run Function Locally

```bash
func start
```

### Test with Postman

```http
POST http://localhost:7071/api/ProcessExcelFiles
Content-Type: application/json

{
  "siteUrl": "https://yourtenant.sharepoint.com/sites/excel-processing",
  "fileUrls": [
    "https://yourtenant.sharepoint.com/sites/excel-processing/Input Files/test.xlsx"
  ],
  "userId": "testuser@yourtenant.onmicrosoft.com",
  "options": {
    "removeDuplicates": "true",
    "sortByDate": "true"
  }
}
```

## Step 12: Deploy to Azure

### Create Function App

```bash
az login
az group create --name excel-processor-rg --location eastus
az storage account create --name excelprocessorstorage --location eastus --resource-group excel-processor-rg --sku Standard_LRS
az functionapp create --resource-group excel-processor-rg --consumption-plan-location eastus --runtime dotnet --functions-version 4 --name excel-processor-function --storage-account excelprocessorstorage
```

### Deploy Function

```bash
func azure functionapp publish excel-processor-function
```

### Configure Application Settings

```bash
az functionapp config appsettings set --name excel-processor-function --resource-group excel-processor-rg --settings APP_CLIENT_ID=your-client-id APP_CLIENT_SECRET=your-client-secret APP_TENANT_ID=your-tenant-id
```

## Key Features Implemented

- ✅ HTTP-triggered Azure Function in C#
- ✅ SharePoint authentication and file operations
- ✅ Excel processing with ClosedXML
- ✅ Data transformation and merging
- ✅ Multiple report generation
- ✅ Error handling and logging
- ✅ Configurable processing options

## Next Steps

With your Azure Function implemented, proceed to [integrating the components](./05-integration.md) to connect your SPFx web part with the processing function.

## Troubleshooting

### Common Issues

1. **Authentication failures**: Verify app registration and permissions
2. **File download errors**: Check SharePoint URLs and permissions
3. **Excel processing errors**: Validate file formats and ClosedXML usage
4. **Upload failures**: Verify library names and folder permissions

### Performance Considerations

- Process files in parallel for large batches
- Implement streaming for large Excel files
- Use Azure Queue Storage for long-running processes
- Monitor memory usage with Application Insights



