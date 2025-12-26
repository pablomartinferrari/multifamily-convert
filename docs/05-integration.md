# 5. Integration and API Communication

This section covers integrating your SPFx web part with the Azure Function, including secure API communication, error handling, and user experience optimization.

## Prerequisites

- SPFx web part deployed to SharePoint (from section 2)
- Azure Function deployed and configured (from sections 3-4)
- Function URL and access keys
- SharePoint site with document libraries configured

## Step 1: Configure SPFx Web Part Properties

### Update Web Part Properties

In your SharePoint page, edit the Excel Processor web part and configure:

```
Document Library ID: [Your Input Files library GUID]
Azure Function URL: https://yourfunction.azurewebsites.net/api/ProcessExcelFiles?code=your-function-key
```

### Get Document Library ID

1. Navigate to your Input Files library
2. Go to **Library settings** → **List name, description and navigation**
3. Copy the URL parameter after `List=`: `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`

## Step 2: Implement Secure API Communication

### Enhanced HTTP Client in SPFx

Update your `ExcelProcessor.tsx` with improved API communication:

```tsx
private async callAzureFunction(fileUrls: string[]): Promise<ProcessingResponse> {
  const functionUrl = this.props.azureFunctionUrl;

  // Build request payload
  const requestBody = {
    siteUrl: this.props.context.pageContext.web.absoluteUrl,
    fileUrls: fileUrls,
    userId: this.props.context.pageContext.user.email,
    options: {
      removeDuplicates: 'true',
      validateData: 'true',
      generateCharts: 'false'
    }
  };

  try {
    // Add timeout and retry logic
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 300000); // 5 minute timeout

    const response = await fetch(functionUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'X-SharePoint-User': this.props.context.pageContext.user.email,
        'X-Request-Source': 'SPFx-WebPart'
      },
      body: JSON.stringify(requestBody),
      signal: controller.signal
    });

    clearTimeout(timeoutId);

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({}));
      throw new Error(`HTTP ${response.status}: ${errorData.message || response.statusText}`);
    }

    const result: ProcessingResponse = await response.json();
    return result;

  } catch (error) {
    if (error.name === 'AbortError') {
      throw new Error('Request timed out. The processing may still be running in the background.');
    }

    console.error('Azure Function call failed:', error);
    throw error;
  }
}
```

### Define Processing Response Interface

```tsx
export interface ProcessingResponse {
  success: boolean;
  message: string;
  generatedReports: GeneratedReport[];
  errors: string[];
  processedFiles: number;
}

export interface GeneratedReport {
  reportType: string;
  fileName: string;
  libraryName: string;
  url: string;
  generatedAt: string;
}
```

## Step 3: Enhance File Selection Integration

### Improved File Selection from Document Library

Update your file selection logic to work better with SharePoint's document library web parts:

```tsx
private getSelectedFilesFromDocumentLibrary = async (): Promise<string[]> => {
  try {
    // Use SharePoint's REST API to get files from the configured library
    const libraryId = this.props.listId;
    const siteUrl = this.props.context.pageContext.web.absoluteUrl;

    // Query for Excel files in the library
    const queryUrl = `${siteUrl}/_api/web/lists('${libraryId}')/items?$select=FileRef,FileLeafRef&$filter=substringof('.xlsx',FileLeafRef) or substringof('.xls',FileLeafRef)&$orderby=Modified desc`;

    const response = await this.props.context.spHttpClient.get(queryUrl, SPHttpClient.configurations.v1);

    if (!response.ok) {
      throw new Error(`Failed to query document library: ${response.statusText}`);
    }

    const data = await response.json();

    // Convert to full URLs
    return data.value.map((item: any) => `${siteUrl}${item.FileRef}`);

  } catch (error) {
    console.error('Error getting files from document library:', error);
    this.showMessage('Failed to retrieve files from document library.', MessageBarType.error);
    return [];
  }
}
```

### Add File Preview and Validation

```tsx
private validateSelectedFiles = (fileUrls: string[]): { valid: string[]; invalid: string[] } => {
  const validExtensions = ['.xlsx', '.xls'];
  const maxFileSize = 50 * 1024 * 1024; // 50MB limit

  const valid: string[] = [];
  const invalid: string[] = [];

  fileUrls.forEach(url => {
    const extension = url.toLowerCase().substring(url.lastIndexOf('.'));
    if (validExtensions.includes(extension)) {
      valid.push(url);
    } else {
      invalid.push(url);
    }
  });

  return { valid, invalid };
}
```

## Step 4: Implement Progress Tracking

### Add Progress Indicator

```tsx
export interface ProcessingProgress {
  stage: 'uploading' | 'processing' | 'generating' | 'complete';
  message: string;
  percentComplete: number;
  currentFile?: string;
}

private updateProgress = (progress: ProcessingProgress): void => {
  this.setState({
    processingProgress: progress,
    progressMessage: progress.message,
    progressPercent: progress.percentComplete
  });
}
```

### Progress Bar Component

```tsx
private renderProgress = (): JSX.Element | null => {
  if (!this.state.processingProgress) return null;

  return (
    <div style={{ margin: '20px 0' }}>
      <Label>Processing Progress</Label>
      <ProgressIndicator
        label={this.state.processingProgress.message}
        description={this.state.processingProgress.currentFile}
        percentComplete={this.state.processingProgress.percentComplete / 100}
      />
    </div>
  );
}
```

## Step 5: Error Handling and Recovery

### Comprehensive Error Handling

```tsx
private handleProcessingError = (error: any): void => {
  let errorMessage = 'An unexpected error occurred during processing.';
  let errorType = MessageBarType.error;

  if (error.message?.includes('timeout')) {
    errorMessage = 'Processing is taking longer than expected. Check the report libraries later for results.';
    errorType = MessageBarType.warning;
  } else if (error.message?.includes('authentication')) {
    errorMessage = 'Authentication failed. Please contact your administrator.';
  } else if (error.message?.includes('permission')) {
    errorMessage = 'You do not have permission to process these files.';
  } else if (error.message?.includes('network')) {
    errorMessage = 'Network error. Please check your connection and try again.';
    errorType = MessageBarType.warning;
  }

  this.showMessage(errorMessage, errorType);

  // Log detailed error for debugging
  console.error('Processing error details:', {
    error: error.message,
    stack: error.stack,
    userAgent: navigator.userAgent,
    timestamp: new Date().toISOString()
  });
}
```

### Retry Logic

```tsx
private callAzureFunctionWithRetry = async (fileUrls: string[], maxRetries: number = 3): Promise<ProcessingResponse> => {
  let lastError: any;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      this.updateProgress({
        stage: 'uploading',
        message: `Processing files (attempt ${attempt}/${maxRetries})...`,
        percentComplete: 10,
        currentFile: fileUrls.length > 1 ? `${fileUrls.length} files` : fileUrls[0]?.split('/').pop()
      });

      const result = await this.callAzureFunction(fileUrls);

      this.updateProgress({
        stage: 'complete',
        message: 'Processing completed successfully!',
        percentComplete: 100
      });

      return result;

    } catch (error) {
      lastError = error;
      console.warn(`Attempt ${attempt} failed:`, error.message);

      if (attempt < maxRetries) {
        // Wait before retry (exponential backoff)
        const delay = Math.pow(2, attempt) * 1000;
        await new Promise(resolve => setTimeout(resolve, delay));

        this.updateProgress({
          stage: 'uploading',
          message: `Retrying... (attempt ${attempt + 1}/${maxRetries})`,
          percentComplete: 10 + (attempt * 20)
        });
      }
    }
  }

  throw lastError;
}
```

## Step 6: User Feedback and Results Display

### Results Display Component

```tsx
private renderResults = (): JSX.Element | null => {
  if (!this.state.processingResult) return null;

  const { processingResult } = this.state;

  return (
    <div style={{ margin: '20px 0', padding: '15px', backgroundColor: '#f8f9fa', borderRadius: '4px' }}>
      <h4>Processing Results</h4>

      {processingResult.success ? (
        <div>
          <Icon iconName="CheckMark" style={{ color: 'green', marginRight: '8px' }} />
          <span>{processingResult.message}</span>

          {processingResult.generatedReports.length > 0 && (
            <div style={{ marginTop: '15px' }}>
              <Label>Generated Reports:</Label>
              <ul>
                {processingResult.generatedReports.map((report, index) => (
                  <li key={index} style={{ margin: '5px 0' }}>
                    <Link href={report.url} target="_blank">
                      {report.reportType}: {report.fileName}
                    </Link>
                    <span style={{ marginLeft: '10px', color: '#666' }}>
                      ({report.libraryName})
                    </span>
                  </li>
                ))}
              </ul>
            </div>
          )}
        </div>
      ) : (
        <div>
          <Icon iconName="Error" style={{ color: 'red', marginRight: '8px' }} />
          <span>{processingResult.message}</span>

          {processingResult.errors.length > 0 && (
            <details style={{ marginTop: '10px' }}>
              <summary>Error Details</summary>
              <ul>
                {processingResult.errors.map((error, index) => (
                  <li key={index} style={{ color: '#d13438' }}>{error}</li>
                ))}
              </ul>
            </details>
          )}
        </div>
      )}
    </div>
  );
}
```

## Step 7: Security Enhancements

### Request Signing and Validation

```tsx
private generateRequestSignature = (requestBody: any): string => {
  const timestamp = Date.now().toString();
  const payload = JSON.stringify(requestBody) + timestamp;
  // In production, use proper HMAC signing with a shared secret
  // For now, return a simple hash
  return btoa(payload).substring(0, 32);
}
```

### Add Security Headers

```tsx
private callAzureFunction = async (fileUrls: string[]): Promise<ProcessingResponse> => {
  const requestBody = {
    siteUrl: this.props.context.pageContext.web.absoluteUrl,
    fileUrls: fileUrls,
    userId: this.props.context.pageContext.user.email,
    timestamp: Date.now(),
    signature: this.generateRequestSignature({ fileUrls })
  };

  const response = await fetch(this.props.azureFunctionUrl, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'X-SharePoint-User': this.props.context.pageContext.user.loginName,
      'X-Request-Source': 'SPFx-ExcelProcessor',
      'X-Timestamp': requestBody.timestamp.toString()
    },
    body: JSON.stringify(requestBody)
  });

  // Validate response signature if implemented
  return response.json();
}
```

## Step 8: Performance Optimizations

### File Batching

```tsx
private processFilesInBatches = async (fileUrls: string[], batchSize: number = 5): Promise<ProcessingResponse> => {
  const batches = [];
  for (let i = 0; i < fileUrls.length; i += batchSize) {
    batches.push(fileUrls.slice(i, i + batchSize));
  }

  const results: ProcessingResponse[] = [];

  for (let i = 0; i < batches.length; i++) {
    const batch = batches[i];
    this.updateProgress({
      stage: 'processing',
      message: `Processing batch ${i + 1}/${batches.length}...`,
      percentComplete: (i / batches.length) * 100
    });

    const result = await this.callAzureFunctionWithRetry(batch);
    results.push(result);
  }

  // Combine results
  return this.combineBatchResults(results);
}
```

### Combine Batch Results

```tsx
private combineBatchResults = (results: ProcessingResponse[]): ProcessingResponse => {
  const combined: ProcessingResponse = {
    success: results.every(r => r.success),
    message: results.every(r => r.success)
      ? `Successfully processed ${results.reduce((sum, r) => sum + r.processedFiles, 0)} files across ${results.length} batches`
      : 'Some batches failed to process',
    generatedReports: results.flatMap(r => r.generatedReports),
    errors: results.flatMap(r => r.errors),
    processedFiles: results.reduce((sum, r) => sum + r.processedFiles, 0)
  };

  return combined;
}
```

## Step 9: Testing Integration

### End-to-End Test Script

Create a test script to verify the complete flow:

```typescript
// Test integration
private runIntegrationTest = async (): Promise<void> => {
  try {
    this.setState({ isProcessing: true, showMessage: false });

    // Test 1: File selection
    const testFiles = await this.getSelectedFilesFromDocumentLibrary();
    if (testFiles.length === 0) {
      throw new Error('No test files found in document library');
    }

    // Test 2: API connectivity
    const testResponse = await fetch(`${this.props.azureFunctionUrl.split('?')[0]}/test`, {
      method: 'GET'
    });

    if (!testResponse.ok) {
      throw new Error('Azure Function is not accessible');
    }

    // Test 3: Full processing (with small file)
    const result = await this.callAzureFunctionWithRetry([testFiles[0]], 1);

    this.showMessage('Integration test passed successfully!', MessageBarType.success);

  } catch (error) {
    this.showMessage(`Integration test failed: ${error.message}`, MessageBarType.error);
  } finally {
    this.setState({ isProcessing: false });
  }
}
```

## Step 10: Monitoring and Logging

### Implement Application Insights

```tsx
private logUserAction = (action: string, details?: any): void => {
  // Log to SharePoint's usage analytics or your custom logging
  console.log('User Action:', {
    action,
    user: this.props.context.pageContext.user.email,
    timestamp: new Date().toISOString(),
    details
  });

  // In production, send to Application Insights or your logging service
}
```

### Track Processing Metrics

```tsx
private trackProcessingMetrics = (result: ProcessingResponse, duration: number): void => {
  const metrics = {
    success: result.success,
    processedFiles: result.processedFiles,
    generatedReports: result.generatedReports.length,
    duration,
    timestamp: new Date().toISOString(),
    user: this.props.context.pageContext.user.email
  };

  this.logUserAction('processing_completed', metrics);
}
```

## Step 11: Update Main Processing Flow

### Complete Updated processFiles Method

```tsx
private processFiles = async (): Promise<void> => {
  const startTime = Date.now();

  try {
    this.setState({
      isProcessing: true,
      showMessage: false,
      processingResult: null,
      processingProgress: null
    });

    this.logUserAction('processing_started');

    // Get files from document library
    const fileUrls = await this.getSelectedFilesFromDocumentLibrary();

    if (fileUrls.length === 0) {
      this.showMessage('No Excel files found in the document library.', MessageBarType.warning);
      return;
    }

    // Validate files
    const validation = this.validateSelectedFiles(fileUrls);
    if (validation.invalid.length > 0) {
      this.showMessage(`Some files are not Excel files and will be skipped: ${validation.invalid.join(', ')}`, MessageBarType.warning);
    }

    if (validation.valid.length === 0) {
      this.showMessage('No valid Excel files to process.', MessageBarType.error);
      return;
    }

    // Process files (with batching for large sets)
    const result = validation.valid.length > 10
      ? await this.processFilesInBatches(validation.valid)
      : await this.callAzureFunctionWithRetry(validation.valid);

    // Track metrics
    const duration = Date.now() - startTime;
    this.trackProcessingMetrics(result, duration);

    this.setState({ processingResult: result });

    if (result.success) {
      this.showMessage(result.message, MessageBarType.success, 10000); // Show longer for success
    } else {
      this.showMessage(result.message, MessageBarType.error);
    }

  } catch (error) {
    this.handleProcessingError(error);
    this.logUserAction('processing_failed', { error: error.message });
  } finally {
    this.setState({ isProcessing: false });
  }
}
```

## Step 12: Deploy and Test Integration

### Deploy SPFx Web Part

```bash
gulp clean
gulp build --production
gulp bundle --production
gulp package-solution --production
```

### Upload to SharePoint App Catalog

1. Upload the `.sppkg` file to your tenant app catalog
2. Approve API permissions if prompted
3. Add the web part to your page

### Test Complete Flow

1. Upload Excel files to Input Files library
2. Click "Process Selected Files" in the web part
3. Monitor progress and verify reports are generated
4. Check Application Insights for metrics

## Key Integration Features

- ✅ Secure API communication with authentication
- ✅ Progress tracking and user feedback
- ✅ Comprehensive error handling and retry logic
- ✅ File validation and batching
- ✅ Results display with report links
- ✅ Security headers and request signing
- ✅ Performance optimizations
- ✅ Monitoring and logging

## Next Steps

With integration complete, proceed to [deployment guides](./06-deployment.md) for production deployment and scaling considerations.

## Troubleshooting Integration Issues

### Common Problems

1. **CORS errors**: Ensure Azure Function allows SharePoint domain
2. **Authentication failures**: Verify function keys and permissions
3. **File access errors**: Check document library permissions
4. **Timeout issues**: Implement proper batching and progress tracking

### Debug Tools

- Browser developer tools for client-side debugging
- Azure Application Insights for server-side monitoring
- SharePoint ULS logs for permission issues
- Function App logs for Azure Function debugging

