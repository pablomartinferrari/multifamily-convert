import * as React from 'react';
import {
  PrimaryButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  ProgressIndicator,
  Label,
  TextField,
  Dropdown,
  IDropdownOption,
  Link,
  Stack,
  Separator
} from '@fluentui/react';
import { HttpClient, HttpClientResponse, IHttpClientOptions, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import type { IMultifamilyProcessorProps } from './IMultifamilyProcessorProps';

export interface IGeneratedReport {
  reportType: string;
  fileName: string;
  url: string;
}

export interface IMultifamilyProcessorState {
  selectedFiles: any[];
  isProcessing: boolean;
  message: string;
  messageType: MessageBarType;
  showMessage: boolean;
  jobNumber: string;
  fileType: string;
  generatedReports: IGeneratedReport[];
  processingProgress: {
    stage: 'uploading' | 'processing' | 'generating' | 'complete';
    message: string;
    percentComplete: number;
    currentFile?: string;
  } | null;
}

const fileTypeOptions: IDropdownOption[] = [
  { key: 'Units', text: 'Units' },
  { key: 'Common Areas', text: 'Common Areas' }
];

export default class MultifamilyProcessor extends React.Component<IMultifamilyProcessorProps, IMultifamilyProcessorState> {

  constructor(props: IMultifamilyProcessorProps) {
    super(props);

    this.state = {
      selectedFiles: [],
      isProcessing: false,
      message: '',
      messageType: MessageBarType.info,
      showMessage: false,
      jobNumber: '',
      fileType: 'Units',
      generatedReports: [],
      processingProgress: null
    };
  }

  public componentDidMount(): void {
    // Start polling for selection changes
    this._pollSelection();
  }

  private _pollSelection = (): void => {
    // Check for selected items in the host list
    if (this.props.context.sdks.microsoftTeams) return; // Not applicable in Teams

    // This is a common way to get selected items in SPFx web parts when they are on a list page
    // or when the context provides it.
    const selectedItems = (this.props.context as any).listView?.selectedItems;
    if (selectedItems && selectedItems.length > 0) {
      const files = selectedItems.map((item: any) => ({
        id: item.id,
        name: item.fileName || item.name,
        url: item.fileRef || item.url
      }));

      // Only update state if selection actually changed to avoid infinite re-renders
      if (JSON.stringify(files) !== JSON.stringify(this.state.selectedFiles)) {
        this.setState({ selectedFiles: files });
      }
    } else {
      if (this.state.selectedFiles.length > 0) {
        this.setState({ selectedFiles: [] });
      }
    }

    // Continue polling
    setTimeout(this._pollSelection, 1000);
  }

  public render(): React.ReactElement<IMultifamilyProcessorProps> {
    return (
      <section style={{ padding: '20px', border: '1px solid #ddd', borderRadius: '4px' }}>
        <h2>Multifamily Excel Processor</h2>
        <p>Process Excel files and generate automated reports based on 40-shot / 2.5% rules.</p>

        {this.state.showMessage && (
          <MessageBar
            messageBarType={this.state.messageType}
            isMultiline={true}
            onDismiss={() => this.setState({ showMessage: false })}
            dismissButtonAriaLabel="Close"
            style={{ marginBottom: '20px' }}
          >
            {this.state.message}
          </MessageBar>
        )}

        {this.renderProgress()}

        <Stack tokens={{ childrenGap: 15 }} style={{ marginBottom: '20px' }}>
          <TextField
            label="Job Number"
            required
            value={this.state.jobNumber}
            onChange={(e, val) => this.setState({ jobNumber: val || '' })}
            placeholder="e.g. 2025-XRF-101"
            disabled={this.state.isProcessing}
          />

          <Dropdown
            label="File Type"
            required
            selectedKey={this.state.fileType}
            options={fileTypeOptions}
            onChange={(e, opt) => this.setState({ fileType: opt?.key as string })}
            disabled={this.state.isProcessing}
          />
        </Stack>

        <div style={{ margin: '20px 0' }}>
          <PrimaryButton
            text={this.state.isProcessing ? "Processing..." : "Process Selected Files"}
            onClick={this.processFiles}
            disabled={this.state.isProcessing || this.state.selectedFiles.length === 0}
            iconProps={this.state.isProcessing ? { iconName: 'Sync' } : { iconName: 'Document' }}
          />

          {this.state.isProcessing && (
            <div style={{ marginTop: '10px' }}>
              <Spinner size={SpinnerSize.small} label="Calling Azure Function..." />
            </div>
          )}
        </div>

        {this.state.generatedReports.length > 0 && (
          <div style={{ marginTop: '20px', padding: '15px', backgroundColor: '#f0f9ff', borderRadius: '4px', border: '1px solid #0078d4' }}>
            <h4 style={{ marginTop: 0 }}>✅ Generated Reports:</h4>
            <ul style={{ listStyleType: 'none', padding: 0 }}>
              {this.state.generatedReports.map((report, index) => (
                <li key={index} style={{ marginBottom: '10px' }}>
                  <strong>{report.reportType}:</strong> <Link href={report.url} target="_blank">{report.fileName}</Link>
                </li>
              ))}
            </ul>
          </div>
        )}

        <Separator>Selected Files</Separator>

        {this.state.selectedFiles.length > 0 ? (
          <div style={{ marginTop: '10px', padding: '15px', backgroundColor: '#f8f9fa', borderRadius: '4px' }}>
            <ul style={{ paddingLeft: '20px' }}>
              {this.state.selectedFiles.map((file, index) => (
                <li key={index} style={{ margin: '5px 0' }}>
                  📄 {file.name}
                </li>
              ))}
            </ul>
          </div>
        ) : (
          <p style={{ fontStyle: 'italic', color: '#666' }}>No files selected from the library above.</p>
        )}
      </section>
    );
  }

  private renderProgress = (): JSX.Element | null => {
    if (!this.state.processingProgress) return null;

    return (
      <div style={{ margin: '20px 0', padding: '15px', backgroundColor: '#e8f4fd', borderRadius: '4px', border: '1px solid #0078d4' }}>
        <Label>Processing Progress</Label>
        <ProgressIndicator
          label={this.state.processingProgress.message}
          description={this.state.processingProgress.currentFile}
          percentComplete={this.state.processingProgress.percentComplete / 100}
          styles={{ progressBar: { backgroundColor: '#0078d4' } }}
        />
      </div>
    );
  }

  private processFiles = async (): Promise<void> => {
    try {
      this.setState({
        isProcessing: true,
        showMessage: false,
        generatedReports: [],
        processingProgress: {
          stage: 'uploading',
          message: 'Preparing data...',
          percentComplete: 10
        }
      });

      const { selectedFiles } = this.state;

      if (selectedFiles.length === 0) {
        this.showMessage('No files selected. Please select files from the list first.', MessageBarType.warning);
        this.setState({ isProcessing: false, processingProgress: null });
        return;
      }

      this.setState({
        processingProgress: {
          stage: 'processing',
          message: `Sending ${selectedFiles.length} files to Azure Function...`,
          percentComplete: 30
        }
      });

      // 2. Prepare the payload for the Azure Function
      const payload = {
        siteUrl: this.props.siteUrl,
        fileUrls: selectedFiles.map(f => f.url),
        jobNumber: this.state.jobNumber || 'AUTO-JOB-' + new Date().getTime(),
        fileType: this.state.fileType,
        userId: this.props.context.pageContext.user.email
      };

      // 3. Call the Azure Function
      const requestOptions: IHttpClientOptions = {
        body: JSON.stringify(payload),
        headers: {
          'Content-Type': 'application/json'
        }
      };

      const functionUrl = this.props.azureFunctionUrl;
      if (!functionUrl) {
        throw new Error('Azure Function URL is not configured in web part properties.');
      }

      const response: HttpClientResponse = await this.props.context.httpClient.post(
        functionUrl,
        HttpClient.configurations.v1,
        requestOptions
      );

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Azure Function failed: ${response.statusText} (${response.status}). ${errorText}`);
      }

      const result = await response.json();

      if (result.success) {
        this.setState({
          generatedReports: result.generatedReports,
          processingProgress: {
            stage: 'complete',
            message: 'Processing completed successfully!',
            percentComplete: 100
          }
        });
        this.showMessage('Successfully processed files and generated reports.', MessageBarType.success);
      } else {
        throw new Error(result.message || 'Unknown error occurred during processing.');
      }

      // Clear progress after a delay
      setTimeout(() => {
        this.setState({ processingProgress: null });
      }, 5000);

    } catch (error) {
      console.error('Error processing files:', error);
      const errorMessage = error instanceof Error ? error.message : 'An unknown error occurred';
      this.showMessage(`Error: ${errorMessage}`, MessageBarType.error);
      this.setState({ processingProgress: null });
    } finally {
      this.setState({ isProcessing: false });
    }
  }

  private getExcelFiles = async (): Promise<any[]> => {
    try {
      const listTitle = this.props.listId || 'Documents';
      const endpoint = `${this.props.siteUrl}/_api/web/lists/getbytitle('${listTitle}')/items?$select=Id,FileLeafRef,FileRef&$filter=substringof('.xlsx',FileLeafRef) or substringof('.xls',FileLeafRef)&$top=20`;

      const response: SPHttpClientResponse = await this.props.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Error fetching files: ${response.statusText}`);
      }

      const data = await response.json();
      return (data.value || []).map((file: any) => ({
        id: file.Id,
        name: file.FileLeafRef || file.Name,
        url: file.FileRef || file.ServerRelativeUrl
      }));

    } catch (error) {
      console.error('Error getting Excel files:', error);
      return [];
    }
  }

  private delay = (ms: number): Promise<void> => {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  private showMessage = (message: string, type: MessageBarType): void => {
    this.setState({
      message,
      messageType: type,
      showMessage: true
    });
  }
}
