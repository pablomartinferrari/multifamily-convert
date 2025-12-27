import * as React from 'react';
import {
  PrimaryButton,
  MessageBar,
  MessageBarType,
  ProgressIndicator,
  Label,
  TextField,
  Dropdown,
  IDropdownOption,
  Link,
  Stack,
  Separator,
  Icon
} from '@fluentui/react';
import { HttpClient, HttpClientResponse, SPHttpClient } from '@microsoft/sp-http';
import type { IMultifamilyProcessorProps } from './IMultifamilyProcessorProps';

export interface IGeneratedReport {
  reportType: string;
  fileName: string;
  url: string;
}

export interface IMultifamilyProcessorState {
  isProcessing: boolean;
  message: string;
  messageType: MessageBarType;
  showMessage: boolean;
  jobNumber: string;
  fileType: string;
  generatedReports: IGeneratedReport[];
  processingProgress: {
    stage: 'searching' | 'processing' | 'generating' | 'complete';
    message: string;
    percentComplete: number;
    foundCount?: number;
  } | undefined;
}

const fileTypeOptions: IDropdownOption[] = [
  { key: 'Units', text: 'Units' },
  { key: 'Common Areas', text: 'Common Areas' }
];

export default class MultifamilyProcessor extends React.Component<IMultifamilyProcessorProps, IMultifamilyProcessorState> {

  constructor(props: IMultifamilyProcessorProps) {
    super(props);

    this.state = {
      isProcessing: false,
      message: '',
      messageType: MessageBarType.info,
      showMessage: false,
      jobNumber: '',
      fileType: 'Units',
      generatedReports: [],
      processingProgress: undefined
    };
  }

  public render(): React.ReactElement<IMultifamilyProcessorProps> {
    return (
      <section style={{ padding: '25px', border: '1px solid #ddd', borderRadius: '8px', backgroundColor: 'white', boxShadow: '0 2px 4px rgba(0,0,0,0.1)' }}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} style={{ marginBottom: '20px' }}>
          <Icon iconName="ExcelDocument" style={{ fontSize: '24px', color: '#217346' }} />
          <h2 style={{ margin: 0 }}>XRF Multifamily Job Processor</h2>
        </Stack>
        
        {this.state.showMessage && (
          <MessageBar
            messageBarType={this.state.messageType}
            isMultiline={true}
            onDismiss={() => this.setState({ showMessage: false })}
            style={{ marginBottom: '20px' }}
          >
            {this.state.message}
          </MessageBar>
        )}

        <Stack tokens={{ childrenGap: 25 }}>
          <p style={{ color: '#666', margin: 0 }}>
            Enter a Job Number to find files in the <strong>{this.props.listId || 'XRF Files'}</strong> library and generate inspection reports.
          </p>

          <Stack horizontal tokens={{ childrenGap: 20 }}>
            <Stack.Item grow={1}>
              <TextField
                label="Job Number (Column Value)"
                required
                value={this.state.jobNumber}
                onChange={(e, val) => this.setState({ jobNumber: val || '' })}
                placeholder="Enter Job Number to search for..."
                disabled={this.state.isProcessing}
              />
            </Stack.Item>
            <Stack.Item grow={1}>
              <Dropdown
                label="Data Type"
                required
                selectedKey={this.state.fileType}
                options={fileTypeOptions}
                onChange={(e, opt) => this.setState({ fileType: opt?.key as string })}
                disabled={this.state.isProcessing}
              />
            </Stack.Item>
          </Stack>

          {this.renderProgress()}

          <PrimaryButton
            text={this.state.isProcessing ? "Processing..." : "Search & Process Files"}
            onClick={() => { void this.processJob(); }}
            disabled={this.state.isProcessing || !this.state.jobNumber}
            iconProps={{ iconName: 'Search' }}
            style={{ height: '45px', fontSize: '16px' }}
          />
        </Stack>

        {this.state.generatedReports.length > 0 && (
          <div style={{ marginTop: '30px', padding: '20px', backgroundColor: '#f0f9ff', borderRadius: '4px', border: '1px solid #0078d4' }}>
            <h3 style={{ marginTop: 0, fontSize: '18px' }}>✅ Reports Generated</h3>
            <Stack tokens={{ childrenGap: 10 }}>
              {this.state.generatedReports.map((report: IGeneratedReport, index: number) => (
                <div key={index} style={{ padding: '10px', backgroundColor: 'white', borderRadius: '4px', borderLeft: '4px solid #0078d4' }}>
                  <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 15 }}>
                    <Icon iconName="ExcelDocument" style={{ color: '#217346', fontSize: '20px' }} />
                    <Stack grow={1}>
                      <span style={{ fontWeight: 'bold' }}>{report.reportType}</span>
                      <span style={{ fontSize: '12px', color: '#666' }}>{report.fileName}</span>
                    </Stack>
                    <Link href={report.url} target="_blank" data-interception="off">
                      Open Report
                    </Link>
                  </Stack>
                </div>
              ))}
            </Stack>
          </div>
        )}
      </section>
    );
  }

  private renderProgress = (): JSX.Element | null => {
    if (!this.state.processingProgress) return null;

    return (
      <div style={{ margin: '10px 0', padding: '15px', backgroundColor: '#f3f2f1', borderRadius: '4px' }}>
        <ProgressIndicator
          label={this.state.processingProgress.message}
          percentComplete={this.state.processingProgress.percentComplete / 100}
        />
      </div>
    );
  }

  private processJob = async (): Promise<void> => {
    try {
      this.setState({
        isProcessing: true,
        showMessage: false,
        generatedReports: [],
        processingProgress: { stage: 'searching', message: `Searching for files with Job Number: ${this.state.jobNumber}...`, percentComplete: 20 }
      });

      // 1. Find files in the library matching the job number and type
      const listName = this.props.listId || 'Documents';
      const fileUrls = await this._findFilesByJobNumber(listName, this.state.jobNumber, this.state.fileType);

      if (fileUrls.length === 0) {
        throw new Error(`No files found with Job Number '${this.state.jobNumber}' and Type '${this.state.fileType}' in library '${listName}'.`);
      }

      this.setState({
        processingProgress: { stage: 'processing', message: `Found ${fileUrls.length} files. Sending to Azure...`, percentComplete: 50 }
      });

      // 2. Call Azure Function
      const payload = {
        siteUrl: this.props.siteUrl,
        fileUrls: fileUrls,
        jobNumber: this.state.jobNumber,
        fileType: this.state.fileType,
        userId: this.props.context.pageContext.user.email
      };

      const response: HttpClientResponse = await this.props.context.httpClient.post(
        this.props.azureFunctionUrl,
        HttpClient.configurations.v1,
        {
          body: JSON.stringify(payload),
          headers: { 'Content-Type': 'application/json' }
        }
      );

      if (!response.ok) throw new Error(`Azure Function error: ${response.statusText}`);

      const result = await response.json();
      if (result.success) {
        this.setState({
          generatedReports: result.generatedReports,
          processingProgress: { stage: 'complete', message: 'Complete!', percentComplete: 100 }
        });
        this.showMessage(`Successfully processed ${fileUrls.length} files.`, MessageBarType.success);
      } else {
        throw new Error(result.message);
      }

    } catch (error) {
      console.error(error);
      this.showMessage(error instanceof Error ? error.message : 'An unexpected error occurred', MessageBarType.error);
    } finally {
      this.setState({ isProcessing: false, processingProgress: undefined });
    }
  }

  private _findFilesByJobNumber = async (listName: string, jobNumber: string, fileType: string): Promise<string[]> => {
    const siteUrl = this.props.siteUrl;
    
    // Using 'InspectionType' (Capitalized) as specified by the user
    const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=FileRef,FileLeafRef,JobNumber,InspectionType&$filter=JobNumber eq '${jobNumber}' and InspectionType eq '${fileType}'`;
    
    console.log(`[XRF Processor] Target Library: ${listName}`);
    console.log(`[XRF Processor] Filtering by Job: ${jobNumber} AND InspectionType: ${fileType}`);
    console.log(`[XRF Processor] Full Request URL: ${endpoint}`);

    try {
      const response = await this.props.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
      
      if (!response.ok) {
        const errorData = await response.json();
        const errorMessage = errorData?.error?.message?.value || response.statusText;
        
        if (response.status === 400 && (errorMessage.indexOf('InspectionType') !== -1 || errorMessage.indexOf('inspectiontype') !== -1)) {
          throw new Error(`The column 'InspectionType' was not found. Please ensure the Internal Name matches exactly (case-sensitive).`);
        }
        
        throw new Error(`SharePoint Error: ${errorMessage}`);
      }
      
      const data = await response.json();
      const items = data.value || [];
      
      return items
        .filter((item: { FileLeafRef: string }) => {
          const name = item.FileLeafRef.toLowerCase();
          return name.endsWith('.xlsx') || name.endsWith('.xls') || name.endsWith('.csv');
        })
        .map((item: { FileRef: string }) => item.FileRef);
    } catch (e) {
      throw e;
    }
  }

  private showMessage = (message: string, type: MessageBarType): void => {
    this.setState({ message, messageType: type, showMessage: true });
  }
}
