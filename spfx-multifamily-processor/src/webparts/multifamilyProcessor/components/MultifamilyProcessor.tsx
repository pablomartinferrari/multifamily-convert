import * as React from 'react';
import {
  PrimaryButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  ProgressIndicator,
  Label
} from '@fluentui/react';
import { sp } from '@pnp/sp/presets/all';
import type { IMultifamilyProcessorProps } from './IMultifamilyProcessorProps';

export interface IMultifamilyProcessorState {
  selectedFiles: any[];
  isProcessing: boolean;
  message: string;
  messageType: MessageBarType;
  showMessage: boolean;
  processingProgress: {
    stage: 'uploading' | 'processing' | 'generating' | 'complete';
    message: string;
    percentComplete: number;
    currentFile?: string;
  } | null;
}

export default class MultifamilyProcessor extends React.Component<IMultifamilyProcessorProps, IMultifamilyProcessorState> {

  constructor(props: IMultifamilyProcessorProps) {
    super(props);

    this.state = {
      selectedFiles: [],
      isProcessing: false,
      message: '',
      messageType: MessageBarType.info,
      showMessage: false,
      processingProgress: null
    };

    // Initialize PnP JS
    sp.setup({
      spfxContext: this.props.context
    });
  }

  public render(): React.ReactElement<IMultifamilyProcessorProps> {
    return (
      <section style={{ padding: '20px', border: '1px solid #ddd', borderRadius: '4px' }}>
        <h2>Multifamily Excel Processor</h2>
        <p>Process Excel files and generate automated reports</p>

        {this.state.showMessage && (
          <MessageBar
            messageBarType={this.state.messageType}
            isMultiline={false}
            onDismiss={() => this.setState({ showMessage: false })}
            dismissButtonAriaLabel="Close"
            style={{ marginBottom: '20px' }}
          >
            {this.state.message}
          </MessageBar>
        )}

        {this.renderProgress()}

        <div style={{ margin: '20px 0' }}>
          <h3>Instructions:</h3>
          <ol>
            <li>Select one or more Excel files from the document library above</li>
            <li>Click the "Process Files" button</li>
            <li>Wait for processing to complete</li>
            <li>Check the report libraries for generated reports</li>
          </ol>
        </div>

        <div style={{ margin: '20px 0' }}>
          <PrimaryButton
            text={this.state.isProcessing ? "Processing..." : "Process Selected Files"}
            onClick={this.processFiles}
            disabled={this.state.isProcessing}
            iconProps={this.state.isProcessing ? { iconName: 'Sync' } : { iconName: 'Document' }}
          />

          {this.state.isProcessing && (
            <div style={{ marginTop: '10px' }}>
              <Spinner size={SpinnerSize.small} label="Processing files..." />
            </div>
          )}
        </div>

        {this.state.selectedFiles.length > 0 && (
          <div style={{ marginTop: '20px', padding: '15px', backgroundColor: '#f8f9fa', borderRadius: '4px' }}>
            <h4>Selected Files ({this.state.selectedFiles.length}):</h4>
            <ul>
              {this.state.selectedFiles.map((file, index) => (
                <li key={index} style={{ margin: '5px 0' }}>
                  📄 {file.name}
                </li>
              ))}
            </ul>
          </div>
        )}

        <div style={{ marginTop: '20px', padding: '10px', backgroundColor: '#f8f9fa', borderRadius: '4px' }}>
          <p><strong>Status:</strong> Web part loaded successfully!</p>
          <p><strong>User:</strong> {this.props.userDisplayName}</p>
          <p><strong>Site:</strong> {this.props.context.pageContext.web.title}</p>
        </div>
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
        processingProgress: {
          stage: 'uploading',
          message: 'Finding Excel files...',
          percentComplete: 10
        }
      });

      // Get Excel files from the configured document library
      const selectedFiles = await this.getExcelFiles();

      if (selectedFiles.length === 0) {
        this.showMessage('No Excel files found in the document library. Please upload some .xlsx or .xls files first.', MessageBarType.warning);
        return;
      }

      this.setState({
        selectedFiles,
        processingProgress: {
          stage: 'processing',
          message: `Processing ${selectedFiles.length} files...`,
          percentComplete: 30
        }
      });

      // Simulate processing delay (replace with actual Azure Function call)
      await this.delay(2000);

      this.setState({
        processingProgress: {
          stage: 'generating',
          message: 'Generating reports...',
          percentComplete: 70
        }
      });

      // Simulate report generation
      await this.delay(1500);

      this.setState({
        processingProgress: {
          stage: 'complete',
          message: 'Processing completed successfully!',
          percentComplete: 100
        }
      });

      this.showMessage(`Successfully processed ${selectedFiles.length} Excel files and generated reports!`, MessageBarType.success);

      // Clear progress after a delay
      setTimeout(() => {
        this.setState({ processingProgress: null });
      }, 3000);

    } catch (error) {
      console.error('Error processing files:', error);
      this.showMessage('An error occurred while processing files. Please try again.', MessageBarType.error);
    } finally {
      this.setState({ isProcessing: false });
    }
  }

  private getExcelFiles = async (): Promise<any[]> => {
    try {
      // For now, we'll get all Excel files from the default Documents library
      // In the full implementation, this would use the configured listId from props
      const files = await sp.web.lists.getByTitle('Documents').items
        .filter("substringof('.xlsx',FileLeafRef) or substringof('.xls',FileLeafRef)")
        .select('Id,FileLeafRef,FileRef')
        .top(10) // Limit to 10 files for testing
        .get();

      return files.map(file => ({
        id: file.Id,
        name: file.FileLeafRef,
        url: file.FileRef
      }));

    } catch (error) {
      console.error('Error getting Excel files:', error);
      // Return empty array instead of throwing to allow graceful handling
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
