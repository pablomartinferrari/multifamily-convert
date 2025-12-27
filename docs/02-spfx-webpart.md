# 2. SPFx Web Part with React

This section covers creating a SharePoint Framework (SPFx) web part using React that provides a button to process selected Excel files from the document library.

## Prerequisites

- Node.js 14+ and npm installed
- SharePoint Framework Yeoman generator: `npm install -g @microsoft/generator-sharepoint`
- Visual Studio Code with SPFx extensions
- Access to SharePoint site for testing

## Step 1: Set Up SPFx Development Environment

### Install SPFx Generator

```bash
npm install -g @microsoft/generator-sharepoint@latest
```

### Verify Installation

```bash
yo --version
gulp --version
```

## Step 2: Scaffold New SPFx Solution

### Create Project Directory

```bash
mkdir spfx-multifamily-processor
cd spfx-multifamily-processor
```

### Generate SPFx Web Part

```bash
yo @microsoft/sharepoint
```

Answer the prompts:

```
? What is your solution name? spfx-multifamily-processor
? Which type of client-side component to create? WebPart
? What is your Web part name? MultifamilyProcessor
? Which template would you like to use? React
? Where do you want to place the files? Use the current folder
? Do you want to allow the tenant admin to deploy the solution to all sites? Y
? Will the components in the solution require permissions to access web APIs? Y
```

### Install Dependencies

```bash
npm install
```

## Step 3: Project Structure Overview

After scaffolding, your project structure should look like:

```
spfx-multifamily-processor/
├── config/
│   ├── package-solution.json
│   ├── serve.json
│   └── write-manifests.json
├── src/
│   ├── webparts/
│   │   └── MultifamilyProcessor/
│   │       ├── components/
│   │       │   ├── MultifamilyProcessor.tsx     # Main React component
│   │       │   └── IMultifamilyProcessorProps.ts
│   │       ├── loc/
│   │       │   └── en-us.js
│   │       ├── MultifamilyProcessorWebPart.ts   # Web part class
│   │       └── MultifamilyProcessorWebPart.manifest.json
│   └── index.ts
├── package.json
├── tsconfig.json
└── gulpfile.js
```

## Step 4: Implement the React Component

### Update IMultifamilyProcessorProps.ts

```typescript
export interface IMultifamilyProcessorProps {
  description: string;
  context: any;
  siteUrl: string;
  listId: string;
  azureFunctionUrl: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
```

### Create MultifamilyProcessor.tsx Component

```tsx
import * as React from 'react';
import { IMultifamilyProcessorProps } from './IMultifamilyProcessorProps';
import { PrimaryButton, MessageBar, MessageBarType, Spinner, SpinnerSize } from '@fluentui/react';
import { sp } from '@pnp/sp/presets/all';

export interface IMultifamilyProcessorState {
  selectedFiles: any[];
  isProcessing: boolean;
  message: string;
  messageType: MessageBarType;
  showMessage: boolean;
}

export default class MultifamilyProcessor extends React.Component<IMultifamilyProcessorProps, IMultifamilyProcessorState> {

  constructor(props: IMultifamilyProcessorProps) {
    super(props);

    this.state = {
      selectedFiles: [],
      isProcessing: false,
      message: '',
      messageType: MessageBarType.info,
      showMessage: false
    };

    // Initialize PnP JS
    sp.setup({
      spfxContext: this.props.context
    });
  }

  public render(): React.ReactElement<IMultifamilyProcessorProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${isDarkTheme ? 'ms-bgColor-neutralDark' : ''} ms-fontColor-neutralPrimary`}>
        <div style={{ padding: '20px', border: '1px solid #ddd', borderRadius: '4px' }}>
          <h2>Excel File Processor</h2>
          <p>{description}</p>

          {this.state.showMessage && (
            <MessageBar
              messageBarType={this.state.messageType}
              isMultiline={false}
              onDismiss={() => this.setState({ showMessage: false })}
              dismissButtonAriaLabel="Close"
            >
              {this.state.message}
            </MessageBar>
          )}

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
            <div style={{ marginTop: '20px' }}>
              <h4>Selected Files ({this.state.selectedFiles.length}):</h4>
              <ul>
                {this.state.selectedFiles.map((file, index) => (
                  <li key={index}>{file.name}</li>
                ))}
              </ul>
            </div>
          )}
        </div>
      </section>
    );
  }

  private processFiles = async (): Promise<void> => {
    try {
      this.setState({
        isProcessing: true,
        showMessage: false
      });

      // Get selected files from document library
      const selectedFiles = await this.getSelectedFiles();

      if (selectedFiles.length === 0) {
        this.showMessage('Please select at least one Excel file from the document library.', MessageBarType.warning);
        return;
      }

      this.setState({ selectedFiles });

      // Call Azure Function
      const result = await this.callAzureFunction(selectedFiles);

      if (result.success) {
        this.showMessage(`Successfully processed ${selectedFiles.length} files. Reports generated.`, MessageBarType.success);
      } else {
        this.showMessage(`Processing failed: ${result.error}`, MessageBarType.error);
      }

    } catch (error) {
      console.error('Error processing files:', error);
      this.showMessage('An error occurred while processing files.', MessageBarType.error);
    } finally {
      this.setState({ isProcessing: false });
    }
  }

  private getSelectedFiles = async (): Promise<any[]> => {
    try {
      // This is a simplified approach - in a real implementation,
      // you'd need to integrate with the document library web part
      // or use SharePoint's selection API

      // For now, we'll get all Excel files from the configured library
      const files = await sp.web.lists.getById(this.props.listId).items
        .filter("substringof('.xlsx',FileLeafRef) or substringof('.xls',FileLeafRef)")
        .select('Id,FileLeafRef,FileRef')
        .get();

      return files.map(file => ({
        id: file.Id,
        name: file.FileLeafRef,
        url: file.FileRef
      }));

    } catch (error) {
      console.error('Error getting files:', error);
      return [];
    }
  }

  private callAzureFunction = async (files: any[]): Promise<{ success: boolean; error?: string }> => {
    try {
      const response = await fetch(this.props.azureFunctionUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          siteUrl: this.props.siteUrl,
          fileUrls: files.map(f => f.url),
          userId: this.props.userDisplayName
        })
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const result = await response.json();
      return result;

    } catch (error) {
      console.error('Error calling Azure Function:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  private showMessage = (message: string, type: MessageBarType): void => {
    this.setState({
      message,
      messageType: type,
      showMessage: true
    });
  }
}
```

## Step 5: Update the Web Part Class

### MultifamilyProcessorWebPart.ts

```typescript
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'MultifamilyProcessorWebPartStrings';
import MultifamilyProcessor from './components/MultifamilyProcessor';
import { IMultifamilyProcessorProps } from './components/IMultifamilyProcessorProps';

export interface IMultifamilyProcessorWebPartProps {
  description: string;
  listId: string;
  azureFunctionUrl: string;
}

export default class MultifamilyProcessorWebPart extends BaseClientSideWebPart<IMultifamilyProcessorWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IMultifamilyProcessorProps> = React.createElement(
      MultifamilyProcessor,
      {
        description: this.properties.description,
        context: this.context,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        listId: this.properties.listId,
        azureFunctionUrl: this.properties.azureFunctionUrl,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private get _isDarkTheme(): boolean {
    return this.context.sdks.microsoftTeams ? this.context.sdks.microsoftTeams.theme === 'dark' : false;
  }

  private get _environmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.pageContext.legacyPageContext.isSiteAdmin ? strings.AppLocalEnvironmentTeams : strings.AppLocalEnvironment;
    }

    return this.context.pageContext.legacyPageContext.isSiteAdmin ? strings.AppLocalEnvironmentSharePoint : strings.AppLocalEnvironment;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('listId', {
                  label: 'Document Library ID',
                  description: 'GUID of the Input Files document library'
                }),
                PropertyPaneTextField('azureFunctionUrl', {
                  label: 'Azure Function URL',
                  description: 'URL of the Azure Function endpoint'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
```

## Step 6: Update Package Configuration

### package-solution.json

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/package-solution.schema.json",
  "solution": {
    "name": "multifamily-processor-client-side-solution",
    "id": "7b8f4c2e-1a9d-4f3e-8b7c-5d2e9f1a3b8c",
    "version": "1.0.0.0",
    "includeClientSideAssets": true,
    "skipFeatureDeployment": true,
    "isDomainIsolated": false,
    "webApiPermissionRequests": [
      {
        "resource": "Microsoft Graph",
        "scope": "Sites.ReadWrite.All"
      },
      {
        "resource": "SharePoint",
        "scope": "Sites.ReadWrite.All"
      }
    ]
  },
  "paths": {
    "zippedPackage": "solution/multifamily-processor.sppkg"
  }
}
```

## Step 7: Add Required Dependencies

```bash
npm install @pnp/sp @pnp/graph @fluentui/react
```

## Step 8: Build and Test Locally

### Build the Solution

```bash
gulp build
```

### Start Local Development Server

```bash
gulp serve
```

### Test in SharePoint Workbench

1. Navigate to your SharePoint site
2. Add `?debug=true&noredir=true&debugManifestsFile=https://localhost:4321/temp/manifests.js` to the URL
3. Add the web part to a page
4. Configure the properties:
   - **Document Library ID**: Get this from your Input Files library URL
   - **Azure Function URL**: Will be configured later

## Step 9: Enhanced File Selection

For better integration with SharePoint document libraries, you can enhance the file selection by using SharePoint's selection API. Here's an improved version:

### Enhanced MultifamilyProcessor.tsx (File Selection)

```tsx
// Add this method to get selected files from SharePoint's selection context
private getSelectedFilesFromSharePoint = async (): Promise<any[]> => {
  try {
    // Use SharePoint's built-in selection API
    const selectedItems = await this.props.context.spHttpClient.get(
      `${this.props.siteUrl}/_api/web/lists('${this.props.listId}')/items?$filter=ID eq ${this.getSelectedItemIds().join(' or ID eq ')}`,
      SPHttpClient.configurations.v1
    );

    const response = await selectedItems.json();
    return response.value.map((item: any) => ({
      id: item.Id,
      name: item.FileLeafRef,
      url: item.FileRef
    }));

  } catch (error) {
    console.error('Error getting selected files:', error);
    return [];
  }
}

// Helper method to get selected item IDs (this would need to be implemented
// based on how you integrate with the document library web part)
private getSelectedItemIds(): number[] {
  // This is a placeholder - in a real implementation, you'd need to
  // communicate with the document library web part or use SharePoint's
  // selection API to get the selected items
  return [1, 2, 3]; // Placeholder IDs
}
```

## Step 10: Error Handling and User Feedback

### Enhanced Error Handling

```tsx
private showMessage = (message: string, type: MessageBarType, duration: number = 5000): void => {
  this.setState({
    message,
    messageType: type,
    showMessage: true
  });

  // Auto-hide success messages after duration
  if (type === MessageBarType.success) {
    setTimeout(() => {
      this.setState({ showMessage: false });
    }, duration);
  }
}

private validateConfiguration = (): boolean => {
  if (!this.props.listId) {
    this.showMessage('Document Library ID is not configured.', MessageBarType.error);
    return false;
  }

  if (!this.props.azureFunctionUrl) {
    this.showMessage('Azure Function URL is not configured.', MessageBarType.error);
    return false;
  }

  return true;
}
```

## Step 11: Package for Deployment

### Build Production Version

```bash
gulp clean
gulp build --production
gulp bundle --production
gulp package-solution --production
```

## Folder Structure Summary

```
spfx-multifamily-processor/
├── src/
│   └── webparts/
│       └── MultifamilyProcessor/
│           ├── components/
│           │   ├── MultifamilyProcessor.tsx       # Main React component
│           │   └── IMultifamilyProcessorProps.ts  # Props interface
│           └── MultifamilyProcessorWebPart.ts     # Web part class
├── config/
│   └── package-solution.json                # Package config
└── dist/                                   # Built solution
```

## Key Features Implemented

- ✅ React-based SPFx web part
- ✅ Button to trigger processing
- ✅ File selection from document library
- ✅ HTTP calls to Azure Function
- ✅ Loading states and user feedback
- ✅ Error handling
- ✅ Configurable properties
- ✅ PnP JS integration for SharePoint operations

## Next Steps

With your SPFx web part created, proceed to [configuring Azure authentication](./03-azure-auth.md) for secure communication between your web part and Azure Function.

## Troubleshooting

### Common Issues

1. **Web part not loading**: Check console for errors, ensure gulp serve is running
2. **PnP JS errors**: Ensure proper initialization with `sp.setup()`
3. **Permission errors**: Verify web API permissions in package-solution.json
4. **File selection not working**: Implement proper integration with document library selection API

### Development Tips

- Use browser developer tools to debug React components
- Test in SharePoint Workbench before deploying to production
- Use `gulp serve --nobrowser` for headless development
- Enable source maps for better debugging: `gulp build --sourceMap`



